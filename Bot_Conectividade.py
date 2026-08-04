"""
Bot Conectividade Social
========================
Automação do envio do arquivo .SFP no site da Conectividade Social V2
(https://conectividadesocialv2.caixa.gov.br/sicns/).

Fluxo (Etapa 5 do processo de recálculo de FGTS):
  1. Acessa o site e a Caixa Postal (login por certificado digital A1 —
     feito pelo NAVEGADOR/Windows, nunca automatizado por este script).
  2. Vai para "Nova Mensagem", serviço "Envio de arquivo SEFIP".
  3. Preenche Nome da Mensagem, Estado e Base de Arrecadação.
  4. Anexa o arquivo .SFP.
  5. Sem --confirmar-envio, PARA antes de clicar "Enviar" e MANTÉM o
     navegador aberto para revisão manual (o envio real é uma ação com
     efeito em sistema externo/oficial).

Mapeado via conectividade_spy.py (ver memory/sefip-automation.md):
  - mat-select[formControlName="funcionalidade"]  — Serviço
  - input  (placeholder="Nome da Mensagem")       — Nome da Mensagem
  - mat-select[name="estado"]                     — Estado
  - mat-select (3º)                               — Base de Arrecadação
  - input#customFile (formControlName="arquivo")  — anexo
  - button "Enviar" / button "Limpar"

Os mat-select são componentes Angular Material — não são <select> HTML
nativo. Interação: clicar para abrir o overlay de opções, depois clicar na
opção pelo texto.

IMPORTANTE sobre seletores: os IDs gerados pelo Angular Material
(mat-select-0, mat-input-0, ...) são numerados sequencialmente por sessão
de página e MUDAM se outros componentes forem criados antes (navegar por
outra tela, reabrir o formulário, etc.). Por isso cada campo é procurado
por uma LISTA de seletores, do mais estável (formControlName/name) para o
mais frágil (id fixo), usando o primeiro que existir.

Uso:
    python Bot_Conectividade.py --arquivo "C:\\...\\IEfG1TyfArf00002.SFP" \\
        --nome-mensagem "GFIP 404 11-2022" --estado "Rio de Janeiro" \\
        --base-arrecadacao "Volta Redonda / RJ"

Diagnóstico (recomendado enquanto o fluxo não estiver estável):
    python Bot_Conectividade.py ... --debug

    Com --debug, tudo que acontece no navegador é interceptado e gravado em
    `debug_conectividade/<timestamp>/`:
      - eventos.jsonl   — log estruturado (rede, console, erros, downloads,
                          diálogos, navegações, marcos do fluxo)
      - eventos.log     — o mesmo em texto legível
      - NN_<marco>.png  — screenshot de cada etapa
      - NN_<marco>.html — HTML da página em cada etapa
      - trace.zip       — trace do Playwright; abrir com:
                          python -m playwright show-trace trace.zip

Autor: Hugo L. Almeida
"""

import argparse
import base64
import json
import os
import re
import sys
import time
from datetime import datetime

from playwright.sync_api import sync_playwright

# Vigia da janela NATIVA do CriptoCNS (assinador da Caixa). Ela não existe no
# DOM — é um processo Win32 —, então Playwright nunca a encontra: o sintoma é
# o "Enviar" ficar desabilitado e o fluxo parar sem erro. Opcional: se o
# módulo não estiver presente, o bot segue funcionando (só não trata sozinho).
try:
    from vigia_criptocns import VigiaCriptoCNS
except ImportError:
    VigiaCriptoCNS = None

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

URL_INICIAL = "https://conectividadesocialv2.caixa.gov.br/sicns/"
USER_DATA_DIR = r"C:\Users\Canella e Santos\AppData\Local\Temp\claude_conectividade_profile"

# Timeout padrão (segundos) que o navegador fica aberto para revisão manual
# quando --confirmar-envio NÃO é informado.
TIMEOUT_REVISAO_PADRAO = 3600

# 🔴 Certificado digital A1 injetado DIRETO no contexto do navegador via
# client_certificates do Playwright — testado em 2026-08-04 depois de
# confirmar que --auto-select-certificate-for-urls (flag de linha de
# comando do Chromium) NÃO funciona neste site: a Conectividade Social não
# usa o handshake TLS mútuo padrão que essa flag intercepta (o popup nativo
# do Windows nunca chegava a aparecer nem a ser resolvido, e a página
# travava esperando "Nova Mensagem" pelos 90s inteiros).
# client_certificates funciona em outra camada (o próprio Chromium assina o
# desafio com a chave do .pfx, sem depender do CryptoAPI do Windows nem de
# qualquer popup) — mais confiável para automação sem supervisão.
# Arquivo fica FORA do repositório (pasta de rede dedicada a certificados,
# nunca dentro do projeto) — client_certificates lê o .pfx diretamente,
# sem precisar instalar o certificado no Windows.
CERTIFICADO_PFX_PATH = (
    r"Z:\003 - CERTIFICADO DIGITAL\102 - CANELLA E SANTOS"
    r"\CANELLA & SANTOS CONTABILIDADE LTDA   SENHA 12345678  V 23.02.2027.pfx"
)
CERTIFICADO_PFX_SENHA = "12345678"

# Subpasta criada DENTRO da pasta do .SFP para receber os artefatos do envio
# (retorno, comprovante PDF e protocolo.json). Mantém tudo da empresa junto:
#   results/404/MD9ghVk735n00008.SFP
#   results/404/protocolo/MD9ghVk735n00008_<retorno>
#   results/404/protocolo/protocolo.json
SUBPASTA_PROTOCOLO = "protocolo"


# ----------------------------------------------------------------------------
# Interceptores / instrumentação
# ----------------------------------------------------------------------------

class Rastreador:
    """
    Camada de observabilidade do fluxo no navegador.

    Registra (quando `ativo`):
      - requisições e respostas XHR/fetch do domínio da Caixa (método, URL,
        status, tempo, corpo JSON truncado) — é aqui que aparece o motivo
        real de uma validação recusada pelo servidor;
      - mensagens de console e erros de página não tratados (Angular quebrando
        silenciosamente é a causa mais comum de "o campo não apareceu");
      - diálogos nativos (alert/confirm) — que travariam o Playwright;
      - downloads iniciados;
      - navegações (framenavigated do frame principal);
      - "marcos" do fluxo, com screenshot + dump do HTML.

    Sempre grava no stdout; com `ativo=False` não escreve arquivos nem anexa
    listeners de rede (custo zero em produção).
    """

    def __init__(self, dir_debug=None, ativo=False):
        self.ativo = ativo
        self.dir_debug = dir_debug
        self.eventos = []
        self._n_marco = 0
        self._t0 = time.time()
        self._req_t = {}
        self._f_jsonl = None
        self._f_log = None

        if self.ativo:
            os.makedirs(self.dir_debug, exist_ok=True)
            self._f_jsonl = open(os.path.join(self.dir_debug, "eventos.jsonl"),
                                 "w", encoding="utf-8")
            self._f_log = open(os.path.join(self.dir_debug, "eventos.log"),
                               "w", encoding="utf-8")

    # -- log básico ---------------------------------------------------------

    def log(self, msg):
        """Mensagem para o usuário (sempre no stdout, e no arquivo se debug)."""
        print(msg, flush=True)
        self._escrever_log(msg)

    def _escrever_log(self, linha):
        if self._f_log:
            self._f_log.write(f"[{time.time() - self._t0:7.2f}s] {linha}\n")
            self._f_log.flush()

    def evento(self, tipo, **dados):
        ev = {"t": round(time.time() - self._t0, 3), "tipo": tipo, **dados}
        self.eventos.append(ev)
        if self._f_jsonl:
            self._f_jsonl.write(json.dumps(ev, ensure_ascii=False) + "\n")
            self._f_jsonl.flush()
            self._escrever_log(f"{tipo}: " + json.dumps(dados, ensure_ascii=False)[:400])
        return ev

    # -- listeners ----------------------------------------------------------

    def anexar(self, context, page):
        """Anexa os interceptores ao contexto/página."""
        # Console e erros valem sempre — são baratos e explicam falhas.
        page.on("console", self._on_console)
        page.on("pageerror", self._on_pageerror)
        # Diálogo nativo travaria o Playwright indefinidamente: registra e
        # aceita, em vez de deixar a automação pendurada.
        page.on("dialog", self._on_dialog)
        page.on("download", lambda d: self.evento(
            "download", sugerido=d.suggested_filename, url=d.url))
        page.on("framenavigated", self._on_navegacao)

        if not self.ativo:
            return

        # Rede só no modo debug: é o listener mais verboso.
        page.on("request", self._on_request)
        page.on("response", self._on_response)
        page.on("requestfailed", lambda r: self.evento(
            "req_falhou", metodo=r.method, url=r.url,
            erro=(r.failure or "")))

        try:
            context.tracing.start(screenshots=True, snapshots=True, sources=True)
            self._escrever_log("tracing iniciado")
        except Exception as e:
            self._escrever_log(f"tracing indisponível: {type(e).__name__}: {e}")

    def _on_console(self, msg):
        if msg.type in ("error", "warning"):
            self.evento("console", nivel=msg.type, texto=msg.text[:500])

    def _on_pageerror(self, err):
        self.evento("erro_pagina", texto=str(err)[:800])

    def _on_dialog(self, dialog):
        self.evento("dialogo", tipo=dialog.type, mensagem=dialog.message[:400])
        self.log(f"⚠️ Diálogo nativo do site: [{dialog.type}] {dialog.message}")
        try:
            dialog.accept()
        except Exception:
            pass

    def _on_navegacao(self, frame):
        try:
            if frame.parent_frame is None:
                self.evento("navegou", url=frame.url)
        except Exception:
            pass

    def _relevante(self, url):
        return "caixa.gov.br" in url

    def _on_request(self, request):
        if request.resource_type in ("xhr", "fetch") and self._relevante(request.url):
            self._req_t[request] = time.time()
            corpo = None
            try:
                corpo = (request.post_data or "")[:1000] or None
            except Exception:
                pass
            self.evento("req", metodo=request.method, url=request.url, corpo=corpo)

    def _on_response(self, response):
        req = response.request
        if req.resource_type not in ("xhr", "fetch") or not self._relevante(response.url):
            return
        ms = None
        if req in self._req_t:
            ms = int((time.time() - self._req_t.pop(req)) * 1000)
        corpo = None
        try:
            ctype = (response.header_value("content-type") or "")
            # Só lê o corpo de respostas textuais: ler binário (ZIP do
            # protocolo, PDF) desperdiça memória e polui o log.
            if "json" in ctype or "text" in ctype:
                corpo = response.text()[:2000]
        except Exception:
            pass
        self.evento("resp", status=response.status, metodo=req.method,
                    url=response.url, ms=ms, corpo=corpo)
        if response.status >= 400:
            self.log(f"⚠️ HTTP {response.status} em {req.method} {response.url}")

    # -- marcos -------------------------------------------------------------

    def marco(self, page, nome):
        """Ponto de checagem do fluxo: screenshot + HTML + evento."""
        self._n_marco += 1
        self.evento("marco", nome=nome, url=_url_segura(page))
        if not self.ativo:
            return
        base = os.path.join(self.dir_debug, f"{self._n_marco:02d}_{_slug(nome)}")
        try:
            page.screenshot(path=base + ".png", full_page=False)
        except Exception as e:
            self._escrever_log(f"screenshot falhou em '{nome}': {type(e).__name__}")
        try:
            with open(base + ".html", "w", encoding="utf-8") as f:
                f.write(page.content())
        except Exception as e:
            self._escrever_log(f"dump html falhou em '{nome}': {type(e).__name__}")

    def encerrar(self, context):
        if self.ativo:
            try:
                context.tracing.stop(path=os.path.join(self.dir_debug, "trace.zip"))
                self.log(f"🧭 Trace salvo. Abrir com:\n"
                         f"   python -m playwright show-trace "
                         f"\"{os.path.join(self.dir_debug, 'trace.zip')}\"")
            except Exception as e:
                self._escrever_log(f"tracing.stop falhou: {type(e).__name__}: {e}")
        for f in (self._f_jsonl, self._f_log):
            try:
                if f:
                    f.close()
            except Exception:
                pass


def _slug(texto):
    return re.sub(r"[^a-z0-9]+", "_", texto.lower()).strip("_")[:40]


# ----------------------------------------------------------------------------
# Token de máquina (JWS) — persistência entre execuções
# ----------------------------------------------------------------------------
#
# COMO O SITE GUARDA O TOKEN (mapeado no DOM real em 2026-07-27):
#   localStorage["<CNPJ 14 dígitos>"] = JWS RS256 (typ=JOSE, com jwk embutido)
#     claims: iss=http://cns.caixa.gov.br, aud=SEFIP, name=<CNPJ>, iat, exp
#     validade observada: ~7 meses.
#   sessionStorage["token"] = JWT da SESSÃO do certificado (~1h) — NÃO é o
#     token de máquina e não deve ser persistido.
#
# O token é POR PERFIL DE NAVEGADOR. Sem ele, o site exige "cadastro de
# máquina" e registra uma máquina NOVA a cada execução (o Manter Máquina da
# Canella tinha 10 máquinas ativas acumuladas por causa disso). Salvando o
# JWS em disco e reinjetando no perfil antes de usar, o cadastro deixa de ser
# necessário em qualquer modo (--cdp ou perfil isolado).

ARQUIVO_TOKEN_PADRAO = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "token_maquina.json")

_JS_LER_TOKENS = """() => {
    const out = {};
    for (let i = 0; i < localStorage.length; i++) {
        const k = localStorage.key(i);
        const v = localStorage.getItem(k) || '';
        if (/^\\d{14}$/.test(k) && v.split('.').length === 3) out[k] = v;
    }
    return out;
}"""


def _decodificar_jws(token):
    """Devolve (name, exp) do payload de um JWS. (None, None) se não der."""
    try:
        partes = token.split(".")
        if len(partes) != 3:
            return None, None
        p = partes[1]
        p += "=" * (-len(p) % 4)
        payload = json.loads(base64.urlsafe_b64decode(p))
        return payload.get("name"), payload.get("exp")
    except Exception:
        return None, None


def _carregar_tokens_do_disco(caminho, rastro):
    if not caminho or not os.path.exists(caminho):
        return {}
    try:
        with open(caminho, encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        rastro.evento("token_disco_ilegivel", erro=f"{type(e).__name__}: {e}")
        return {}


def _salvar_token_no_disco(caminho, cnpj, token, rastro):
    dados = _carregar_tokens_do_disco(caminho, rastro)
    _, exp = _decodificar_jws(token)
    if dados.get(cnpj, {}).get("token") == token:
        return False
    dados[cnpj] = {
        "token": token,
        "exp": exp,
        "exp_legivel": (datetime.fromtimestamp(exp).isoformat() if exp else None),
        "salvo_em": datetime.now().isoformat(timespec="seconds"),
    }
    try:
        with open(caminho, "w", encoding="utf-8") as f:
            json.dump(dados, f, ensure_ascii=False, indent=1)
        rastro.log(f"💾 Token de máquina de {cnpj} salvo em {caminho}")
        rastro.evento("token_salvo", cnpj=cnpj, exp=exp)
        return True
    except Exception as e:
        rastro.log(f"⚠️ Não consegui salvar o token: {type(e).__name__}: {e}")
        return False


def _gerenciar_token_maquina(page, caminho, rastro, injetar=True):
    """
    Sincroniza o token de máquina entre o perfil do navegador e o disco.
    Chamar LOGO APÓS o goto (precisa estar na origem certa para ver o
    localStorage do site).

    - Token no navegador  → salva/atualiza no disco.
    - Sem token, mas há um válido no disco → injeta e recarrega a página.
    - Sem token em lugar nenhum → avisa que o site vai pedir cadastro de
      máquina (e que isso criará mais uma entrada no Manter Máquina).

    Retorna: "presente" | "reinjetado" | "ausente" | "expirado".
    """
    try:
        no_navegador = page.evaluate(_JS_LER_TOKENS) or {}
    except Exception as e:
        rastro.evento("token_leitura_falhou", erro=f"{type(e).__name__}: {e}")
        return "ausente"

    if no_navegador:
        for cnpj, token in no_navegador.items():
            _, exp = _decodificar_jws(token)
            venc = datetime.fromtimestamp(exp).strftime("%d/%m/%Y") if exp else "?"
            rastro.log(f"🔑 Token de máquina presente no perfil ({cnpj}, "
                       f"válido até {venc}).")
            # NÃO promete pular o cadastro: em 27/07/2026 o modal de cadastro
            # apareceu MESMO com este token válido no localStorage. O que
            # dispensa o cadastro é outro controle do lado da Caixa (suspeita:
            # o endpoint common-api/.../serial/). Guardar o token continua útil,
            # mas não é garantia de fluxo curto.
            if caminho:
                _salvar_token_no_disco(caminho, cnpj, token, rastro)
        rastro.evento("token_status", status="presente",
                      cnpjs=list(no_navegador))
        return "presente"

    do_disco = _carregar_tokens_do_disco(caminho, rastro)
    if not do_disco:
        rastro.log("⚠️ Nenhum token de máquina no perfil nem em disco — o site "
                   "vai pedir CADASTRO DE MÁQUINA e registrar mais uma máquina. "
                   "Após esta execução o token será salvo para as próximas.")
        rastro.evento("token_status", status="ausente")
        return "ausente"

    if not injetar:
        rastro.log("ℹ️ Há token em disco, mas a injeção está desativada "
                   "(--nao-injetar-token).")
        rastro.evento("token_status", status="ausente_sem_injecao")
        return "ausente"

    agora = time.time()
    validos = {c: d for c, d in do_disco.items()
               if not d.get("exp") or d["exp"] > agora + 300}
    if not validos:
        rastro.log("⚠️ O token em disco está EXPIRADO — será necessário novo "
                   "cadastro de máquina.")
        rastro.evento("token_status", status="expirado")
        return "expirado"

    for cnpj, dados in validos.items():
        page.evaluate("([k, v]) => localStorage.setItem(k, v)",
                      [cnpj, dados["token"]])
        rastro.log(f"💉 Token de máquina de {cnpj} reinjetado no perfil "
                   f"(evita novo cadastro de máquina).")
    page.reload(wait_until="domcontentloaded")
    time.sleep(1.5)

    conferido = page.evaluate(_JS_LER_TOKENS) or {}
    ok = bool(conferido)
    rastro.evento("token_status", status="reinjetado", confirmado=ok,
                  cnpjs=list(conferido))
    if not ok:
        rastro.log("⚠️ A injeção não sobreviveu ao reload — o site pode ter "
                   "limpado o localStorage. Seguirá pelo cadastro de máquina.")
        return "ausente"
    return "reinjetado"


def _url_segura(page):
    try:
        return page.url
    except Exception:
        return None


# ----------------------------------------------------------------------------
# Helpers de localização (seletores resilientes)
# ----------------------------------------------------------------------------

def _localizar(page, seletores, descricao, rastro, timeout=10000):
    """
    Tenta cada seletor da lista, em ordem, e devolve o primeiro que estiver
    visível. Divide o timeout entre as tentativas para não multiplicar a
    espera total pelo número de candidatos.
    """
    por_tentativa = max(1500, int(timeout / max(1, len(seletores))))
    ultimo_erro = None
    for sel in seletores:
        try:
            loc = page.locator(sel).first
            loc.wait_for(state="visible", timeout=por_tentativa)
            rastro.evento("seletor_ok", campo=descricao, seletor=sel)
            return loc
        except Exception as e:
            ultimo_erro = e
            rastro.evento("seletor_falhou", campo=descricao, seletor=sel,
                          erro=type(e).__name__)
    rastro.marco(page, f"falha_{descricao}")
    raise RuntimeError(
        f"Campo '{descricao}' não encontrado. Seletores tentados: {seletores}. "
        f"Último erro: {type(ultimo_erro).__name__}"
    )


def _selecionar_mat_select(page, seletores, texto_opcao, descricao, rastro):
    """
    Angular Material mat-select não é <select> nativo — clica para abrir o
    overlay de opções e clica na opção correspondente.

    A opção é buscada primeiro por texto EXATO (evita, por exemplo, escolher
    "Baixar Token" quando se quer "Baixar", ou um estado cujo nome é prefixo
    de outro) e só cai para "contém" se o exato não existir.
    """
    rastro.log(f"📋 Selecionando {descricao} = {texto_opcao!r}...")
    campo = _localizar(page, seletores, descricao, rastro)
    campo.click()
    # O overlay do Material é renderizado fora do mat-select (no
    # cdk-overlay-container), por isso a busca é global na página. O filtro
    # `:visible` é essencial: overlays de selects já fechados podem continuar
    # no DOM (ocultos), e sem ele o script tentaria clicar na opção de OUTRO
    # campo.
    page.wait_for_selector("mat-option:visible", timeout=8000)
    visiveis = page.locator("mat-option:visible")

    exato = visiveis.filter(has_text=re.compile(rf"^\s*{re.escape(texto_opcao)}\s*$"))
    if exato.count() > 0:
        opcao = exato.first
    else:
        opcao = visiveis.filter(has_text=texto_opcao).first
        rastro.evento("opcao_aproximada", campo=descricao, texto=texto_opcao)

    try:
        opcao.wait_for(state="visible", timeout=5000)
    except Exception:
        disponiveis = []
        try:
            disponiveis = page.locator("mat-option:visible").all_inner_texts()[:40]
        except Exception:
            pass
        rastro.evento("opcao_nao_encontrada", campo=descricao,
                      procurado=texto_opcao, disponiveis=disponiveis)
        rastro.marco(page, f"opcoes_{descricao}")
        raise RuntimeError(
            f"Opção {texto_opcao!r} não encontrada em {descricao}. "
            f"Opções visíveis: {disponiveis}"
        )

    opcao.click()
    time.sleep(0.3)


def _conferir_anexo(page, arquivo, rastro):
    """
    Confirma que o Angular REGISTROU o anexo — não basta o set_input_files
    ter rodado sem erro.

    O rótulo visível ("Selecione os Arquivos") é estático (o "+ Adicionar" é
    CSS `::after`), então ele NÃO muda ao anexar: olhar para a tela engana.
    A prova real é o arquivo aparecer na lista `mat-list-item` com o ícone
    check.png, e o botão Enviar ficar habilitado.

    Retorna (anexo_ok, enviar_habilitado).
    """
    nome = os.path.basename(arquivo)
    item = page.locator("mat-list-item", has_text=nome)
    anexo_ok = False
    try:
        item.first.wait_for(state="visible", timeout=8000)
        anexo_ok = item.locator("img[src*='check']").count() > 0 or item.count() > 0
    except Exception:
        pass

    # O site pode ACEITAR o arquivo na lista (com ícone de check!) e mesmo
    # assim recusá-lo na validação de conteúdo, escrevendo a razão em vermelho
    # logo abaixo do nome. Confirmado em 2026-07-29:
    #   "Não é possível enviar o arquivo. A inscrição do responsável no
    #    arquivo deve ser igual á inscrição da empresa logado."
    # Sem ler isso, o bot só veria "Enviar desabilitado" e reportaria um
    # erro genérico — escondendo a causa real, que é de CADASTRO, não do bot.
    erro_arquivo = _ler_erro_do_anexo(page, item, rastro)

    # O Angular pode habilitar o Enviar de forma assíncrona (depois de validar
    # o arquivo), então não basta amostrar uma vez: espera até 8s.
    enviar_habilitado = False
    limite = time.time() + 8
    while time.time() < limite:
        try:
            botao = _botao(page, "Enviar")
            if botao is not None and botao.is_enabled(timeout=1000):
                enviar_habilitado = True
                break
        except Exception:
            pass
        time.sleep(0.5)

    rastro.evento("conferencia_anexo", arquivo=nome, anexo_ok=anexo_ok,
                  enviar_habilitado=enviar_habilitado, erro_arquivo=erro_arquivo)

    if erro_arquivo:
        rastro.log("❌ O SITE RECUSOU O ARQUIVO:")
        rastro.log(f"   {erro_arquivo}")
        for dica in _explicar_recusa(erro_arquivo):
            rastro.log(f"   → {dica}")
        rastro.marco(page, "arquivo_recusado")
    elif anexo_ok and enviar_habilitado:
        rastro.log(f"✅ Anexo confirmado na lista ({nome}) e botão Enviar habilitado.")
    else:
        rastro.log(f"⚠️ Conferência do anexo FALHOU — arquivo na lista: {anexo_ok}, "
                   f"Enviar habilitado: {enviar_habilitado}. O formulário pode "
                   f"não estar realmente pronto para envio.")
        rastro.marco(page, "anexo_nao_confirmado")
    return anexo_ok, enviar_habilitado, erro_arquivo


# Mensagens de recusa conhecidas do site → explicação acionável. A chave é um
# trecho estável (sem acento/caixa) da mensagem; o site escreve "á" onde
# deveria ser "à", então casamos por pedaços curtos e tolerantes.
RECUSAS_CONHECIDAS = [
    (("inscricao do responsavel", "inscricao da empresa"),
     ["O CNPJ do responsável DENTRO do .SFP não bate com o certificado logado.",
      "Não é erro do bot nem do arquivo em si: é cadastro.",
      "Ou o SEFIP gerou o arquivo com outro responsável, ou é preciso entrar "
      "em 'Acessar Empresa Outorgante' com a empresa correta.",
      "Confira o responsável no SEFIP (Cadastro de Responsável) antes de reenviar."]),
    (("arquivo ja foi enviado",),
     ["Este arquivo já consta como transmitido — confira em 'Itens Enviados' "
      "antes de reenviar, para não duplicar."]),
    (("layout", "versao"),
     ["O layout/versão do arquivo não é o aceito pela Caixa — regerar no SEFIP."]),
    (("competencia",),
     ["Competência do arquivo inconsistente com o serviço selecionado."]),
]


def _sem_acento(texto):
    import unicodedata
    return "".join(c for c in unicodedata.normalize("NFD", texto or "")
                   if unicodedata.category(c) != "Mn").lower()


def _explicar_recusa(mensagem):
    """Traduz a mensagem do site em próximos passos. Vazio se desconhecida."""
    alvo = _sem_acento(mensagem)
    for marcas, dicas in RECUSAS_CONHECIDAS:
        if all(m in alvo for m in marcas):
            return dicas
    return ["Recusa não catalogada — trate manualmente e me avise o texto "
            "para eu adicionar o tratamento."]


def _ler_erro_do_anexo(page, item, rastro):
    """
    Lê a mensagem de recusa que o site escreve dentro do item do anexo.

    Estratégia em camadas porque não sabemos a classe exata do elemento:
    procura texto em vermelho / classes de erro dentro do item e, como
    último recurso, qualquer linha do item que não seja o nome do arquivo.
    """
    seletores = [".erro", ".error", ".mat-error", ".text-danger",
                 "[class*='erro']", "[class*='error']", "small", "span"]
    try:
        alvo = item.first
    except Exception:
        return None

    for sel in seletores:
        try:
            loc = alvo.locator(sel)
            for i in range(min(loc.count(), 6)):
                txt = (loc.nth(i).inner_text() or "").strip()
                if _parece_erro(txt):
                    return " ".join(txt.split())
        except Exception:
            continue

    # Fallback: o texto inteiro do item, tirando o nome do arquivo.
    try:
        inteiro = " ".join((alvo.inner_text() or "").split())
        for linha in inteiro.split("  "):
            if _parece_erro(linha):
                return linha.strip()
        if _parece_erro(inteiro):
            return inteiro
    except Exception:
        pass
    return None


def _parece_erro(texto):
    if not texto or len(texto) < 15:
        return False
    alvo = _sem_acento(texto)
    marcas = ("nao e possivel", "nao foi possivel", "erro", "invalid",
              "deve ser igual", "nao permitido", "recusado", "falha")
    return any(m in alvo for m in marcas)


def _botao(page, texto, exato=True):
    """
    Localiza um botão pelo texto, de forma robusta ao markup do Angular
    Material.

    ATENÇÃO (confirmado no DOM real em 2026-07-27): os botões do site são
    `<button><span class="mat-button-wrapper">Enviar</span>...</button>`.
    O seletor `button:text-is("Enviar")` casa ZERO elementos nesse markup,
    porque o texto não é filho direto do <button>. Por isso usamos o NOME
    ACESSÍVEL (get_by_role), que é calculado a partir dos descendentes.

    Em modo exato NÃO há fallback para "contém": "Baixar" jamais pode acabar
    clicando em "Baixar Token".
    """
    # Case-insensitive de propósito: o site renderiza "SIM" em maiúsculas
    # (text-transform), e get_by_role com exact=True é case-SENSITIVE — foi
    # exatamente o que fez o clique em 'Sim' falhar em 2026-07-27. Com regex
    # de linha inteira, "Baixar" continua não casando "Baixar Token".
    padrao = (re.compile(rf"^\s*{re.escape(texto)}\s*$", re.I) if exato
              else re.compile(re.escape(texto), re.I))
    tentativas = [
        page.get_by_role("button", name=padrao),
        page.locator("button").filter(has_text=padrao),
    ]
    for loc in tentativas:
        try:
            if loc.count() > 0:
                return loc.first
        except Exception:
            continue
    return None


def _textos_dos_botoes(page, limite=25):
    """Lista os botões visíveis — usado nas mensagens de erro."""
    try:
        return [t.strip() for t in
                page.locator("button:visible").all_inner_texts()][:limite]
    except Exception:
        return []


def _clicar_botao(page, texto, rastro, timeout=8000, exato=True):
    """Clica um botão pelo texto (exato por padrão)."""
    loc = _botao(page, texto, exato=exato)
    if loc is None:
        disponiveis = _textos_dos_botoes(page)
        rastro.evento("botao_nao_encontrado", botao=texto, disponiveis=disponiveis)
        rastro.marco(page, f"sem_botao_{_slug(texto)}")
        raise RuntimeError(
            f"Botão {texto!r} não encontrado. Botões visíveis: {disponiveis}")
    loc.wait_for(state="visible", timeout=timeout)
    rastro.evento("clique", botao=texto, exato=exato)
    loc.click()


def _visivel_agora(page, seletor, timeout=8000):
    """
    Locator.is_visible() IGNORA o parâmetro timeout (retorna na hora), então
    um elemento que ainda vai renderizar seria dado como ausente. Aqui a
    espera é feita de verdade via wait_for.
    """
    try:
        page.locator(seletor).first.wait_for(state="visible", timeout=timeout)
        return True
    except Exception:
        return False


# ----------------------------------------------------------------------------
# Fluxo principal
# ----------------------------------------------------------------------------

# Seletores por campo, do mais estável para o mais frágil.
SEL_SERVICO = [
    'mat-select[formControlName="funcionalidade"]',
    'mat-select[name="servicos"]',
    'mat-select[placeholder="Selecione o serviço"]',
    "#mat-select-0",
]
SEL_NOME_MENSAGEM = [
    'input[placeholder="Nome da Mensagem"]',
    'input[formControlName="nomeMensagem"]',
    "#mat-input-0",
]
SEL_ESTADO = [
    'mat-select[name="estado"]',
    'mat-select[placeholder="Selecione o estado"]',
    "#mat-select-1",
]
# A Base de Arrecadação NÃO tem formControlName nem name no DOM real
# (confirmado no dump de 2026-07-27): os únicos atributos semânticos são
# placeholder e aria-label. Por isso eles vêm primeiro aqui, e o id só
# como último recurso.
SEL_BASE = [
    'mat-select[placeholder="Selecione a Base de Arrecadação"]',
    'mat-select[aria-label="Selecione a Base de Arrecadação"]',
    "#mat-select-2",
    "mat-select >> nth=2",
]
SEL_ARQUIVO = [
    "#customFile",
    'input[type="file"][formControlName="arquivo"]',
    'input[type="file"]',
]


def enviar_arquivo_sefip(arquivo: str, nome_mensagem: str, estado: str,
                         base_arrecadacao: str, confirmar_envio: bool = False,
                         dir_download: str = None, debug: bool = False,
                         dir_debug: str = None,
                         timeout_revisao: int = TIMEOUT_REVISAO_PADRAO,
                         cdp: str = None, arquivo_token: str = ARQUIVO_TOKEN_PADRAO,
                         injetar_token: bool = True, nome_maquina: str = None,
                         vigiar_criptocns: bool = True,
                         criptocns_nao_perguntar: bool = False):
    """
    Preenche o formulário "Nova Mensagem" (Envio de arquivo SEFIP) e anexa
    o arquivo.

    Com `confirmar_envio=False` (padrão), para antes de clicar Enviar e
    MANTÉM o navegador aberto até você fechá-lo (ou até `timeout_revisao`
    segundos). Isso é essencial: o processo do Chrome está atrelado ao
    processo Python — se a função retornasse aqui, o bloco
    `with sync_playwright()` fecharia o driver e o navegador cairia junto.

    Com `confirmar_envio=True`, conduz o fluxo pós-Enviar completo (desafio
    de token → cadastro de máquina → baixar token → protocolo de envio →
    baixar arquivo de retorno) e salva os downloads em `dir_download`.

    Retorna um dict: {"status", "protocolo", "token", "dir_debug", "erro"}.
    """
    if debug and not dir_debug:
        carimbo = datetime.now().strftime("%Y%m%d_%H%M%S")
        dir_debug = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                 "debug_conectividade", carimbo)

    rastro = Rastreador(dir_debug=dir_debug, ativo=debug)
    log = rastro.log
    resultado = {"status": "erro", "protocolo": None, "token": None,
                 "token_maquina": None,
                 "dir_debug": dir_debug if debug else None, "erro": None}

    # Auto-seleciona o certificado PJ CANELLA E SANTOS (CNPJ 06.310.711/0001-49,
    # thumbprint 9661049FC0CB41A96D2463A8A0179B805DD17D15) para o domínio da
    # Conectividade Social, evitando o popup nativo do Windows/Chrome a cada
    # sessão nova. Filtro por ISSUER (emissor) + CN do assunto.
    politica_certificado = (
        '{"pattern":"https://conectividadesocialv2.caixa.gov.br",'
        '"filter":{"ISSUER":{"CN":"AC SyngularID Multipla"},'
        '"SUBJECT":{"CN":"CANELLA E SANTOS CONTABILIDADE LTDA:06310711000149"}}}'
    )

    with sync_playwright() as p:
        if cdp:
            # Anexa a um Chrome JÁ ABERTO (iniciado com
            # --remote-debugging-port=9222). Vantagens: reaproveita a sessão
            # já autenticada por certificado, mantém as extensões do usuário
            # (inclusive a do Claude) e não disputa o lock do perfil.
            # Aqui NÃO se aplica --auto-select-certificate-for-urls: essa flag
            # só vale no start do navegador; se a sessão ainda não estiver
            # autenticada, o popup de certificado aparece e você escolhe.
            log(f"🔌 Conectando ao Chrome existente em {cdp} ...")
            navegador = p.chromium.connect_over_cdp(cdp)
            if not navegador.contexts:
                raise RuntimeError(
                    f"Conectado a {cdp}, mas o Chrome não expôs nenhum contexto.")
            context = navegador.contexts[0]
            # Abre uma ABA NOVA em vez de sequestrar a aba atual do usuário.
            page = context.new_page()
        else:
            navegador = None
            client_certificates = None
            if os.path.exists(CERTIFICADO_PFX_PATH):
                client_certificates = [{
                    "origin": "https://conectividadesocialv2.caixa.gov.br",
                    "pfxPath": CERTIFICADO_PFX_PATH,
                    "passphrase": CERTIFICADO_PFX_SENHA,
                }]
            else:
                log(f"⚠️ Certificado não encontrado em '{CERTIFICADO_PFX_PATH}' — "
                    "seguindo sem client_certificates (pode pedir seleção manual).")
            context = p.chromium.launch_persistent_context(
                USER_DATA_DIR,
                channel="chrome",
                headless=False,
                accept_downloads=True,
                viewport={"width": 1400, "height": 900},
                args=[f"--auto-select-certificate-for-urls=[{politica_certificado}]"],
                client_certificates=client_certificates,
            )
            page = context.pages[0] if context.pages else context.new_page()
        rastro.anexar(context, page)

        if debug:
            log(f"🔎 Modo debug: artefatos em {dir_debug}")

        # Precisa começar ANTES de qualquer navegação: o CriptoCNS pode pedir
        # autorização já na abertura da Caixa Postal, não só no Enviar.
        vigia = None
        if vigiar_criptocns and VigiaCriptoCNS is not None:
            vigia = VigiaCriptoCNS(
                marcar_nao_perguntar=criptocns_nao_perguntar, log=log).iniciar()
        elif vigiar_criptocns:
            log("ℹ️ vigia_criptocns.py não encontrado ao lado do script — se a "
                "janela do CriptoCNS aparecer, clique 'Aceitar' manualmente.")

        try:
            log(f"🌐 Acessando {URL_INICIAL} ...")
            page.goto(URL_INICIAL, wait_until="domcontentloaded")
            time.sleep(2)
            rastro.marco(page, "inicial")

            # Precisa vir DEPOIS do goto (localStorage é por origem) e ANTES do
            # envio, para que o cadastro de máquina nem chegue a ser pedido.
            resultado["token_maquina"] = _gerenciar_token_maquina(
                page, arquivo_token, rastro, injetar=injetar_token)

            # A página inicial (sicns/) tem os botões de serviço (Caixa Postal,
            # Manter Máquina, etc.) — não é a Caixa Postal em si. Clicar em
            # "Caixa Postal" primeiro; isso dispara o desafio de certificado
            # digital, se a sessão ainda não estiver autenticada.
            try:
                page.click("text=Caixa Postal", timeout=10000)
                time.sleep(1.5)
            except Exception:
                log("ℹ️ Botão 'Caixa Postal' não encontrado na tela inicial — "
                    "talvez já esteja em outra tela. Prosseguindo.")

            # Aguarda o menu "Nova Mensagem" (indica que passou pelo
            # certificado e chegou na Caixa Postal). Se não aparecer em 15s,
            # provavelmente está esperando o usuário escolher o certificado no
            # popup do Windows/Chrome — dá mais tempo para resolver.
            try:
                page.wait_for_selector("text=Nova Mensagem", timeout=15000)
            except Exception:
                log("⚠️ Não encontrei o menu 'Nova Mensagem' — pode ser necessário "
                    "selecionar o certificado manualmente (popup do Windows/Chrome) "
                    "ou fazer login nesta janela. Aguardando até 90s.")
                rastro.marco(page, "aguardando_certificado")
                page.wait_for_selector("text=Nova Mensagem", timeout=90000)

            rastro.marco(page, "caixa_postal")

            log("🖱️ Clicando em 'Nova Mensagem'...")
            page.click("text=Nova Mensagem")
            time.sleep(1.5)
            rastro.marco(page, "nova_mensagem")

            # O serviço "Envio de arquivo SEFIP" NEM SEMPRE vem pré-selecionado
            # — em sessão nova aparece vazio ("Selecione o serviço") e precisa
            # ser escolhido antes de os demais campos aparecerem/habilitarem.
            _selecionar_mat_select(page, SEL_SERVICO, "Envio de arquivo SEFIP",
                                   "servico", rastro)
            time.sleep(0.8)

            log(f"✏️ Preenchendo Nome da Mensagem: {nome_mensagem!r}")
            _localizar(page, SEL_NOME_MENSAGEM, "nome_mensagem", rastro).fill(nome_mensagem)
            time.sleep(0.3)

            _selecionar_mat_select(page, SEL_ESTADO, estado, "estado", rastro)
            time.sleep(0.5)

            # A Base de Arrecadação só popula opções DEPOIS do Estado ser
            # escolhido (dependência confirmada no mapeamento).
            _selecionar_mat_select(page, SEL_BASE, base_arrecadacao,
                                   "base_arrecadacao", rastro)
            time.sleep(0.5)

            log(f"📎 Anexando arquivo: {arquivo}")
            _localizar(page, SEL_ARQUIVO, "arquivo", rastro)  # garante presença
            page.set_input_files(SEL_ARQUIVO[0], arquivo)
            time.sleep(1)

            anexo_ok, enviar_ok, erro_arquivo = _conferir_anexo(page, arquivo, rastro)
            resultado["anexo_ok"] = anexo_ok
            resultado["enviar_habilitado"] = enviar_ok
            resultado["erro_arquivo"] = erro_arquivo

            # Recusa do site é caso à parte: não é falha do bot nem coisa que
            # retentar resolva. Sai com status próprio para o orquestrador
            # poder separar "erro técnico" de "arquivo/cadastro inválido" e
            # seguir para a próxima empresa em vez de abortar o lote.
            if erro_arquivo:
                resultado["status"] = "recusado_pelo_site"
                resultado["erro"] = erro_arquivo
                resultado["acao_sugerida"] = _explicar_recusa(erro_arquivo)
                rastro.marco(page, "formulario_recusado")
                if not confirmar_envio:
                    log("🔍 Mantendo o navegador aberto para você conferir a "
                        "mensagem do site...")
                    _aguardar_revisao_manual(page, rastro, min(timeout_revisao, 600))
                return resultado

            rastro.marco(page, "formulario_preenchido")
            log("=" * 70)
            log("Formulário preenchido.")
            log("=" * 70)

            if confirmar_envio and not (anexo_ok and enviar_ok):
                # Não clicar Enviar às cegas: um envio com anexo não registrado
                # geraria uma mensagem vazia no sistema oficial da Caixa.
                raise RuntimeError(
                    "Abortando o envio: o anexo não foi confirmado na lista de "
                    "arquivos e/ou o botão Enviar está desabilitado."
                )

            if not confirmar_envio:
                resultado["status"] = "preenchido_aguardando_revisao"
                log("ℹ️ --confirmar-envio NÃO informado: parando antes do Enviar.")
                log(f"👉 Revise a tela e clique 'Enviar' manualmente. O navegador "
                    f"fica aberto por até {timeout_revisao}s ou até você fechá-lo "
                    f"(Ctrl+C aqui também encerra).")
                _aguardar_revisao_manual(page, rastro, timeout_revisao)
                return resultado

            r = _confirmar_envio_e_baixar(page, rastro, dir_download=dir_download,
                                          nome_maquina=nome_maquina)
            resultado.update(r)
            # Se o cadastro de máquina acabou de acontecer, o token só existe
            # agora — captura para que a PRÓXIMA execução não precise repetir.
            if arquivo_token:
                try:
                    for cnpj, token in (page.evaluate(_JS_LER_TOKENS) or {}).items():
                        _salvar_token_no_disco(arquivo_token, cnpj, token, rastro)
                except Exception:
                    pass
            resultado["status"] = "enviado" if r.get("protocolo") else "enviado_sem_protocolo"
            return resultado

        except Exception as e:
            resultado["erro"] = f"{type(e).__name__}: {e}"
            log(f"❌ Falha: {resultado['erro']}")
            rastro.evento("excecao", texto=resultado["erro"])
            rastro.marco(page, "excecao")
            if not confirmar_envio:
                # Em modo de revisão, mantém a tela para inspeção visual em vez
                # de derrubar o Chrome junto com o processo.
                log("🔍 Mantendo o navegador aberto para inspeção do erro...")
                _aguardar_revisao_manual(page, rastro, min(timeout_revisao, 600))
            return resultado

        finally:
            if vigia is not None:
                resultado["criptocns_tratadas"] = vigia.tratadas
                vigia.parar()
            rastro.encerrar(context)
            # Com --cdp o navegador é do USUÁRIO: nunca fechar o contexto (isso
            # derrubaria as abas dele). Só desconecta.
            if cdp:
                try:
                    navegador.close()  # apenas encerra a conexão CDP
                except Exception:
                    pass
            elif confirmar_envio:
                try:
                    context.close()
                except Exception:
                    pass


def _aguardar_revisao_manual(page, rastro, timeout_s):
    """
    Segura o processo Python (e portanto o Chrome) enquanto o usuário revisa
    a tela. Sai quando a página/contexto é fechado, no timeout, ou com Ctrl+C.
    """
    limite = time.time() + timeout_s
    try:
        while time.time() < limite:
            if page.is_closed():
                rastro.log("🚪 Navegador fechado pelo usuário — encerrando.")
                return
            time.sleep(1)
        rastro.log(f"⏲️ Timeout de revisão ({timeout_s}s) atingido — encerrando.")
    except KeyboardInterrupt:
        rastro.log("⌨️ Ctrl+C — encerrando.")


def _confirmar_envio_e_baixar(page, rastro, dir_download=None, nome_maquina=None):
    """
    Clica 'Enviar' e conduz o fluxo pós-envio até baixar o arquivo de
    retorno (protocolo). Passos mapeados manualmente (2026-07-24),
    implementados de forma defensiva — cada etapa tem timeout próprio e loga
    claramente se algo não bater, em vez de travar esperando input.

    Fluxo esperado:
      1. Clicar 'Enviar'.
      2. Se aparecer desafio de token (1ª vez nessa máquina/navegador):
         clicar 'Não possuo token'.
      3. Modal 'Aviso de cadastro de máquina obrigatório para envio SEFIP':
         marcar 'Aceito os termos...' e clicar 'Sim'.
      4. Modal 'Máquina Cadastrada': clicar 'Baixar Token'.
      5. Tela 'Protocolo de Envio de Arquivos': clicar 'Baixar'.

    Retorna {"token": caminho|None, "protocolo": caminho|None}.
    """
    log = rastro.log
    saida = {"token": None, "protocolo": None}

    log("▶ Clicando 'Enviar'...")
    _clicar_botao(page, "Enviar", rastro)
    return _conduzir_pos_envio(page, rastro, dir_download, saida,
                               nome_maquina=nome_maquina)


# Telas possíveis depois do Enviar. A ORDEM em que aparecem varia: com token
# de máquina válido o site pula direto para o protocolo; sem token passa pelo
# desafio e pelo cadastro. Por isso não se espera uma sequência fixa — espera-se
# a PRIMEIRA tela que aparecer e reage-se a ela.
TELAS_POS_ENVIO = {
    "protocolo": "text=Protocolo de Envio de Arquivos",
    "maquina_cadastrada": "text=Máquina Cadastrada",
    "cadastro_maquina": "text=cadastro de máquina",
    "desafio_token": ":is(button, a):has-text('Não possuo token')",
    "erro": ".mat-error, .erro, text=/erro|falha|inv[áa]lid/i",
}


def _detectar_tela(page):
    """
    Devolve o nome da primeira tela conhecida que estiver visível, ou None.
    Usa is_visible() de propósito: aqui ele retorna na hora (sem esperar), que
    é o comportamento correto para varrer vários candidatos em loop.
    """
    for nome, sel in TELAS_POS_ENVIO.items():
        try:
            if page.locator(sel).first.is_visible():
                return nome
        except Exception:
            continue
    return None


def _conduzir_pos_envio(page, rastro, dir_download, saida, timeout=120,
                        nome_maquina=None):
    """
    Máquina de estados do pós-envio: observa qual tela apareceu e age, até
    chegar no protocolo (ou estourar o tempo). Tolera o fluxo curto (token
    válido → protocolo direto) e o longo (desafio → cadastro → token →
    protocolo), em qualquer ordem.
    """
    log = rastro.log
    limite = time.time() + timeout
    ja_visto = set()

    while time.time() < limite:
        tela = _detectar_tela(page)

        if tela is None:
            time.sleep(0.5)
            continue

        # Evita reprocessar a mesma tela em loop caso o clique não avance.
        if tela in ja_visto and tela != "protocolo":
            time.sleep(1)
            if _detectar_tela(page) == tela:
                log(f"⚠️ Tela '{tela}' não avançou após a ação — parando para "
                    f"inspeção manual.")
                rastro.marco(page, f"travou_{tela}")
                return saida
            continue
        ja_visto.add(tela)
        rastro.evento("tela_pos_envio", tela=tela)
        rastro.marco(page, f"tela_{tela}")

        if tela == "desafio_token":
            log("🔘 Desafio de token — clicando 'Não possuo token'...")
            page.locator(TELAS_POS_ENVIO["desafio_token"]).first.click()
            time.sleep(1.5)

        elif tela == "cadastro_maquina":
            log("🔘 Cadastro de máquina — aceitando termos...")
            if nome_maquina:
                try:
                    campo = page.locator("input[matinput][required]").first
                    campo.fill(nome_maquina[:30])
                    log(f"✏️ Nome da máquina: {nome_maquina[:30]!r}")
                except Exception as e:
                    log(f"ℹ️ Não consegui definir o nome da máquina "
                        f"({type(e).__name__}) — mantendo o sugerido pelo site.")
            _aceitar_termos_maquina(page, rastro)
            # O botão "Sim" NÃO existe no DOM antes do aceite (confirmado no
            # dump real): o Angular só o renderiza quando o checkbox é marcado.
            # Por isso espera-se o botão aparecer em vez de clicar direto.
            limite_sim = time.time() + 15
            while time.time() < limite_sim and _botao(page, "Sim") is None:
                time.sleep(0.5)
            _clicar_botao(page, "Sim", rastro, timeout=8000)
            time.sleep(2)

        elif tela == "maquina_cadastrada":
            log("🔘 Máquina cadastrada — baixando token...")
            try:
                saida["token"] = _baixar(page, rastro, "Baixar Token",
                                         dir_download, timeout=20000)
            except Exception as e:
                log(f"⚠️ Falha ao baixar o token ({type(e).__name__}) — seguindo.")
            time.sleep(1.5)

        elif tela == "protocolo":
            log("📄 Tela de Protocolo de Envio — ENVIO CONCLUÍDO.")
            _registrar_protocolo(page, rastro, saida)
            saida["dados"] = _extrair_dados_protocolo(page, rastro)
            # Prefixo pelo NRA: amarra os dois artefatos ao envio.
            prefixo = saida["dados"].get("nra")

            try:
                saida["protocolo"] = _baixar(page, rastro, "Baixar",
                                             dir_download, timeout=30000,
                                             prefixo=prefixo)
            except Exception as e:
                log(f"⚠️ Não consegui baixar o arquivo de retorno "
                    f"({type(e).__name__}). O ENVIO já foi feito — dá para "
                    f"baixar manualmente em 'Itens Enviados'.")

            # O comprovante em PDF é o documento que fica no dossiê da empresa.
            # É best-effort: se falhar, não invalida o envio nem o retorno.
            try:
                saida["pdf"] = _baixar(page, rastro, "Salvar PDF",
                                       dir_download, timeout=20000,
                                       prefixo=prefixo)
            except Exception as e:
                log(f"ℹ️ Não consegui salvar o PDF do comprovante "
                    f"({type(e).__name__}) — o protocolo já está no "
                    f"protocolo.json e em 'Itens Enviados'.")

            saida["pasta"] = dir_download
            _gravar_manifesto(dir_download, saida, rastro)
            rastro.marco(page, "final")
            return saida

        elif tela == "erro":
            texto = ""
            try:
                texto = page.locator(TELAS_POS_ENVIO["erro"]).first.inner_text()[:300]
            except Exception:
                pass
            log(f"❌ O site retornou erro: {texto!r}")
            rastro.evento("erro_site", texto=texto)
            return saida

    log("⚠️ Nenhuma tela conhecida apareceu dentro do tempo — verifique "
        "manualmente (o envio PODE ter ocorrido; confira 'Itens Enviados').")
    rastro.marco(page, "timeout_pos_envio")
    return saida


def _aceitar_termos_maquina(page, rastro):
    """
    Marca "Aceito os termos de cadastro de máquina".

    O mat-checkbox do Angular Material esconde o <input type="checkbox"> real
    com a classe `cdk-visually-hidden` e desenha um quadrado estilizado. Logo,
    `.check()` do Playwright falha ("element is not visible") — é preciso
    clicar no LABEL/componente, e confirmar o estado pelo input nativo.
    """
    nativo = page.locator(
        "input.mat-checkbox-input, mat-checkbox input[type=checkbox]").first

    def marcado():
        try:
            return nativo.is_checked()
        except Exception:
            return False

    if marcado():
        return True

    estrategias = [
        ("label", lambda: page.locator(
            "mat-checkbox label.mat-checkbox-layout").first.click()),
        ("mat-checkbox", lambda: page.locator("mat-checkbox").first.click()),
        ("texto", lambda: page.get_by_text("Aceito os termos").first.click()),
        ("check-force", lambda: nativo.check(force=True)),
    ]
    for nome, acao in estrategias:
        try:
            acao()
            time.sleep(0.4)
            if marcado():
                rastro.evento("termos_aceitos", via=nome)
                rastro.log(f"☑️ Termos aceitos (via {nome}).")
                return True
        except Exception as e:
            rastro.evento("aceite_falhou", via=nome, erro=type(e).__name__)

    rastro.marco(page, "aceite_falhou")
    raise RuntimeError(
        "Não consegui marcar 'Aceito os termos de cadastro de máquina' — "
        "sem isso o botão 'Sim' nem chega a aparecer.")


def _registrar_protocolo(page, rastro, saida):
    """Captura o número do protocolo exibido na tela, para log e conferência."""
    try:
        texto = page.locator("body").inner_text()
        # O protocolo real é um UUID (ex.: fc54b77d-f142-4f56-936d-9a7a0a809384),
        # confirmado no envio de 27/07/2026 — começa com LETRA, então a regex
        # antiga (que exigia dígito inicial) nunca casava. UUID primeiro;
        # formato numérico fica como alternativa.
        m = re.search(r"\b[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-"
                      r"[0-9a-f]{4}-[0-9a-f]{12}\b", texto, re.I)
        if not m:
            m = re.search(r"protocolo[^0-9]{0,40}([0-9][0-9./\- ]{6,})", texto, re.I)
        if m:
            saida["numero_protocolo"] = (m.group(1) if m.groups() else m.group(0)).strip()
            rastro.log(f"🧾 Protocolo: {saida['numero_protocolo']}")
            rastro.evento("protocolo", numero=saida["numero_protocolo"])
    except Exception:
        pass


def _baixar(page, rastro, texto_botao, dir_download, timeout=20000, prefixo=None):
    """
    Clica um botão que dispara download e salva o arquivo.

    Com `prefixo` (normalmente o NRA do arquivo enviado), o nome final vira
    `<NRA>_<nome sugerido>` — assim os artefatos ficam associados ao envio
    mesmo quando a Caixa devolve nomes genéricos, e a Etapa 6 consegue
    localizá-los sem adivinhação.
    """
    with page.expect_download(timeout=timeout) as info:
        _clicar_botao(page, texto_botao, rastro, timeout=timeout)
    download = info.value
    if not dir_download:
        rastro.log(f"ℹ️ Download '{download.suggested_filename}' recebido, mas "
                   f"sem pasta de destino: arquivo NÃO foi salvo.")
        return None
    os.makedirs(dir_download, exist_ok=True)
    nome = download.suggested_filename
    if prefixo and not nome.startswith(prefixo):
        nome = f"{prefixo}_{nome}"
    caminho = os.path.join(dir_download, nome)
    download.save_as(caminho)
    rastro.log(f"💾 Salvo em: {caminho}")
    return caminho


# Campos do "Protocolo de Envio de Arquivos" (texto confirmado no envio real
# de 27/07/2026). Viram protocolo.json ao lado dos downloads, para a Etapa 6
# consumir sem precisar reabrir o site nem ler PDF.
CAMPOS_PROTOCOLO = [
    ("armazenado_em", r"armazenados na Caixa Econ[ôo]mica Federal em\s*([\d/]+\s+[\d:]+)"),
    ("arquivo", r"protocolo do arquivo\s+(\S+?\.SFP)"),
    ("transmissor", r"Transmissor:\s*(.+)"),
    ("inscricao_transmissor", r"Inscri[çc][ãa]o do Transmissor:\s*(\S+)"),
    ("responsavel", r"Respons[áa]vel:\s*(.+)"),
    ("inscricao_responsavel", r"Inscri[çc][ãa]o do Respons[áa]vel:\s*(\S+)"),
    ("competencia", r"Compet[êe]ncia:\s*([\d/]+)"),
    ("nra", r"NRA:\s*(\S+)"),
    ("base_processamento", r"Base de Processamento:\s*(.+)"),
    ("codigo_recolhimento", r"C[óo]digo de Recolhimento:\s*(\d+)"),
]


def _extrair_dados_protocolo(page, rastro):
    """Lê os campos do comprovante direto da tela, como dict."""
    dados = {}
    try:
        texto = page.locator("body").inner_text()
    except Exception:
        return dados
    for chave, padrao in CAMPOS_PROTOCOLO:
        m = re.search(padrao, texto)
        if m:
            dados[chave] = m.group(1).strip()
    rastro.evento("dados_protocolo", **dados)
    return dados


def _gravar_manifesto(dir_download, saida, rastro):
    """
    Grava protocolo.json na pasta de destino — é o contrato com a Etapa 6:
    ela lê daqui o número do protocolo, o NRA e os caminhos dos arquivos,
    sem depender de parsear PDF nem de reabrir a Conectividade.
    """
    if not dir_download:
        return None
    try:
        os.makedirs(dir_download, exist_ok=True)
        caminho = os.path.join(dir_download, "protocolo.json")
        registro = dict(saida)
        registro["gerado_em"] = datetime.now().isoformat(timespec="seconds")
        with open(caminho, "w", encoding="utf-8") as f:
            json.dump(registro, f, ensure_ascii=False, indent=1)
        rastro.log(f"🧾 Manifesto salvo em: {caminho}")
        return caminho
    except Exception as e:
        rastro.log(f"⚠️ Não consegui gravar o manifesto: {type(e).__name__}: {e}")
        return None


def main():
    p = argparse.ArgumentParser(description="Bot Conectividade Social — envio de arquivo SEFIP")
    p.add_argument("--arquivo", required=True, help="Caminho completo do arquivo .SFP")
    p.add_argument("--nome-mensagem", dest="nome_mensagem", required=True,
                   help="Texto para o campo 'Nome da Mensagem'")
    p.add_argument("--estado", required=True, help="Texto da opção de Estado (ex.: 'Rio de Janeiro')")
    p.add_argument("--base-arrecadacao", dest="base_arrecadacao", required=True,
                   help="Texto da opção de Base de Arrecadação (ex.: 'Volta Redonda / RJ')")
    p.add_argument("--confirmar-envio", dest="confirmar_envio", action="store_true",
                   help="Clica Enviar automaticamente (sem isso, para antes para revisão manual)")
    p.add_argument("--dir-download", dest="dir_download",
                   help="Pasta de destino dos artefatos. Se omitido, usa "
                        f"<pasta do .SFP>/{SUBPASTA_PROTOCOLO}/ — assim os "
                        "arquivos ficam junto da empresa, prontos para a Etapa 6.")
    p.add_argument("--subpasta", default=SUBPASTA_PROTOCOLO,
                   help=f"Nome da subpasta criada ao lado do .SFP (padrão: {SUBPASTA_PROTOCOLO})")
    p.add_argument("--debug", action="store_true",
                   help="Grava trace, screenshots, HTML, rede e console em debug_conectividade/")
    p.add_argument("--dir-debug", dest="dir_debug",
                   help="Pasta dos artefatos de debug (padrão: debug_conectividade/<timestamp>)")
    p.add_argument("--cdp", nargs="?", const="http://localhost:9222", default=None,
                   help="Anexa a um Chrome já aberto via CDP em vez de abrir um "
                        "perfil novo (ex.: --cdp ou --cdp http://localhost:9222). "
                        "O Chrome precisa ter sido iniciado com "
                        "--remote-debugging-port=9222.")
    p.add_argument("--sem-vigia-criptocns", dest="vigiar_criptocns",
                   action="store_false",
                   help="Não trata automaticamente a janela nativa do CriptoCNS "
                        "(por padrão o bot clica 'Aceitar' nela).")
    p.add_argument("--criptocns-nao-perguntar", dest="criptocns_nao_perguntar",
                   action="store_true",
                   help="Marca 'Não perguntar novamente' na janela do CriptoCNS "
                        "(configuração PERSISTENTE da máquina).")
    p.add_argument("--nome-maquina", dest="nome_maquina",
                   help="Nome da máquina no cadastro (máx. 30 chars). Sem isso, "
                        "mantém o sugerido pelo site (CNPJ_data hora).")
    p.add_argument("--arquivo-token", dest="arquivo_token",
                   default=ARQUIVO_TOKEN_PADRAO,
                   help=f"JSON onde o token de máquina é guardado/reinjetado "
                        f"(padrão: {os.path.basename(ARQUIVO_TOKEN_PADRAO)}). "
                        f"Contém credencial — não versionar.")
    p.add_argument("--nao-injetar-token", dest="injetar_token",
                   action="store_false",
                   help="Não reinjeta o token salvo (força o caminho de cadastro "
                        "de máquina — útil para testar esse fluxo).")
    p.add_argument("--timeout-revisao", dest="timeout_revisao", type=int,
                   default=TIMEOUT_REVISAO_PADRAO,
                   help=f"Segundos que o navegador fica aberto para revisão manual "
                        f"(padrão: {TIMEOUT_REVISAO_PADRAO})")
    args = p.parse_args()

    if not os.path.exists(args.arquivo):
        print(f"❌ Arquivo não encontrado: {args.arquivo}")
        sys.exit(1)

    # Destino padrão: subpasta ao lado do próprio .SFP, para que retorno,
    # comprovante e manifesto fiquem na pasta da empresa — é o que a Etapa 6
    # vai consumir.
    if not args.dir_download:
        args.dir_download = os.path.join(
            os.path.dirname(os.path.abspath(args.arquivo)), args.subpasta)
    if args.confirmar_envio:
        print(f"📁 Artefatos do envio irão para: {args.dir_download}")

    resultado = enviar_arquivo_sefip(
        args.arquivo, args.nome_mensagem, args.estado, args.base_arrecadacao,
        confirmar_envio=args.confirmar_envio, dir_download=args.dir_download,
        debug=args.debug, dir_debug=args.dir_debug,
        timeout_revisao=args.timeout_revisao, cdp=args.cdp,
        arquivo_token=args.arquivo_token, injetar_token=args.injetar_token,
        nome_maquina=args.nome_maquina,
        vigiar_criptocns=args.vigiar_criptocns,
        criptocns_nao_perguntar=args.criptocns_nao_perguntar,
    )

    print("🏁 " + json.dumps(resultado, ensure_ascii=False))

    # Códigos de saída distintos para o orquestrador decidir o que fazer:
    #   0 = ok   |   2 = o SITE recusou (cadastro/arquivo: exige ação humana,
    #   retentar não adianta)   |   1 = erro técnico (pode valer retry)
    status = resultado.get("status", "erro")
    if status == "recusado_pelo_site":
        print("↳ Recusa do site: NÃO adianta repetir — corrija o cadastro/arquivo.")
        sys.exit(2)
    sys.exit(0 if status != "erro" else 1)


if __name__ == "__main__":
    main()
