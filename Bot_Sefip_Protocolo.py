"""
Bot SEFIP — Protocolo (Etapa 6)
===============================
Valida no SEFIP o arquivo de protocolo devolvido pela Conectividade Social
depois do envio (Etapa 5, feita pelo Bot_Conectividade.py).

O QUE A ETAPA 5 DEIXA PRONTO
----------------------------
Na pasta da empresa:

    results/404/
        MD9ghVk735n00008.SFP                          <- arquivo enviado
        protocolo/
            FC54B77D-...-9A7A0A809384.xml             <- protocolo assinado
            protocolo270720261541.pdf                 <- comprovante
            protocolo.json                            <- manifesto (opcional)

ANATOMIA DO XML DE PROTOCOLO (real, envio de 27/07/2026)
--------------------------------------------------------
    <?xml-stylesheet type="text/xsl"?>   <- por isso o navegador mostra como página
    <header>
      fileversion = W08.40        srv_comp  = 202211     (competência AAAAMM)
      srv_codrec  = 115           filetype  = SFP
      who         = <CNPJ>        filename  = <nome da mensagem>
    </header>
    <body SERVICE="RE"><info>C:\\CAIXA\\MD9ghVk735n00008.sfp</info></body>
    <footer>
      protocol_date / protocol_time / protocol_id
      <signature> PKCS#7 da Caixa sobre data+hora+protocolo+NRA </signature>
    </footer>

Implicações práticas:
  - o <info> aponta para onde o SEFIP espera encontrar o .sfp ORIGINAL;
  - a assinatura é da Caixa: o XML não pode ser editado nem remontado;
  - competência e código de recolhimento no XML têm que bater com o .sfp,
    senão estamos validando o protocolo de outro envio.

ESTADO DESTE SCRIPT
-------------------
As partes que NÃO dependem da tela do SEFIP já estão prontas e testadas:
localizar os artefatos, ler o XML, ler o manifesto e conferir a coerência
entre eles. A navegação na tela ainda não está escrita porque o caminho de
menu do SEFIP precisa ser mapeado na máquina real — para isso existe o
modo `--spy`, que enumera os menus e as janelas do SEFIP:

    python Bot_Sefip_Protocolo.py --spy

Rode com o SEFIP ABERTO (e este script como Administrador, pois o SEFIP roda
elevado — sem isso a UIPI bloqueia leitura/cliques). Ele imprime a árvore de
menus com os IDs de comando; com essa saída dá para escrever a navegação
exata, sem chute.

Conferência dos artefatos, sem tocar no SEFIP:

    python Bot_Sefip_Protocolo.py --conferir ".\\results\\404"

Autor: Hugo L. Almeida
"""

import argparse
import glob
import json
import os
import re
import sys
import time
import xml.etree.ElementTree as ET
from datetime import datetime

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

# Só existem no Windows; o restante do módulo (parsing/conferência) funciona
# em qualquer lugar, o que permite testar sem o SEFIP por perto.
try:
    import win32con
    import win32gui
    import win32gui_struct
    import win32process
    import win32api
    from pywinauto.keyboard import send_keys
    WIN32 = True
except ImportError:
    WIN32 = False

SUBPASTA_PROTOCOLO = "protocolo"

# IDs de comando do menu do SEFIP, mapeados via --spy em 2026-07-28. Owner-
# drawn (sem texto legível) — se o SEFIP for atualizado, reconfirmar com
# --spy antes de usar (ver sefip-automation.md para o dump completo).
CMD_RELATORIOS_GRF = 93
CMD_FERRAMENTAS_BACKUP = 101
CMD_FERRAMENTAS_LIMPAR_BASE = 110

BM_CLICK = 0x00F5


# ----------------------------------------------------------------------------
# Leitura dos artefatos (independente do SEFIP)
# ----------------------------------------------------------------------------

def ler_xml_protocolo(caminho: str) -> dict:
    """
    Lê o XML de protocolo devolvido pela Conectividade.

    Devolve um dict com os campos do header/footer, o caminho do .sfp que o
    protocolo referencia e se a assinatura está presente.
    """
    with open(caminho, "rb") as f:
        bruto = f.read()

    # O XML vem em iso-8859-1 e com PI de stylesheet; o ElementTree lida bem,
    # mas normalizamos a leitura para não depender do encoding declarado.
    try:
        raiz = ET.fromstring(bruto)
    except ET.ParseError:
        raiz = ET.fromstring(bruto.decode("iso-8859-1", "replace"))

    dados = {"arquivo_xml": os.path.abspath(caminho)}
    for param in raiz.iter("param"):
        nome = param.get("NAME")
        if nome:
            dados[nome] = (param.text or "").strip()

    for tag in ("protocol_date", "protocol_time", "protocol_id"):
        el = raiz.find(f".//{tag}")
        if el is not None:
            dados[tag] = (el.text or "").strip()

    info = raiz.find(".//body/info")
    dados["arquivo_referenciado"] = (info.text or "").strip() if info is not None else None
    if dados.get("arquivo_referenciado"):
        dados["nra"] = os.path.splitext(
            os.path.basename(dados["arquivo_referenciado"].replace("\\", "/")))[0]

    assinatura = raiz.find(".//signature")
    texto_ass = (assinatura.text or "").strip() if assinatura is not None else ""
    dados["assinado"] = len(texto_ass) > 100
    dados["tamanho_assinatura"] = len(texto_ass)

    # srv_comp vem como AAAAMM; a competência "humana" é MM/AAAA.
    comp = dados.get("srv_comp") or ""
    if re.fullmatch(r"\d{6}", comp):
        dados["competencia"] = f"{comp[4:]}/{comp[:4]}"

    return dados


def carregar_manifesto(pasta_protocolo: str) -> dict:
    caminho = os.path.join(pasta_protocolo, "protocolo.json")
    if not os.path.exists(caminho):
        return {}
    try:
        with open(caminho, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def localizar_artefatos(pasta_empresa: str, subpasta: str = SUBPASTA_PROTOCOLO) -> dict:
    """
    Encontra, na pasta da empresa, o .SFP e os artefatos do envio.

    Funciona mesmo sem protocolo.json (por exemplo, quando os downloads foram
    feitos manualmente pelo navegador) — o XML sozinho já basta.
    """
    pasta_empresa = os.path.abspath(pasta_empresa)
    pasta_prot = os.path.join(pasta_empresa, subpasta)

    sfps = sorted(glob.glob(os.path.join(pasta_empresa, "*.SFP")) +
                  glob.glob(os.path.join(pasta_empresa, "*.sfp")))
    xmls = sorted(glob.glob(os.path.join(pasta_prot, "*.xml")),
                  key=os.path.getmtime, reverse=True)
    pdfs = sorted(glob.glob(os.path.join(pasta_prot, "*.pdf")),
                  key=os.path.getmtime, reverse=True)

    return {
        "pasta_empresa": pasta_empresa,
        "pasta_protocolo": pasta_prot,
        "sfp": sfps[0] if sfps else None,
        "sfps_encontrados": sfps,
        "xml": xmls[0] if xmls else None,
        "xmls_encontrados": xmls,
        "pdf": pdfs[0] if pdfs else None,
        "manifesto": carregar_manifesto(pasta_prot),
    }


def conferir_coerencia(artefatos: dict) -> dict:
    """
    Garante que o protocolo corresponde ao .SFP daquela pasta.

    Validar o protocolo errado é o pior erro possível nesta etapa: o SEFIP
    aceitaria um selo de outro envio e a empresa ficaria com um comprovante
    que não corresponde à guia. Por isso a checagem é explícita e o script
    deve se recusar a prosseguir quando não bate.
    """
    problemas, avisos, info = [], [], {}

    if not artefatos.get("sfp"):
        problemas.append("Nenhum arquivo .SFP encontrado na pasta da empresa.")
    if not artefatos.get("xml"):
        problemas.append("Nenhum XML de protocolo encontrado em "
                         f"'{artefatos.get('pasta_protocolo')}'.")
    if problemas:
        return {"ok": False, "problemas": problemas, "avisos": avisos, "info": info}

    xml = ler_xml_protocolo(artefatos["xml"])
    info["xml"] = xml

    nra_sfp = os.path.splitext(os.path.basename(artefatos["sfp"]))[0]
    nra_xml = xml.get("nra")
    info["nra_sfp"], info["nra_xml"] = nra_sfp, nra_xml
    if nra_xml and nra_sfp.upper() != nra_xml.upper():
        problemas.append(
            f"O protocolo é de OUTRO arquivo: XML aponta para {nra_xml!r}, "
            f"mas a pasta tem {nra_sfp!r}.")

    if not xml.get("assinado"):
        problemas.append("O XML não tem assinatura da Caixa — protocolo inválido.")

    if len(artefatos.get("xmls_encontrados") or []) > 1:
        avisos.append(
            f"Há {len(artefatos['xmls_encontrados'])} XMLs na pasta; usando o "
            f"mais recente ({os.path.basename(artefatos['xml'])}). Confira se "
            f"não há protocolo de envio antigo misturado.")

    man = artefatos.get("manifesto") or {}
    if man:
        num_man = (man.get("numero_protocolo") or "").upper()
        num_xml = (xml.get("protocol_id") or "").upper()
        if num_man and num_xml and num_man != num_xml:
            problemas.append(
                f"protocolo.json diz {num_man}, mas o XML é {num_xml}.")
        comp_man = (man.get("dados") or {}).get("competencia")
        if comp_man and xml.get("competencia") and comp_man != xml["competencia"]:
            problemas.append(f"Competência divergente: manifesto {comp_man} "
                             f"x XML {xml['competencia']}.")
    else:
        avisos.append("Sem protocolo.json (downloads manuais?) — a conferência "
                      "usou apenas o XML.")

    # O SEFIP procura o .sfp no caminho gravado no protocolo.
    ref = xml.get("arquivo_referenciado")
    if ref:
        info["caminho_esperado_pelo_sefip"] = ref
        if not os.path.exists(ref):
            avisos.append(
                f"O protocolo referencia {ref!r}, que não existe nesta máquina. "
                f"Se o SEFIP exigir o arquivo nesse caminho, será preciso copiar "
                f"{os.path.basename(artefatos['sfp'])} para lá antes de validar.")

    return {"ok": not problemas, "problemas": problemas, "avisos": avisos, "info": info}


# ----------------------------------------------------------------------------
# Spy: mapeamento do SEFIP (para escrever a navegação sem chute)
# ----------------------------------------------------------------------------

def _cls(h):
    try: return win32gui.GetClassName(h) or ""
    except Exception: return ""


def _txt(h):
    try: return win32gui.GetWindowText(h) or ""
    except Exception: return ""


def _achar_sefip():
    """
    Mesma heurística do Bot_Sefip.py: classe/título TfrmPrincipalSEFIP.

    NÃO filtra por IsWindowVisible: quando o SEFIP é minimizado, a janela
    principal real (TfrmPrincipalSEFIP) fica com IsWindowVisible=False —
    é a janela-sombra do processo (TApplication, classe T mas sem ser
    "Tfrm...") que aparece minimizada na barra de tarefas. Se filtrássemos
    por visível aqui, sobraria só essa janela-sombra (título também
    "Sefip", casando no 2º critério) — testado na prática: rodar
    --limpar-base com o SEFIP minimizado conectou na TApplication por
    engano, o WM_COMMAND não teve efeito (TApplication não tem o menu de
    verdade) e o comando falhou com segurança, mas sem fazer nada.

    Por isso: prioriza SEMPRE matches por classe TfrmXxxSEFIP (a janela
    principal de verdade) sobre o critério mais fraco de título — mesmo
    que a Tfrm esteja oculta/minimizada, ela é preferível.
    """
    por_classe, por_titulo = [], []

    def _cb(h, _):
        try:
            c, t = _cls(h), _txt(h)
            if c.startswith("Tfrm") and "SEFIP" in c.upper():
                por_classe.append(h)
            elif t.upper().startswith("SEFIP") and c.startswith("T"):
                por_titulo.append(h)
        except Exception:
            pass
        return True

    win32gui.EnumWindows(_cb, None)
    achados = por_classe or por_titulo
    return achados


# ----------------------------------------------------------------------------
# Automação de tela — Etapa 6 (mapeada e testada ao vivo em 2026-07-28)
# ----------------------------------------------------------------------------

def _enum_filhos(hwnd, pred=None):
    """Todos os descendentes (não só filhos diretos) que satisfazem pred(h)."""
    achados = []

    def _cb(h, _):
        try:
            if pred is None or pred(h):
                achados.append(h)
        except Exception:
            pass
        return True

    try:
        win32gui.EnumChildWindows(hwnd, _cb, None)
    except Exception:
        pass
    return achados


def _norm(s):
    return (s or "").strip().lower().replace("&", "")


def _forcar_foreground(hwnd) -> bool:
    """
    SetForegroundWindow sozinho FALHA SILENCIOSAMENTE (sem exceção, sem
    trocar o foco de verdade) quando chamado por um processo que não está
    em foreground/não teve input recente — é a proteção "focus stealing
    prevention" do Windows. Isso é invisível em testes manuais interativos
    (o terminal que roda o script já está em foco), mas quebra sempre que
    o script roda desacoplado do terminal ativo (ex.: via ferramenta de
    background) — sintoma observado: WM_SETTEXT/BM_CLICK não fazem efeito
    nenhum, sem nenhum erro Python, porque a janela nunca ganhou foco real.

    Fix padrão do Windows: anexar a thread de input do processo atual à
    thread da janela alvo com AttachThreadInput antes de chamar
    SetForegroundWindow — isso contorna a proteção. Sempre desanexar depois
    (senão as duas threads ficam permanentemente ligadas).

    Também restaura a janela se estiver minimizada (SW_RESTORE) — sem isso
    SetForegroundWindow não traz uma janela minimizada de volta à tela, só
    troca o foco "por baixo" sem o usuário ver nada acontecer.

    🔴 Caso especial do SEFIP (testado 2026-07-28): ao minimizar o SEFIP, a
    janela principal (TfrmPrincipalSEFIP) reporta IsIconic=False mesmo
    estando efetivamente oculta (IsWindowVisible=False) — quem carrega o
    estado "minimizado" de verdade é a janela-sombra do processo
    (TApplication, dona/owner da principal via GetWindow(GW_OWNER), a que
    aparece na barra de tarefas). Restaurar só a TfrmPrincipalSEFIP não
    traz nada de volta à tela — é preciso restaurar a OWNER. Por isso,
    além de checar IsIconic(hwnd), também checamos e restauramos o owner.
    """
    try:
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        owner = win32gui.GetWindow(hwnd, win32con.GW_OWNER)
        if owner and win32gui.IsIconic(owner):
            win32gui.ShowWindow(owner, win32con.SW_RESTORE)
    except Exception:
        pass
    try:
        tid_atual = win32api.GetCurrentThreadId()
        tid_alvo, _ = win32process.GetWindowThreadProcessId(hwnd)
        anexado = False
        if tid_alvo != tid_atual:
            win32process.AttachThreadInput(tid_atual, tid_alvo, True)
            anexado = True
        try:
            win32gui.SetForegroundWindow(hwnd)
        finally:
            if anexado:
                win32process.AttachThreadInput(tid_atual, tid_alvo, False)
        return True
    except Exception:
        try:
            win32gui.SetForegroundWindow(hwnd)
        except Exception:
            pass
        return False


def _achar_popup(titulos=None, classes=("#32770", "TMessageForm")):
    """Primeiro popup visível cujo título bate (substring, case-insensitive)."""
    titulos = [t.lower() for t in (titulos or [])]

    def _pred(h):
        c = _cls(h)
        if not any(k in c for k in classes) and c not in classes:
            return False
        if not titulos:
            return True
        t = _txt(h).lower()
        return any(k in t for k in titulos)

    achados = []

    def _cb(h, _):
        try:
            if win32gui.IsWindowVisible(h) and _pred(h):
                achados.append(h)
        except Exception:
            pass
        return True

    win32gui.EnumWindows(_cb, None)
    return achados[0] if achados else 0


def _clicar_botao_por_texto(hwnd_pai, texto_alvo, classes=("TButton", "TBitBtn", "Button"),
                            assincrono=False):
    """
    Clica um botão descendente por texto (ignora &). `assincrono=True` (via
    PostMessage) é obrigatório para botões que disparam processamento
    bloqueante no SEFIP — mesma lição documentada no Bot_Sefip.py (Salvar de
    Modalidades, Executar). Traz a janela pra frente antes do clique: os
    popups TMessageForm do SEFIP não respondem a SendMessage sem foreground
    prévio (testado — sem isso o BM_CLICK não fecha o popup).
    """
    alvo = _norm(texto_alvo)
    try:
        _forcar_foreground(hwnd_pai)
        time.sleep(0.3)
    except Exception:
        pass
    for h in _enum_filhos(hwnd_pai, lambda h: _cls(h) in classes):
        if _norm(_txt(h)) == alvo:
            try:
                if assincrono:
                    win32gui.PostMessage(h, BM_CLICK, 0, 0)
                else:
                    win32gui.SendMessage(h, BM_CLICK, 0, 0)
                return True
            except Exception:
                return False
    return False


def _preencher_dialogo_abrir_arquivo(caminho_arquivo, timeout=10):
    """
    Preenche e confirma um diálogo nativo clássico do Windows (Abrir/Salvar,
    campo Nome ctrl_id=1152, botão ação ctrl_id=1) — usado em 3 pontos da
    Etapa 6: 'GRF - Arquivo ICP', 'Seleção do Arquivo de Saída' e 'Salvar
    arquivo backup'. GetWindowText no campo Nome não é confiável (retorna
    vazio mesmo com o texto visível na tela) — não usar para verificação,
    só WM_SETTEXT para escrever.
    """
    t0 = time.time()
    hdialog = 0
    while time.time() - t0 < timeout:
        hdialog = _achar_popup(classes=("#32770",))
        if hdialog:
            break
        time.sleep(0.3)
    if not hdialog:
        return False

    hedit = 0
    for h in _enum_filhos(hdialog, lambda h: _cls(h) == "Edit"):
        if win32gui.GetDlgCtrlID(h) == 1152:
            hedit = h
            break
    if not hedit:
        return False

    habrir = 0
    for h in _enum_filhos(hdialog, lambda h: _cls(h) == "Button"):
        if win32gui.GetDlgCtrlID(h) == 1:
            habrir = h
            break
    if not habrir:
        return False

    _forcar_foreground(hdialog)
    time.sleep(0.3)
    win32gui.SendMessage(hedit, win32con.WM_SETTEXT, 0, caminho_arquivo)
    time.sleep(0.3)
    win32gui.PostMessage(habrir, BM_CLICK, 0, 0)
    time.sleep(1)
    return True


class ValidacaoProtocolo:
    """
    Automação de tela da Etapa 6: valida o protocolo (Relatórios > GRF),
    gera a guia em PDF, faz backup e (opcionalmente) limpa a base do SEFIP.

    Todos os passos foram mapeados e testados ao vivo em 2026-07-28 (ver
    sefip-automation.md para o relato completo, incluindo os hwnds/ctrl_ids
    de cada tela). Os IDs de comando de menu (CMD_*) são owner-drawn — se o
    SEFIP for atualizado, reconfirmar com `--spy` antes de usar.
    """

    def __init__(self, logger=None):
        self.log = logger or (lambda msg: print(msg, flush=True))
        self.main_hwnd = 0

    def conectar(self) -> bool:
        janelas = _achar_sefip()
        if not janelas:
            self.log("❌ SEFIP não encontrado. Abra o sistema antes de iniciar.")
            return False
        self.main_hwnd = janelas[0]
        self.log(f"✅ SEFIP encontrado: '{_txt(self.main_hwnd)}' (hwnd={self.main_hwnd})")
        return True

    def _enviar_comando_menu(self, id_cmd):
        _forcar_foreground(self.main_hwnd)
        time.sleep(0.3)
        win32gui.PostMessage(self.main_hwnd, win32con.WM_COMMAND, id_cmd, 0)
        time.sleep(1)

    # ══════════════════════════════════════════════════════════════════
    # Relatórios > GRF: validar protocolo + gerar guia em PDF
    # ══════════════════════════════════════════════════════════════════

    def validar_protocolo_e_gerar_guia(self, caminho_xml_protocolo, caminho_sfp,
                                        caminho_pdf_saida, timeout=60) -> bool:
        """
        Executa o fluxo completo mapeado:
          Relatórios > GRF > abre o XML de protocolo > (erro 1156 esperado,
          sempre acontece — não é bug) > OK > seleciona o .SFP > Validar
          Arquivo > Visualizar > Ctrl+I > Imprimir(OK) > Salvar PDF.

        `caminho_xml_protocolo`: XML de protocolo assinado (ver
        localizar_artefatos/ler_xml_protocolo).
        `caminho_sfp`: o .SFP a selecionar na tela 'Localização do Arquivo
        .SFP' quando o erro 1156 aparecer — IMPORTANTE: use o .SFP GENUÍNO
        gerado pela Etapa 4, não o ZIP de retorno da Conectividade (mesmo
        nome, conteúdo diferente — ver nota em sefip-automation.md, o teste
        de mapeamento desta função usou o arquivo errado de propósito só
        para validar a estrutura de telas, não os dados da guia).
        `caminho_pdf_saida`: caminho completo (com nome) do PDF da guia.
        """
        self.log("📄 Etapa 6 — Relatórios > GRF (Arquivo ICP)")
        self._enviar_comando_menu(CMD_RELATORIOS_GRF)

        if not _preencher_dialogo_abrir_arquivo(caminho_xml_protocolo):
            self.log("❌ Diálogo 'GRF - Arquivo ICP' não abriu ou não foi possível preencher.")
            return False

        # Erro 1156 é ESPERADO sempre (o XML referencia um caminho fixo
        # C:\CAIXA\<nra>.sfp que nunca existe nesta máquina) — não tratar
        # como falha, só fechar o popup e seguir para a seleção manual do
        # .SFP, que é o próximo passo normal do fluxo.
        t0 = time.time()
        herro = 0
        while time.time() - t0 < timeout:
            herro = _achar_popup(titulos=["erro"])
            if herro:
                break
            time.sleep(0.4)
        if herro:
            self.log("ℹ️ Popup 'Erro' 1156 (esperado) — fechando com OK.")
            if not _clicar_botao_por_texto(herro, "OK", assincrono=True):
                self.log("⚠️ Não consegui fechar o popup de erro 1156 — verifique manualmente.")
                return False
            time.sleep(1)
        else:
            self.log("⚠️ Popup de erro 1156 não apareceu (inesperado, mas seguindo).")

        # Tela 'Localização do Arquivo .SFP' — clicar no ícone de busca
        # (TBitBtn sem texto) para abrir o diálogo 'Seleção do Arquivo de
        # Saída' e selecionar o .SFP. NÃO usar 'Buscar em todo o
        # computador' (lento, lista arquivos demais) nem digitar caminho
        # direto no TEdit da tela (não aceita digitação, testado).
        t0 = time.time()
        hbusca = 0
        while time.time() - t0 < timeout:
            hbusca = _achar_popup(titulos=["localização do arquivo"],
                                  classes=("TfrmTela_Busca", "#32770"))
            if hbusca:
                break
            time.sleep(0.4)
        if not hbusca:
            self.log("❌ Tela 'Localização do Arquivo .SFP' não apareceu.")
            return False

        hicone = 0
        for h in _enum_filhos(hbusca, lambda h: _cls(h) == "TBitBtn"):
            if not _txt(h):
                hicone = h
                break
        if not hicone:
            self.log("❌ Ícone de busca de arquivo não encontrado na tela de Localização.")
            return False

        self.log("🔎 Selecionando o arquivo .SFP...")
        _forcar_foreground(hbusca)
        time.sleep(0.3)
        win32gui.PostMessage(hicone, BM_CLICK, 0, 0)
        time.sleep(1.5)

        if not _preencher_dialogo_abrir_arquivo(caminho_sfp):
            self.log("❌ Diálogo 'Seleção do Arquivo de Saída' não abriu ou não foi possível preencher.")
            return False

        # Botão 'Validar Arquivo' só habilita depois do grid populado —
        # aguardar antes de clicar.
        t0 = time.time()
        hvalidar = 0
        while time.time() - t0 < timeout:
            for h in _enum_filhos(hbusca, lambda h: _cls(h) == "TButton"):
                if _norm(_txt(h)) == "validar arquivo" and win32gui.IsWindowEnabled(h):
                    hvalidar = h
                    break
            if hvalidar:
                break
            time.sleep(0.4)
        if not hvalidar:
            self.log("❌ Botão 'Validar Arquivo' não habilitou — arquivo não foi selecionado corretamente.")
            return False

        self.log("▶ Clicando 'Validar Arquivo'...")
        _forcar_foreground(hbusca)
        time.sleep(0.3)
        win32gui.PostMessage(hvalidar, BM_CLICK, 0, 0)
        time.sleep(2)

        # Tela 'GRF' (TfrmGridGuias) com a guia gerada — Visualizar.
        t0 = time.time()
        hgrid = 0
        while time.time() - t0 < timeout:
            hgrid = _achar_popup(classes=("TfrmGridGuias",))
            if hgrid:
                break
            time.sleep(0.4)
        if not hgrid:
            self.log("❌ Tela de guias (TfrmGridGuias) não apareceu após Validar Arquivo.")
            return False

        self.log("👁️ Clicando 'Visualizar'...")
        if not _clicar_botao_por_texto(hgrid, "Visualizar", classes=("TBitBtn",), assincrono=True):
            self.log("❌ Botão 'Visualizar' não encontrado.")
            return False
        time.sleep(2)

        # Preview (TQRPreview, não legível via win32) — Ctrl+I abre o
        # diálogo nativo de Impressão do Windows.
        t0 = time.time()
        hpreview = 0
        while time.time() - t0 < timeout:
            hpreview = _achar_popup(titulos=["grf"], classes=("TForm",))
            if hpreview:
                break
            time.sleep(0.4)
        if not hpreview:
            self.log("❌ Janela de preview da guia não apareceu.")
            return False

        self.log("🖨️ Enviando Ctrl+I (Imprimir)...")
        _forcar_foreground(hpreview)
        time.sleep(0.3)
        send_keys("^i")
        time.sleep(1.5)

        t0 = time.time()
        himprimir = 0
        while time.time() - t0 < timeout:
            himprimir = _achar_popup(titulos=["imprimir"])
            if himprimir:
                break
            time.sleep(0.4)
        if not himprimir:
            self.log("❌ Diálogo 'Imprimir' não apareceu após Ctrl+I.")
            return False

        self.log("▶ Confirmando impressora (Microsoft Print To PDF) — OK...")
        if not _clicar_botao_por_texto(himprimir, "OK", classes=("Button",), assincrono=True):
            self.log("❌ Botão OK do diálogo Imprimir não encontrado.")
            return False
        time.sleep(1.5)

        if not _preencher_dialogo_salvar_moderno(caminho_pdf_saida, timeout=timeout):
            self.log("❌ Diálogo 'Salvar Saída de Impressão como' falhou.")
            return False

        if not os.path.exists(caminho_pdf_saida):
            self.log(f"⚠️ PDF não encontrado em '{caminho_pdf_saida}' após salvar — confira manualmente.")
        else:
            self.log(f"✅ Guia salva em PDF: {caminho_pdf_saida}")

        # Fecha o preview e a tela de guias, voltando limpo.
        try:
            win32gui.PostMessage(hpreview, win32con.WM_CLOSE, 0, 0)
            time.sleep(1)
        except Exception:
            pass
        _clicar_botao_por_texto(hgrid, "Fechar", classes=("TBitBtn",), assincrono=True)
        time.sleep(1)
        return True

    # ══════════════════════════════════════════════════════════════════
    # Ferramentas > Fazer backup
    # ══════════════════════════════════════════════════════════════════

    def fazer_backup(self, caminho_backup, timeout_processamento=180) -> bool:
        """
        Ferramentas > Fazer backup > OK > salva no caminho informado. O
        SEFIP acrescenta '.zip' automaticamente ao nome dado (o backup sai
        sempre compactado) e o processamento é LENTO (~2min testado) —
        timeout_processamento generoso é importante aqui.
        """
        self.log("💾 Etapa 6 — Ferramentas > Fazer backup")
        self._enviar_comando_menu(CMD_FERRAMENTAS_BACKUP)

        hbackup = _achar_popup(classes=("TfrmBackup",))
        if not hbackup:
            self.log("❌ Tela 'Backup do SEFIP' não apareceu.")
            return False

        if not _clicar_botao_por_texto(hbackup, "OK", classes=("TBitBtn",), assincrono=True):
            self.log("❌ Botão OK da tela de backup não encontrado.")
            return False
        time.sleep(1)

        if not _preencher_dialogo_abrir_arquivo(caminho_backup):
            self.log("❌ Diálogo 'Salvar arquivo backup' não abriu ou não foi possível preencher.")
            return False

        self.log(f"⏳ Aguardando o backup ser gerado (pode levar minutos)...")
        t0 = time.time()
        ultimo_heartbeat = t0
        hsucesso = 0
        while time.time() - t0 < timeout_processamento:
            hsucesso = _achar_popup(titulos=["informação", "informacao"])
            if hsucesso:
                break
            if time.time() - ultimo_heartbeat > 10:
                self.log(f"⏳ ...ainda gerando backup ({int(time.time()-t0)}s)")
                ultimo_heartbeat = time.time()
            time.sleep(0.5)

        if not hsucesso:
            self.log("⚠️ Timeout aguardando confirmação de sucesso do backup — confira manualmente.")
            return False

        self.log("🔔 Backup concluído — fechando popup de confirmação.")
        _clicar_botao_por_texto(hsucesso, "OK", classes=("TButton",), assincrono=True)
        time.sleep(1)

        caminho_zip = caminho_backup + ".zip"
        if os.path.exists(caminho_zip):
            self.log(f"✅ Backup salvo em: {caminho_zip}")
        elif os.path.exists(caminho_backup):
            self.log(f"✅ Backup salvo em: {caminho_backup}")
        else:
            self.log(f"⚠️ Arquivo de backup não encontrado em '{caminho_backup}' nem '{caminho_zip}' — confira manualmente.")
        return True

    # ══════════════════════════════════════════════════════════════════
    # Ferramentas > Limpar Base de Dados
    # ══════════════════════════════════════════════════════════════════

    def limpar_base_dados(self, confirmar=False, timeout=15) -> bool:
        """
        Ferramentas > Limpar Base de Dados > Sim/Não.

        AÇÃO DESTRUTIVA E IRREVERSÍVEL — por isso `confirmar` é obrigatório
        e explícito (padrão False só abre e cancela, sem apagar nada). Só
        chamar com confirmar=True depois de já ter gerado a guia e feito o
        backup — não há como desfazer.
        """
        self.log("🧹 Etapa 6 — Ferramentas > Limpar Base de Dados")
        self._enviar_comando_menu(CMD_FERRAMENTAS_LIMPAR_BASE)

        t0 = time.time()
        hconfirma = 0
        while time.time() - t0 < timeout:
            hconfirma = _achar_popup(titulos=["confirmação", "confirmacao"])
            if hconfirma:
                break
            time.sleep(0.4)
        if not hconfirma:
            self.log("❌ Popup de confirmação de limpeza não apareceu.")
            return False

        if not confirmar:
            self.log("ℹ️ confirmar=False — cancelando (Não), SEM limpar a base.")
            _clicar_botao_por_texto(hconfirma, "Não", classes=("TButton",), assincrono=True)
            time.sleep(1)
            return True

        self.log("⚠️ Confirmando limpeza da base de dados (Sim) — ação irreversível.")
        _clicar_botao_por_texto(hconfirma, "Sim", classes=("TButton",), assincrono=True)
        time.sleep(1)

        # Segundo popup, de conclusão ("Informação", OK único) — testado ao
        # vivo (2026-07-28): sem tratar este popup, ele fica pendente na
        # tela mesmo com a limpeza já executada.
        t0 = time.time()
        hinfo = 0
        while time.time() - t0 < timeout:
            hinfo = _achar_popup(titulos=["informação", "informacao"])
            if hinfo:
                break
            time.sleep(0.4)
        if hinfo:
            self.log("🔔 Popup de conclusão detectado — fechando com OK.")
            _clicar_botao_por_texto(hinfo, "OK", classes=("TButton",), assincrono=True)
            time.sleep(1)
        else:
            self.log("⚠️ Popup de conclusão não apareceu dentro do timeout — confira manualmente.")

        return True


def _preencher_dialogo_salvar_moderno(caminho_arquivo, timeout=15):
    """
    Preenche o diálogo Explorer MODERNO 'Salvar Saída de Impressão como'
    (diferente do diálogo clássico — sem Edit ctrl_id=1152 direto). O campo
    de nome de arquivo real está mais fundo na árvore: Edit ctrl_id=1001,
    localizável só com EnumChildWindows recursivo completo (_enum_filhos já
    enumera todos os descendentes, não só filhos diretos).
    """
    t0 = time.time()
    hdialog = 0
    while time.time() - t0 < timeout:
        hdialog = _achar_popup(titulos=["salvar saída de impressão"], classes=("#32770",))
        if hdialog:
            break
        time.sleep(0.4)
    if not hdialog:
        return False

    hedit = 0
    for h in _enum_filhos(hdialog, lambda h: _cls(h) == "Edit"):
        if win32gui.GetDlgCtrlID(h) == 1001:
            hedit = h
            break
    if not hedit:
        return False

    hsalvar = 0
    for h in _enum_filhos(hdialog, lambda h: _cls(h) == "Button"):
        if win32gui.GetDlgCtrlID(h) == 1:
            hsalvar = h
            break
    if not hsalvar:
        return False

    _forcar_foreground(hdialog)
    time.sleep(0.3)
    win32gui.SendMessage(hedit, win32con.WM_SETTEXT, 0, caminho_arquivo)
    time.sleep(0.3)
    win32gui.PostMessage(hsalvar, BM_CLICK, 0, 0)
    time.sleep(1.5)
    return True


def _dump_menu(hmenu, profundidade=0, caminho="", saida=None):
    """
    Enumera recursivamente a barra de menus do SEFIP.

    Os itens do nível 0 (Arquivo/Relatórios/Ferramentas/...) têm texto legível
    via GetMenuItemInfo. Mas os SUBitens são owner-drawn pelo VCL — o texto
    vem sempre vazio (mesma limitação já documentada para o menu Arquivo em
    sefip-automation.md). Por isso, quando o rótulo vem vazio, usamos a
    POSIÇÃO (índice, 1-based) como identificador — é o que permite
    correlacionar cada linha impressa aqui com o que o usuário vê/conta na
    tela (ex.: "7º item do menu Ferramentas" = índice 6).
    """
    saida = saida if saida is not None else []
    try:
        total = win32gui.GetMenuItemCount(hmenu)
    except Exception:
        return saida

    for i in range(total):
        try:
            buf, _extras = win32gui_struct.EmptyMENUITEMINFO()
            win32gui.GetMenuItemInfo(hmenu, i, True, buf)
            info = win32gui_struct.UnpackMENUITEMINFO(buf)
            rotulo = (info.text or "").strip()
            sub = info.hSubMenu or 0
            id_cmd = info.wID if not sub else -1
        except Exception:
            rotulo, sub, id_cmd = "", 0, -1

        rotulo_exibicao = rotulo if rotulo else f"(item #{i})"
        completo = f"{caminho} > {rotulo_exibicao}" if caminho else rotulo_exibicao
        saida.append({"caminho": completo, "nivel": profundidade, "indice": i,
                      "id": None if sub else id_cmd, "tem_submenu": bool(sub),
                      "rotulo_legivel": bool(rotulo)})
        if sub:
            _dump_menu(sub, profundidade + 1, completo, saida)
    return saida


def spy():
    """Imprime menus e janelas do SEFIP para mapear a Etapa 6."""
    if not WIN32:
        print("❌ Este modo só roda no Windows (precisa de pywin32).")
        return 1

    janelas = _achar_sefip()
    if not janelas:
        print("❌ SEFIP não encontrado. Abra o SEFIP e rode de novo.")
        print("   Dica: rode este script como Administrador — o SEFIP roda "
              "elevado e a UIPI bloqueia a leitura de outro nível.")
        return 1

    hwnd = janelas[0]
    print("=" * 74)
    print(f"SEFIP: '{_txt(hwnd)}'  classe={_cls(hwnd)}  hwnd={hwnd}")
    print("=" * 74)

    hmenu = win32gui.GetMenu(hwnd)
    if not hmenu:
        print("\n⚠️ A janela principal não expôs menu via GetMenu — o SEFIP pode "
              "usar menu desenhado (VCL TMainMenu não-nativo). Nesse caso "
              "mapearemos por teclado/controles.")
    else:
        print("\n── ÁRVORE DE MENUS ──")
        print("(subitens são owner-drawn no VCL — sem texto legível; use a "
              "POSIÇÃO [#n, contando do topo, 0-based] pra bater com a tela)")
        for item in _dump_menu(hmenu):
            prefixo = "  " * item["nivel"]
            marca = "▸" if item["tem_submenu"] else " "
            id_txt = "" if item["id"] in (None, -1, 0) else f"   [id={item['id']}]"
            rotulo = item["caminho"].split(" > ")[-1]
            pos_txt = "" if item["rotulo_legivel"] else f"   (pos={item['indice']})"
            print(f"  {prefixo}{marca} {rotulo}{id_txt}{pos_txt}")

    print("\n── JANELAS TOP-LEVEL VISÍVEIS (classe / título) ──")
    def _cb(h, _):
        try:
            if win32gui.IsWindowVisible(h) and _txt(h):
                c = _cls(h)
                if c.startswith("T") or c == "#32770":
                    print(f"   {c:28} {_txt(h)[:60]}")
        except Exception:
            pass
        return True
    win32gui.EnumWindows(_cb, None)

    print("\n👉 Me mande esta saída inteira: com os rótulos e IDs reais eu escrevo "
          "a navegação da validação de protocolo sem chute.")
    return 0


# ----------------------------------------------------------------------------
# CLI
# ----------------------------------------------------------------------------

def _imprimir_conferencia(pasta):
    art = localizar_artefatos(pasta)
    print("=" * 74)
    print(f"Pasta da empresa : {art['pasta_empresa']}")
    print(f"Arquivo .SFP     : {art['sfp'] or '(não encontrado)'}")
    print(f"XML do protocolo : {art['xml'] or '(não encontrado)'}")
    print(f"Comprovante PDF  : {art['pdf'] or '(não encontrado)'}")
    print(f"Manifesto        : {'sim' if art['manifesto'] else 'não'}")
    print("=" * 74)

    res = conferir_coerencia(art)
    xml = res["info"].get("xml") or {}
    if xml:
        print(f"\nProtocolo   : {xml.get('protocol_id')}")
        print(f"Data/hora   : {xml.get('protocol_date')} {xml.get('protocol_time')}")
        print(f"Competência : {xml.get('competencia')}   "
              f"Cód. recolhimento: {xml.get('srv_codrec')}")
        print(f"NRA         : {xml.get('nra')}")
        print(f"Assinatura  : {'presente' if xml.get('assinado') else 'AUSENTE'} "
              f"({xml.get('tamanho_assinatura')} chars)")
        print(f"SEFIP espera o .sfp em: {xml.get('arquivo_referenciado')}")

    for a in res["avisos"]:
        print(f"\n⚠️  {a}")
    for p in res["problemas"]:
        print(f"\n❌ {p}")
    print("\n" + ("✅ Artefatos coerentes — prontos para validar no SEFIP."
                  if res["ok"] else "❌ NÃO validar no SEFIP com estes arquivos."))
    return 0 if res["ok"] else 1


def _esta_elevado() -> bool:
    """Ver Bot_Sefip.py — mesmo motivo: o SEFIP roda elevado, UIPI bloqueia
    cliques/teclas silenciosamente sem admin."""
    try:
        import ctypes
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def main():
    p = argparse.ArgumentParser(
        description="Bot SEFIP — Etapa 6 (validação do protocolo)")
    p.add_argument("--spy", action="store_true",
                   help="Mapeia menus/janelas do SEFIP (rodar com o SEFIP aberto, como Admin)")
    p.add_argument("--conferir", metavar="PASTA_EMPRESA",
                   help="Confere os artefatos do envio sem tocar no SEFIP")
    p.add_argument("--validar-protocolo", action="store_true",
                   help="Relatórios > GRF: valida o XML de protocolo e gera a guia em PDF")
    p.add_argument("--xml", help="Caminho do XML de protocolo (para --validar-protocolo)")
    p.add_argument("--sfp", help="Caminho do .SFP a selecionar quando o SEFIP pedir "
                                 "(erro 1156 esperado) — use o .SFP GENUÍNO da Etapa 4")
    p.add_argument("--pdf-saida", dest="pdf_saida",
                   help="Caminho completo (com nome) do PDF da guia a salvar")
    p.add_argument("--backup", action="store_true", help="Ferramentas > Fazer backup")
    p.add_argument("--dir-backup", dest="dir_backup",
                   help="Caminho completo (com nome, sem .zip) do arquivo de backup")
    p.add_argument("--limpar-base", action="store_true",
                   help="Ferramentas > Limpar Base de Dados (abre e cancela por padrão)")
    p.add_argument("--confirmar-limpeza", action="store_true",
                   help="Confirma de fato a limpeza da base (AÇÃO IRREVERSÍVEL) — "
                        "sem esta flag, --limpar-base só abre e clica Não")
    args = p.parse_args()

    if args.spy:
        sys.exit(spy())
    if args.conferir:
        sys.exit(_imprimir_conferencia(args.conferir))

    if args.validar_protocolo or args.backup or args.limpar_base:
        if not WIN32:
            print("❌ Este modo só roda no Windows (precisa de pywin32/pywinauto).")
            sys.exit(1)
        if not _esta_elevado():
            print("❌ Este processo NÃO está rodando como Administrador.")
            print("   O SEFIP roda elevado — sem admin, a UIPI bloqueia cliques/teclas")
            print("   de forma SILENCIOSA. Abra um PowerShell 'Executar como Administrador'.")
            sys.exit(1)

        bot = ValidacaoProtocolo()
        if not bot.conectar():
            sys.exit(1)

        if args.validar_protocolo:
            if not (args.xml and args.sfp and args.pdf_saida):
                print("❌ --validar-protocolo requer --xml, --sfp e --pdf-saida.")
                sys.exit(1)
            if not bot.validar_protocolo_e_gerar_guia(args.xml, args.sfp, args.pdf_saida):
                sys.exit(1)

        if args.backup:
            if not args.dir_backup:
                print("❌ --backup requer --dir-backup.")
                sys.exit(1)
            if not bot.fazer_backup(args.dir_backup):
                sys.exit(1)

        if args.limpar_base:
            if not bot.limpar_base_dados(confirmar=args.confirmar_limpeza):
                sys.exit(1)

        print("🏁 Etapa 6 concluída.")
        return

    p.print_help()
    print("\nExemplos:")
    print('  python Bot_Sefip_Protocolo.py --conferir ".\\results\\404"')
    print("  python Bot_Sefip_Protocolo.py --spy")
    print('  python Bot_Sefip_Protocolo.py --validar-protocolo --xml "...\\protocolo\\X.xml" '
          '--sfp "...\\resultado.SFP" --pdf-saida "...\\Guia.pdf"')
    print('  python Bot_Sefip_Protocolo.py --backup --dir-backup "...\\backup_404"')
    print("  python Bot_Sefip_Protocolo.py --limpar-base --confirmar-limpeza")


if __name__ == "__main__":
    main()
