r"""
Loop_FGTS.py — roda o recálculo de FGTS para várias rodadas (empresa +
competência), cada rodada de PONTA A PONTA (Domínio -> SEFIP, com backup
logo após o Executar -> Conectividade -> validar protocolo/guia) antes de
passar para a próxima.

🔴 CAUSA RAIZ REAL DO "ERRO 1117" (investigado e resolvido em 2026-08-04,
com bisseção binária de comprimento de caminho — ver histórico da sessão):
o diálogo nativo "Seleção do Arquivo de Saída" (Etapa 6 — Relatórios > GRF)
tem um limite de caminho de ~120 caracteres no campo Nome — BEM menor que o
MAX_PATH de 260 do Windows. Caminhos mais longos que isso fazem a Etapa 6
falhar com "Erro 1117 - arquivo de dados (.SFP) é inválido ou foi violado",
mesmo com o .SFP em si perfeitamente válido. Testado e confirmado: 119
caracteres funciona, 122 caracteres falha (limite exato entre os dois, não
determinado com mais precisão). Hipóteses descartadas ao longo da
investigação, TODAS refutadas por teste controlado: ordem do backup
(rodadas sem qualquer backup também falhavam), espaços no nome da pasta
(nome sem espaço mas longo também falhava), profundidade de pastas (pasta
extra de 1 caractere funcionava normalmente).

FIX: usar uma pasta RAIZ curta para `--dir-raiz` (ex.: "C:\FGTS", não um
caminho longo tipo ".\Users\...\Desktop\...\canella_dombot_sefip\results").
Isso garante margem suficiente mesmo com nomes de empresa/funcionário
longos na subpasta. Testado com sucesso: C:\FGTS\25\<funcionario>\<comp>\
(~70-90 caracteres típicos) sempre passa na Etapa 6.

O backup roda logo após o Executar, dentro do próprio Bot_Sefip.py (ver
_etapa_sefip) — voltou a ser assim depois de confirmado que a ordem do
backup NUNCA foi a causa do Erro 1117 (era o comprimento do caminho). Rodar
o backup aqui (não no final do lote inteiro) continua sendo necessário
pelo motivo original: 'Ferramentas > Fazer backup' captura a base INTEIRA
carregada no momento, não algo específico da competência — se ficasse para
o final de várias rodadas, o Import da rodada seguinte já teria
sobrescrito a base das rodadas anteriores.

Config: um JSON (lista de dicts) com uma entrada por rodada. Ver
`loop_rodadas.json` para o formato esperado (chaves = mesmos parâmetros do
Orquestrador_FGTS.py: empresa, empresa_nome, comp_inicial, comp_final,
data_atraso, responsavel, funcionario (lista), nome_mensagem, estado,
base_arrecadacao, com_backup (bool)).

Uso:
  python Loop_FGTS.py --config loop_rodadas.json --dir-raiz "C:\FGTS"
  python Loop_FGTS.py --config loop_rodadas.json --dir-raiz "C:\FGTS" --ate-etapa envio
"""

import argparse
import glob
import json
import os
import subprocess
import sys
import time

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

DIR_SCRIPT = os.path.dirname(os.path.abspath(__file__))


def _nome_pasta_competencia(comp_inicial: str, comp_final: str) -> str:
    ini = (comp_inicial or "").replace("/", "-")
    fim = (comp_final or "").replace("/", "-")
    if not ini:
        return "sem-competencia"
    if fim and fim != ini:
        return f"{ini}_a_{fim}"
    return ini


_CARACTERES_INVALIDOS_PASTA = '<>:"/\\|?*'

# Máximo de caracteres do nome da subpasta do funcionário. Ver
# _pasta_rodada: o caminho final é
#   {dir_raiz}\{empresa}\{funcionario}\{MM-AAAA}\{arquivo}
# e o diálogo "Seleção do Arquivo de Saída" da Etapa 6 (Relatórios > GRF)
# tem limite de caminho de ~120 caracteres no campo Nome — confirmado por
# bisseção ao vivo em 2026-08-04 (119 funciona, 122 falha; não é sobre
# espaço, backup ou profundidade de pasta, só comprimento total). Truncar
# aqui é a rede de segurança para nomes de funcionário muito longos —
# ainda assim use um --dir-raiz curto (ex.: "C:\FGTS"), já que o
# comprimento da raiz também consome a mesma margem de ~120 caracteres.
_MAX_CHARS_NOME_FUNCIONARIO = 20


def _nome_pasta_funcionario(funcionario: list) -> str:
    """
    O recálculo de FGTS é feito por funcionário — a mesma empresa pode ter
    várias rodadas independentes (um funcionário por vez), então a
    competência sozinha não é suficiente para não misturar resultados.
    Usa o primeiro nome/CPF da lista (uma rodada normalmente processa um
    funcionário só), sanitizado para ser um nome de pasta válido no
    Windows e curto o bastante para não estourar o limite de caminho da
    Etapa 6 (ver _MAX_CHARS_NOME_FUNCIONARIO).
    """
    if not funcionario:
        return "sem-funcionario"
    nome = str(funcionario[0]).strip()
    for c in _CARACTERES_INVALIDOS_PASTA:
        nome = nome.replace(c, "_")
    nome = "_".join(nome.split())
    nome = nome[:_MAX_CHARS_NOME_FUNCIONARIO].rstrip("_")
    return nome or "sem-funcionario"


def _pasta_rodada(dir_raiz: str, rodada: dict) -> str:
    return os.path.join(
        dir_raiz,
        rodada["empresa"],
        _nome_pasta_funcionario(rodada.get("funcionario")),
        _nome_pasta_competencia(rodada["comp_inicial"], rodada["comp_final"]),
    )


def _rotulo(rodada: dict) -> str:
    func = _nome_pasta_funcionario(rodada.get("funcionario"))
    return f"empresa {rodada['empresa']} / {func} / {rodada['comp_inicial']}"


def _rodar(args_cli, descricao):
    print("=" * 78)
    print(f"▶ {descricao}")
    print("   " + " ".join(f'"{a}"' if " " in a else a for a in args_cli))
    print("=" * 78)
    resultado = subprocess.run([sys.executable] + args_cli, cwd=DIR_SCRIPT)
    return resultado.returncode


def _fechar_janelas_residuais_dominio(log=print):
    """
    Rede de segurança ENTRE blocos do loop: fecha qualquer janela do
    Domínio (classe FNWND3190) diferente da janela principal — cobre
    "Seleção de Empregados", "Troca de empresas", popups "Atenção" e
    qualquer outra sobra, não só o caso "Seleção" original. Testado ao
    vivo: mesmo com o subprocesso da rodada anterior reportando sucesso e
    encerrando, o Domínio às vezes ainda está fechando essas janelas
    internamente quando a rodada seguinte manda F8 — o F8 então esbarra em
    "para trocar de empresa você terá que fechar todas as janelas
    abertas". Usa AttachThreadInput (mesma técnica do
    DomBot_GFIP._forcar_foreground) para garantir que o ESC realmente
    chegue à janela, mesmo que ela não esteja em foreground real.
    """
    try:
        import win32gui
        import win32process
        import win32api
        import win32con
    except ImportError:
        return

    def _forcar_foreground(hwnd) -> bool:
        try:
            fg = win32gui.GetForegroundWindow()
            tid_fg, _ = win32process.GetWindowThreadProcessId(fg) if fg else (0, 0)
            tid_alvo, _ = win32process.GetWindowThreadProcessId(hwnd)
            tid_atual = win32api.GetCurrentThreadId()
            attached_fg = attached_alvo = False
            if tid_fg and tid_fg != tid_atual:
                win32process.AttachThreadInput(tid_atual, tid_fg, True)
                attached_fg = True
            if tid_alvo and tid_alvo != tid_atual:
                win32process.AttachThreadInput(tid_atual, tid_alvo, True)
                attached_alvo = True
            try:
                if win32gui.IsIconic(hwnd):
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(hwnd)
                return True
            finally:
                if attached_fg:
                    win32process.AttachThreadInput(tid_atual, tid_fg, False)
                if attached_alvo:
                    win32process.AttachThreadInput(tid_atual, tid_alvo, False)
        except Exception:
            return False

    from pywinauto.keyboard import send_keys

    # Acha a janela principal do Domínio (maior FNWND3190 visível com
    # "omínio" no título) para não fechá-la por engano — só suas filhas.
    hwnd_principal = [0]

    def _cb_principal(hwnd, _):
        try:
            if not win32gui.IsWindowVisible(hwnd):
                return
            if win32gui.GetClassName(hwnd) != "FNWND3190":
                return
            t = win32gui.GetWindowText(hwnd)
            if "omínio" in t or "OMÍNIO" in t or "ominio" in t.lower():
                hwnd_principal[0] = hwnd
        except Exception:
            pass

    win32gui.EnumWindows(_cb_principal, None)
    if not hwnd_principal[0]:
        return  # Domínio nem está aberto — nada a fazer aqui

    # Traz o Domínio para frente ANTES de procurar — algumas janelas
    # filhas (ex.: "Seleção de Empregados") ficam marcadas como visíveis
    # mas escondidas por trás da principal, e só reagem a ESC quando o
    # Domínio realmente está em foreground.
    _forcar_foreground(hwnd_principal[0])
    time.sleep(0.3)

    # Allowlist de palavras-chave de popups/telas modais REAIS conhecidas
    # do Domínio. Bug encontrado ao vivo (2026-08-04): filtrar só por
    # "tem título" não basta — "Barra Domínio Atendimento" é um painel
    # FIXO e legítimo (sempre visível, sempre com título), e ESC nele
    # nunca fecha nada — o loop insistia até esgotar as 6 tentativas a
    # cada rodada, desperdiçando ~10s por vez. Restrito às telas que
    # sabemos ser popups/modais transitórios que de fato bloqueiam o F8.
    TERMOS_POPUP_DOMINIO = ["seleção", "selecao", "troca de empresas", "aviso", "atenção",
                            "atencao", "informação", "informacao"]

    def _eh_popup_conhecido(hwnd) -> bool:
        try:
            t = win32gui.GetWindowText(hwnd).lower()
        except Exception:
            return False
        return any(termo in t for termo in TERMOS_POPUP_DOMINIO)

    for tentativa in range(6):
        alvo = [0]

        # 🔴 Bug encontrado ao vivo: EnumWindows só enumera janelas
        # TOP-LEVEL. A janela "Seleção de Empregados" (e outras) pode
        # existir como FILHA da janela principal do Domínio (visível via
        # EnumChildWindows, GetParent == hwnd do Domínio) sem nunca
        # aparecer em EnumWindows — por isso as checagens anteriores
        # sempre relatavam "nenhuma janela residual" mesmo com ela
        # confirmadamente aberta e bloqueando o F8 da troca de empresa
        # seguinte. Verificamos as DUAS formas agora.
        def _cb_top(hwnd, _):
            try:
                if hwnd == hwnd_principal[0]:
                    return
                if not win32gui.IsWindowVisible(hwnd):
                    return
                if win32gui.GetClassName(hwnd) != "FNWND3190":
                    return
                if not _eh_popup_conhecido(hwnd):
                    return
                alvo[0] = hwnd
            except Exception:
                pass

        win32gui.EnumWindows(_cb_top, None)

        if not alvo[0]:
            def _cb_filho(hwnd, _):
                try:
                    if not win32gui.IsWindowVisible(hwnd):
                        return
                    if win32gui.GetClassName(hwnd) != "FNWND3190":
                        return
                    if not _eh_popup_conhecido(hwnd):
                        return
                    alvo[0] = hwnd
                except Exception:
                    pass
            try:
                win32gui.EnumChildWindows(hwnd_principal[0], _cb_filho, None)
            except Exception:
                pass

        if not alvo[0]:
            if tentativa == 0:
                log("✔ Nenhuma janela residual do Domínio (FNWND3190) a fechar.")
            return

        titulo_alvo = win32gui.GetWindowText(alvo[0])
        log(f"🔻 [Loop] Janela do Domínio residual '{titulo_alvo}' encontrada (tentativa "
            f"{tentativa + 1}/6) — forçando foco e fechando.")
        _forcar_foreground(hwnd_principal[0])
        time.sleep(0.1)
        _forcar_foreground(alvo[0])
        time.sleep(0.1)
        send_keys('{ESC}')
        time.sleep(0.3)


ETAPAS = ["dominio", "sefip", "envio", "validar"]


def _etapa_dominio(rodada, dir_raiz):
    dir_saida_rodada = _pasta_rodada(dir_raiz, rodada)
    # Pausa curta antes de checar: rede de segurança leve (a causa raiz
    # real — a janela "Seleção de Empregados" existindo como FILHA
    # oculta, invisível a EnumWindows — já foi corrigida em
    # _fechar_janelas_residuais_dominio). 1s é suficiente para o Domínio
    # assentar após o processo anterior encerrar.
    time.sleep(1)
    _fechar_janelas_residuais_dominio()
    rc = _rodar([
        os.path.join(DIR_SCRIPT, "DomBot_GFIP.py"),
        "--empresa", rodada["empresa"],
        "--empresa-nome", rodada["empresa_nome"],
        "--comp-inicial", rodada["comp_inicial"],
        "--comp-final", rodada["comp_final"],
        "--data-atraso", rodada["data_atraso"],
        "--dir-saida", dir_raiz,
        "--subpasta", _nome_pasta_funcionario(rodada.get("funcionario")),
    ] + (["--cod-recolhimento", rodada["cod_recolhimento"]] if rodada.get("cod_recolhimento") else [])
      + (["--responsavel", rodada["responsavel"]] if rodada.get("responsavel") else []),
        f"DomBot_GFIP — {_rotulo(rodada)}")
    _fechar_janelas_residuais_dominio()
    if rc != 0:
        return False, f"DomBot_GFIP falhou (código {rc})"
    caminho_re = os.path.join(dir_saida_rodada, "Sefip.re")
    if not os.path.exists(caminho_re):
        return False, f"Esperava '{caminho_re}' após o DomBot_GFIP mas não achei"
    return True, None


def _etapa_sefip(rodada, dir_raiz):
    dir_saida_rodada = _pasta_rodada(dir_raiz, rodada)
    caminho_re = os.path.join(dir_saida_rodada, "Sefip.re")
    # Backup roda AQUI, dentro do próprio Bot_Sefip.py, logo após o
    # Executar desta rodada — protege o estado antes do próximo Import
    # (rodada seguinte) sobrescrever a base. A ordem do backup NÃO causa
    # o Erro 1117 (hipótese testada e descartada — ver docstring do topo
    # do arquivo); a causa real é o limite de ~120 caracteres no diálogo
    # da Etapa 6, mitigado usando um --dir-raiz curto.
    args_backup = []
    if rodada.get("com_backup", True):
        caminho_backup = rodada.get("dir_backup") or os.path.join(dir_saida_rodada, "backup_sefip")
        args_backup = ["--dir-backup", caminho_backup]
    rc = _rodar([
        os.path.join(DIR_SCRIPT, "Bot_Sefip.py"),
        "--tudo",
        "--arquivo", caminho_re,
        "--funcionario", *rodada["funcionario"],
        "--salvar",
        "--dir-saida", dir_saida_rodada,
    ] + args_backup,
        f"Bot_Sefip — {_rotulo(rodada)}")
    if rc != 0:
        return False, f"Bot_Sefip falhou (código {rc})"
    sfps = glob.glob(os.path.join(dir_saida_rodada, "*.SFP")) + \
           glob.glob(os.path.join(dir_saida_rodada, "*.sfp"))
    if not sfps:
        return False, f"Nenhum .SFP encontrado em '{dir_saida_rodada}'"
    return True, None


def _etapa_envio(rodada, dir_raiz, criptocns_nao_perguntar):
    dir_saida_rodada = _pasta_rodada(dir_raiz, rodada)
    sfps = sorted(glob.glob(os.path.join(dir_saida_rodada, "*.SFP")) +
                  glob.glob(os.path.join(dir_saida_rodada, "*.sfp")),
                  key=os.path.getmtime, reverse=True)
    caminho_sfp = sfps[0]
    rc = _rodar([
        os.path.join(DIR_SCRIPT, "Bot_Conectividade.py"),
        "--arquivo", caminho_sfp,
        "--nome-mensagem", rodada["nome_mensagem"],
        "--estado", rodada["estado"],
        "--base-arrecadacao", rodada["base_arrecadacao"],
        "--confirmar-envio",
    ] + (["--criptocns-nao-perguntar"] if criptocns_nao_perguntar else []),
        f"Bot_Conectividade — {_rotulo(rodada)}")
    if rc != 0:
        return False, f"Bot_Conectividade falhou (código {rc})"
    return True, None


def _etapa_validar(rodada, dir_raiz):
    dir_saida_rodada = _pasta_rodada(dir_raiz, rodada)
    pasta_protocolo = os.path.join(dir_saida_rodada, "protocolo")
    xmls = sorted(glob.glob(os.path.join(pasta_protocolo, "*.xml")),
                  key=os.path.getmtime, reverse=True)
    if not xmls:
        return False, f"Nenhum XML de protocolo em '{pasta_protocolo}' (envio pode não ter concluído)", None
    caminho_xml = xmls[0]

    sfps = sorted(glob.glob(os.path.join(dir_saida_rodada, "*.SFP")) +
                  glob.glob(os.path.join(dir_saida_rodada, "*.sfp")),
                  key=os.path.getmtime, reverse=True)
    caminho_sfp = sfps[0]
    nome_pdf = f"Guia_{os.path.splitext(os.path.basename(caminho_sfp))[0]}.pdf"
    caminho_pdf = os.path.join(dir_saida_rodada, nome_pdf)

    rc = _rodar([
        os.path.join(DIR_SCRIPT, "Bot_Sefip_Protocolo.py"),
        "--validar-protocolo",
        "--xml", caminho_xml,
        "--sfp", caminho_sfp,
        "--pdf-saida", caminho_pdf,
    ], f"Bot_Sefip_Protocolo (validar) — {_rotulo(rodada)}")

    if rc == 2:
        # Código 2 = 'Erro 1117 - arquivo .SFP inválido ou foi violado'.
        # 🔴 Causa raiz real (ver docstring do topo do arquivo): o campo
        # Nome do diálogo dessa tela tem limite de ~120 caracteres — NÃO
        # é sobre ordem de backup (hipótese testada e descartada). Se
        # esse erro aparecer, o mais provável é --dir-raiz produzindo um
        # caminho longo demais para esta rodada específica (empresa +
        # subpasta do funcionário + competência) — use uma raiz mais
        # curta (ex.: "C:\FGTS").
        return False, "Erro 1117 (.SFP inválido/violado) — provável caminho longo demais (limite ~120 caracteres no diálogo da Etapa 6); use --dir-raiz mais curto", "erro_1117"
    if rc != 0:
        return False, f"Validação do protocolo falhou (código {rc})", None
    return True, None, None


def processar_rodada_completa(rodada, dir_raiz, criptocns_nao_perguntar, ate_etapa):
    """
    Roda uma rodada de PONTA A PONTA: Domínio -> SEFIP (com backup logo
    após o Executar, dentro do próprio Bot_Sefip.py) -> Conectividade ->
    validar protocolo/guia.

    `ate_etapa`: para depois desta etapa (uma de ETAPAS), não executa as
    seguintes. None = roda tudo.

    Retorna (sucesso: bool, motivo_erro: str|None, codigo_erro: str|None).
    codigo_erro é usado pelo chamador para decidir se a falha é isolada
    (ex.: 'erro_1117') ou deve abortar o lote inteiro.
    """
    print(f"\n########## Rodada: {_rotulo(rodada)} ##########\n")

    ok, motivo = _etapa_dominio(rodada, dir_raiz)
    if not ok:
        return False, motivo, None
    if ate_etapa == "dominio":
        return True, None, None

    ok, motivo = _etapa_sefip(rodada, dir_raiz)
    if not ok:
        return False, motivo, None
    if ate_etapa == "sefip":
        return True, None, None

    ok, motivo = _etapa_envio(rodada, dir_raiz, criptocns_nao_perguntar)
    if not ok:
        return False, motivo, None
    if ate_etapa == "envio":
        return True, None, None

    ok, motivo, codigo = _etapa_validar(rodada, dir_raiz)
    if not ok:
        return False, motivo, codigo

    return True, None, None


def main():
    p = argparse.ArgumentParser(description="Roda o recálculo de FGTS ponta a ponta por rodada (empresa+competência)")
    p.add_argument("--config", required=True, help="Caminho do JSON com a lista de rodadas")
    p.add_argument("--dir-raiz", dest="dir_raiz", required=True, help="Pasta raiz de saída (ex.: .\\results)")
    p.add_argument("--ate-etapa", dest="ate_etapa", choices=ETAPAS, default=None,
                   help="Para cada rodada depois desta etapa (dominio/sefip/envio/validar/backup). "
                        "Padrão: roda todas as etapas, incluindo backup por último.")
    p.add_argument("--criptocns-nao-perguntar", dest="criptocns_nao_perguntar", action="store_true",
                   help="Marca 'Não perguntar novamente' nos popups do CriptoCNS durante o envio")
    args = p.parse_args()

    args.dir_raiz = os.path.abspath(args.dir_raiz)

    with open(args.config, encoding="utf-8") as f:
        rodadas = json.load(f)
    if not isinstance(rodadas, list) or not rodadas:
        print("❌ Config precisa ser uma lista JSON não-vazia de rodadas.")
        sys.exit(1)

    print(f"📋 {len(rodadas)} rodada(s) carregada(s) de '{args.config}' — cada uma ponta a ponta "
          f"(Domínio → SEFIP → Conectividade → validar guia → backup):")
    for r in rodadas:
        print(f"   - {_rotulo(r)}")

    rodadas_com_falha_isolada = []
    for rodada in rodadas:
        ok, motivo, codigo = processar_rodada_completa(
            rodada, args.dir_raiz, args.criptocns_nao_perguntar, args.ate_etapa)
        if ok:
            continue
        if codigo == "erro_1117":
            # Falha isolada conhecida — não deveria mais acontecer agora
            # que o backup roda por último, mas se acontecer não trava o
            # lote inteiro (só essa rodada precisa ser revista).
            print(f"⚠️ {_rotulo(rodada)}: {motivo}. Continuando com as demais rodadas do lote.")
            rodadas_com_falha_isolada.append((rodada, motivo))
            continue
        print(f"❌ {_rotulo(rodada)}: {motivo}. Loop interrompido.")
        sys.exit(1)

    print("\n########## Todas as rodadas processadas. ##########")
    if rodadas_com_falha_isolada:
        print(f"\n⚠️ {len(rodadas_com_falha_isolada)} rodada(s) com falha isolada:")
        for r, motivo in rodadas_com_falha_isolada:
            print(f"   - {_rotulo(r)}: {motivo}")
    print("\nLembrete: 'Ferramentas > Limpar Base de Dados' continua manual e por rodada — "
          "use 'python Bot_Sefip_Protocolo.py --limpar-base --confirmar-limpeza' quando "
          "tiver certeza de que quer resetar o SEFIP.")


if __name__ == "__main__":
    main()
