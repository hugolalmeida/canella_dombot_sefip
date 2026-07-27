"""
SEFIP Spy — inspetor de janelas via Win32 puro
==============================================
O Inspect.exe / UI Automation NÃO conseguem ler a janela do SEFIP
(erro 0x80070005 "Acesso negado"), porque o SEFIP é uma aplicação
Delphi/VCL antiga que não expõe provider UI Automation.

Esta ferramenta substitui o Inspect para esse caso: usa apenas mensagens
Win32 clássicas (EnumWindows / EnumChildWindows / GetWindowText /
GetClassName / GetDlgCtrlID), que funcionam perfeitamente em apps VCL —
exatamente a técnica que o DomBot_GFIP.py já usa para o Domínio.

Uso:
    python sefip_spy.py                 # lista TODAS as janelas top-level
    python sefip_spy.py --sefip         # foca a janela principal do SEFIP e mapeia sua árvore
    python sefip_spy.py --title SEFIP   # mapeia a 1ª janela cujo título contém "SEFIP"
    python sefip_spy.py --pid 17688     # mapeia todas as janelas de um PID
    python sefip_spy.py --dialogs       # lista só diálogos/popups (#32770 e VCL modais) — útil p/ mapear avisos
    python sefip_spy.py --watch         # monitora aberturas/fechamentos de janela em tempo real (Ctrl+C p/ sair)

Saída: árvore identada com classe, ctrl_id, texto e handle de cada controle.
Os ctrl_ids e classes revelados aqui são o que o Bot_Sefip.py vai usar
para localizar botões/campos via PostMessage/SendMessage.

Autor: Hugo L. Almeida
"""

import argparse
import sys
import time

import win32gui
import win32process
import win32con

# Console do Windows costuma ser cp1252 e quebra com caracteres Unicode.
# Reconfigura stdout/stderr para UTF-8 (com replace) para não travar a saída.
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


# ── Leitura de propriedades de um controle ───────────────────────────────────

def _texto(hwnd: int) -> str:
    try:
        return win32gui.GetWindowText(hwnd) or ""
    except Exception:
        return ""


def _classe(hwnd: int) -> str:
    try:
        return win32gui.GetClassName(hwnd) or ""
    except Exception:
        return ""


def _ctrl_id(hwnd: int) -> int:
    try:
        return win32gui.GetDlgCtrlID(hwnd)
    except Exception:
        return 0


def _rect(hwnd: int):
    try:
        l, t, r, b = win32gui.GetWindowRect(hwnd)
        return (r - l, b - t)  # largura, altura
    except Exception:
        return (0, 0)


def _pid_de(hwnd: int) -> int:
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        return pid
    except Exception:
        return 0


def _descrever(hwnd: int) -> str:
    cls = _classe(hwnd)
    txt = _texto(hwnd)
    cid = _ctrl_id(hwnd)
    w, h = _rect(hwnd)
    vis = "V" if win32gui.IsWindowVisible(hwnd) else "-"
    partes = [f"[{vis}] {cls}"]
    if cid:
        partes.append(f"id={cid}")
    if txt:
        partes.append(f"'{txt}'")
    partes.append(f"({w}x{h})")
    partes.append(f"hwnd={hwnd}")
    return "  ".join(partes)


# ── Enumeração da árvore de filhos ───────────────────────────────────────────

def _mapear_filhos(hwnd: int, nivel: int = 1, max_nivel: int = 6):
    """Imprime recursivamente os filhos de uma janela, identados por nível."""
    if nivel > max_nivel:
        return
    filhos = []

    def _cb(ch, _):
        filhos.append(ch)
        return True

    try:
        win32gui.EnumChildWindows(hwnd, _cb, None)
    except Exception:
        return

    # EnumChildWindows já retorna TODA a descendência (netos inclusive) de forma
    # achatada; para desenhar a árvore de verdade, filtramos só os filhos diretos
    # e recursamos. GetParent identifica o pai imediato.
    diretos = []
    for ch in filhos:
        try:
            if win32gui.GetParent(ch) == hwnd:
                diretos.append(ch)
        except Exception:
            pass

    for ch in diretos:
        print("    " * nivel + "└─ " + _descrever(ch))
        _mapear_filhos(ch, nivel + 1, max_nivel)


# ── Listagem de janelas top-level ────────────────────────────────────────────

def _listar_toplevel(filtro_titulo: str = None, filtro_pid: int = None,
                     apenas_dialogos: bool = False):
    janelas = []

    def _cb(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return True
        txt = _texto(hwnd)
        cls = _classe(hwnd)
        pid = _pid_de(hwnd)

        if apenas_dialogos and cls != "#32770" and "TMessageForm" not in cls and "TForm" not in cls:
            return True
        if filtro_titulo and filtro_titulo.lower() not in txt.lower():
            return True
        if filtro_pid and pid != filtro_pid:
            return True
        # Ignora janelas sem título e sem classe relevante (ruído do shell)
        if not txt and not apenas_dialogos and not filtro_pid:
            return True
        janelas.append((hwnd, txt, cls, pid))
        return True

    try:
        win32gui.EnumWindows(_cb, None)
    except Exception:
        pass
    return janelas


def cmd_listar(args):
    print("=" * 78)
    print("JANELAS TOP-LEVEL VISÍVEIS")
    if args.title:
        print(f"Filtro título: contém '{args.title}'")
    if args.pid:
        print(f"Filtro PID: {args.pid}")
    if args.dialogs:
        print("Filtro: apenas diálogos/formulários (#32770, TMessageForm, TForm)")
    print("=" * 78)

    janelas = _listar_toplevel(args.title, args.pid, args.dialogs)
    if not janelas:
        print("Nenhuma janela encontrada com esses filtros.")
        return

    for hwnd, txt, cls, pid in janelas:
        print(f"\n► {_descrever(hwnd)}  [pid={pid}]")
        if args.tree:
            _mapear_filhos(hwnd, max_nivel=args.depth)


def _achar_sefip() -> int:
    """Retorna hwnd da janela principal do SEFIP (classe TfrmPrincipalSEFIP ou título)."""
    achado = [0]

    def _cb(hwnd, _):
        # Nunca retornar False: o pywin32 trata "parar cedo" como erro (win error 6).
        # Continuamos enumerando e ficamos com o primeiro match.
        if achado[0]:
            return True
        if not win32gui.IsWindowVisible(hwnd):
            return True
        cls = _classe(hwnd)
        txt = _texto(hwnd)
        # Match preciso: a classe da janela VCL do SEFIP começa com "Tfrm" e
        # contém "SEFIP" (ex.: TfrmPrincipalSEFIP). Assim evitamos falsos
        # positivos como o VS Code, cujo TÍTULO pode conter "sefip".
        if cls.startswith("Tfrm") and "SEFIP" in cls.upper():
            achado[0] = hwnd
            return True
        # Fallback secundário: título começando com "SEFIP -" e classe VCL (T*)
        if txt.upper().startswith("SEFIP") and cls.startswith("T"):
            achado[0] = hwnd
        return True

    try:
        win32gui.EnumWindows(_cb, None)
    except Exception:
        pass
    return achado[0]


def cmd_sefip(args):
    hwnd = _achar_sefip()
    if not hwnd:
        print("❌ Janela do SEFIP não encontrada. Abra o SEFIP antes de rodar o spy.")
        sys.exit(1)

    print("=" * 78)
    print("ÁRVORE DA JANELA PRINCIPAL DO SEFIP")
    print("=" * 78)
    print(f"► {_descrever(hwnd)}  [pid={_pid_de(hwnd)}]")
    _mapear_filhos(hwnd, max_nivel=args.depth)
    print("\n" + "=" * 78)
    print("DICA: anote os ctrl_id (id=) e classes dos campos/botões que você")
    print("      quer automatizar. O Bot_Sefip.py vai usá-los via PostMessage.")
    print("=" * 78)


def cmd_snapshot(args):
    """
    Captura AGORA, sem buffering, todas as janelas relevantes visíveis
    (SEFIP principal + popups/diálogos VCL + #32770) com seus filhos.
    Ideal para congelar o estado de uma tela/popup que está aberto neste momento.
    """
    print("=" * 78)
    print("SNAPSHOT — janelas relevantes visíveis neste momento")
    print("=" * 78)

    alvos = []

    def _cb(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return True
        cls = _classe(hwnd)
        txt = _texto(hwnd)
        # Janela principal do SEFIP, formulários VCL (TfrmXxx/TForm), popups
        # modais (#32770, TMessageForm) e qualquer T* com título.
        eh_sefip = cls.startswith("Tfrm")
        eh_dialog = cls in ("#32770",) or "TMessageForm" in cls
        eh_form_titulado = cls.startswith("T") and bool(txt)
        if eh_sefip or eh_dialog or eh_form_titulado:
            alvos.append(hwnd)
        return True

    try:
        win32gui.EnumWindows(_cb, None)
    except Exception:
        pass

    if not alvos:
        print("Nenhuma janela relevante visível agora.")
        return

    for hwnd in alvos:
        print(f"\n► {_descrever(hwnd)}  [pid={_pid_de(hwnd)}]")
        _mapear_filhos(hwnd, max_nivel=args.depth)


def cmd_watch(args):
    """Monitora aberturas/fechamentos de janela top-level em tempo real."""
    print("=" * 78)
    print("MONITOR DE JANELAS — abra/feche telas no SEFIP para vê-las aqui")
    print("Ctrl+C para sair")
    print("=" * 78)

    def _snapshot():
        atual = {}

        def _cb(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                atual[hwnd] = (_texto(hwnd), _classe(hwnd), _pid_de(hwnd))
            return True

        try:
            win32gui.EnumWindows(_cb, None)
        except Exception:
            pass
        return atual

    def _relevante(txt, cls):
        return bool(txt) or cls == "#32770" or "TForm" in cls or "TMessageForm" in cls

    anterior = _snapshot()
    try:
        while True:
            time.sleep(0.6)
            atual = _snapshot()
            for hwnd, info in atual.items():
                if hwnd not in anterior:
                    txt, cls, pid = info
                    if _relevante(txt, cls):
                        print(f"\n{'━'*70}")
                        print(f"➕ ABRIU:  [{cls}] '{txt}'  hwnd={hwnd} pid={pid}")
                        # Mapeia a árvore da janela recém-aberta automaticamente.
                        if args.tree:
                            _mapear_filhos(hwnd, max_nivel=args.depth)
                        print(f"{'━'*70}")
            for hwnd, info in anterior.items():
                if hwnd not in atual:
                    txt, cls, pid = info
                    if _relevante(txt, cls):
                        print(f"➖ FECHOU: [{cls}] '{txt}'  hwnd={hwnd}")
            anterior = atual
    except KeyboardInterrupt:
        print("\nMonitor encerrado.")


def main():
    p = argparse.ArgumentParser(description="Inspetor Win32 para o SEFIP (substitui o Inspect)")
    p.add_argument("--sefip", action="store_true", help="Mapeia a árvore da janela principal do SEFIP")
    p.add_argument("--title", help="Filtra/lista janelas cujo título contém este texto")
    p.add_argument("--pid", type=int, help="Filtra janelas deste PID")
    p.add_argument("--dialogs", action="store_true", help="Lista apenas diálogos/popups (#32770, TMessageForm, TForm)")
    p.add_argument("--snapshot", action="store_true", help="Captura AGORA todas as janelas/popups visíveis com seus filhos (sem buffering)")
    p.add_argument("--watch", action="store_true", help="Monitora aberturas/fechamentos de janela em tempo real")
    p.add_argument("--tree", action="store_true", help="Ao listar top-level, também mapeia a árvore de cada uma")
    p.add_argument("--depth", type=int, default=6, help="Profundidade máxima da árvore (padrão 6)")
    args = p.parse_args()

    if args.snapshot:
        cmd_snapshot(args)
    elif args.watch:
        cmd_watch(args)
    elif args.sefip:
        cmd_sefip(args)
    else:
        cmd_listar(args)


if __name__ == "__main__":
    main()
