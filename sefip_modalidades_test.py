"""
Teste de manipulação da janela "Modalidades" (Ctrl+M) do SEFIP via Win32
========================================================================
Objetivo: derriscar a parte mais crítica da automação — ler e controlar os
dois TListView (esquerda/direita), os combos de modalidade e os botões de
transferência (<<, <, >, >>) da janela TfrmAdmModalidades.

A janela usa ctrl_ids VCL instáveis, então localizamos os controles por
CLASSE + POSIÇÃO (painel esquerdo x direito), e os botões por TEXTO.

FASES:
  --ler        (padrão) Só LEITURA. Não altera nada. 100% seguro.
               Lista itens dos dois TListView, opções dos combos e os botões.
  --selecionar-mod "9"   Seleciona no combo DESTINO (direita) a modalidade
               cujo texto começa/contém o valor dado. ALTERA a tela.
  --clicar >>  Clica um botão de transferência por texto (<<, <, >, >>) ou
               'salvar'/'cancelar'. ALTERA a tela.
  --achar-func "CARLOS"  Procura nos dois TListView um item contendo o texto
               (nome/CPF) e informa em qual lista/índice está. Só leitura.

Uso típico de teste (com a janela Modalidades aberta):
  python sefip_modalidades_test.py --ler
  python sefip_modalidades_test.py --achar-func "CARLOS AUGUSTO"
  python sefip_modalidades_test.py --selecionar-mod "9"
  python sefip_modalidades_test.py --clicar ">>"

IMPORTANTE: rodar em PowerShell "Executar como Administrador" (SEFIP é elevado).
"""

import argparse
import sys
import time

import win32gui
import win32con
import win32api
import win32process
import ctypes
from ctypes import wintypes

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


# ── Constantes Win32 ─────────────────────────────────────────────────────────

BM_CLICK   = 0x00F5
CB_GETCOUNT      = 0x0146
CB_GETLBTEXTLEN  = 0x0149
CB_GETLBTEXT     = 0x0148
CB_SETCURSEL     = 0x014E
CB_GETCURSEL     = 0x0147

LVM_FIRST          = 0x1000
LVM_GETITEMCOUNT   = LVM_FIRST + 4
LVM_GETITEMTEXTW   = LVM_FIRST + 115
LVM_SETITEMSTATE   = LVM_FIRST + 43
LVM_GETNEXTITEM    = LVM_FIRST + 12
LVM_ENSUREVISIBLE  = LVM_FIRST + 19

LVNI_SELECTED = 0x0002
LVIS_SELECTED = 0x0002
LVIS_FOCUSED  = 0x0001


# ── Helpers de leitura de controle ───────────────────────────────────────────

def _cls(h):
    try: return win32gui.GetClassName(h) or ""
    except Exception: return ""

def _txt(h):
    try: return win32gui.GetWindowText(h) or ""
    except Exception: return ""

def _rect(h):
    try: return win32gui.GetWindowRect(h)
    except Exception: return (0, 0, 0, 0)


def _achar_modalidades():
    """Retorna hwnd da janela TfrmAdmModalidades, ou 0."""
    achado = [0]
    def _cb(h, _):
        if achado[0]:
            return True
        if not win32gui.IsWindowVisible(h):
            return True
        if _cls(h) == "TfrmAdmModalidades" or _txt(h) == "Modalidades":
            achado[0] = h
        return True
    try:
        win32gui.EnumWindows(_cb, None)
    except Exception:
        pass
    return achado[0]


def _coletar_controles(hwnd):
    """
    Enumera descendentes e agrupa por tipo, anotando posição x para separar
    painel esquerdo (x menor) de direito (x maior).
    Retorna dict com listas de (hwnd, x_left, texto).
    """
    listviews = []
    combos = []
    bitbtns = []

    def _cb(h, _):
        c = _cls(h)
        l, t, r, b = _rect(h)
        if c == "SysListView32" or c == "TListView":
            listviews.append((h, l))
        elif "ComboBox" in c or c == "TComboBoxPesqRef":
            combos.append((h, l, _txt(h)))
        elif c in ("TBitBtn", "TButton"):
            bitbtns.append((h, l, _txt(h)))
        return True

    try:
        win32gui.EnumChildWindows(hwnd, _cb, None)
    except Exception:
        pass

    listviews.sort(key=lambda x: x[1])   # esquerda primeiro
    combos.sort(key=lambda x: x[1])
    return {"listviews": listviews, "combos": combos, "bitbtns": bitbtns}


# ── Leitura de TListView via mensagens LVM_* (cross-process) ─────────────────

def _lv_count(hlv):
    return win32gui.SendMessage(hlv, LVM_GETITEMCOUNT, 0, 0)


def _lv_item_text(hlv, index, pid):
    """
    Lê o texto do item `index` de um ListView em OUTRO processo.
    Requer alocar a struct LVITEM e o buffer de texto na memória do processo
    dono do ListView (VirtualAllocEx), enviar LVM_GETITEMTEXTW e ler de volta.
    """
    PROCESS_VM = 0x0008 | 0x0010 | 0x0020  # READ|WRITE|OPERATION
    hproc = ctypes.windll.kernel32.OpenProcess(PROCESS_VM, False, pid)
    if not hproc:
        return "<sem acesso ao processo>"

    TEXT_LEN = 260
    buf_size = TEXT_LEN * 2  # WCHAR

    # LVITEMW: iSubItem em offset; pszText ponteiro; cchTextMax
    class LVITEMW(ctypes.Structure):
        _fields_ = [
            ("mask", wintypes.UINT),
            ("iItem", ctypes.c_int),
            ("iSubItem", ctypes.c_int),
            ("state", wintypes.UINT),
            ("stateMask", wintypes.UINT),
            ("pszText", ctypes.c_void_p),
            ("cchTextMax", ctypes.c_int),
            ("iImage", ctypes.c_int),
            ("lParam", ctypes.c_void_p),
            ("iIndent", ctypes.c_int),
            ("iGroupId", ctypes.c_int),
            ("cColumns", wintypes.UINT),
            ("puColumns", ctypes.c_void_p),
            ("piColFmt", ctypes.c_void_p),
            ("iGroup", ctypes.c_int),
        ]

    kernel32 = ctypes.windll.kernel32
    MEM_COMMIT = 0x1000
    PAGE_RW = 0x04

    remote_text = kernel32.VirtualAllocEx(hproc, 0, buf_size, MEM_COMMIT, PAGE_RW)
    remote_item = kernel32.VirtualAllocEx(hproc, 0, ctypes.sizeof(LVITEMW), MEM_COMMIT, PAGE_RW)

    try:
        item = LVITEMW()
        item.iSubItem = 0
        item.pszText = remote_text
        item.cchTextMax = TEXT_LEN

        written = ctypes.c_size_t(0)
        kernel32.WriteProcessMemory(hproc, remote_item, ctypes.byref(item),
                                    ctypes.sizeof(LVITEMW), ctypes.byref(written))

        win32gui.SendMessage(hlv, LVM_GETITEMTEXTW, index, remote_item)

        local_buf = (ctypes.c_wchar * TEXT_LEN)()
        read = ctypes.c_size_t(0)
        kernel32.ReadProcessMemory(hproc, remote_text, ctypes.byref(local_buf),
                                   buf_size, ctypes.byref(read))
        return local_buf.value
    finally:
        MEM_RELEASE = 0x8000
        kernel32.VirtualFreeEx(hproc, remote_text, 0, MEM_RELEASE)
        kernel32.VirtualFreeEx(hproc, remote_item, 0, MEM_RELEASE)
        kernel32.CloseHandle(hproc)


def _lv_listar(hlv, pid, limite=200):
    n = _lv_count(hlv)
    itens = []
    for i in range(min(n, limite)):
        itens.append(_lv_item_text(hlv, i, pid))
    return n, itens


# ── Leitura de combo (mesmo processo? não — cross-process também) ────────────

def _combo_listar(hcombo, pid):
    """
    Combos VCL (TComboBoxPesqRef) podem não responder a CB_GETLBTEXT como um
    ComboBox nativo. Tentamos CB_* e reportamos o que vier.
    """
    n = win32gui.SendMessage(hcombo, CB_GETCOUNT, 0, 0)
    atual = win32gui.SendMessage(hcombo, CB_GETCURSEL, 0, 0)
    txt_atual = _txt(hcombo)
    return n, atual, txt_atual


# ── Comandos ─────────────────────────────────────────────────────────────────

def cmd_ler(hwnd, pid):
    ctrls = _coletar_controles(hwnd)
    lvs = ctrls["listviews"]
    combos = ctrls["combos"]
    btns = ctrls["bitbtns"]

    print("=" * 70)
    print(f"Janela Modalidades hwnd={hwnd} pid={pid}")
    print("=" * 70)

    print(f"\n── ListViews encontrados: {len(lvs)} ──")
    for i, (hlv, x) in enumerate(lvs):
        lado = "ESQUERDA (Origem)" if i == 0 else ("DIREITA (Destino)" if i == 1 else f"#{i}")
        print(f"\n[{lado}] hwnd={hlv} x={x}")
        try:
            n, itens = _lv_listar(hlv, pid)
            print(f"  {n} item(ns):")
            for j, it in enumerate(itens):
                print(f"    {j:3d}: {it}")
        except Exception as e:
            print(f"  ERRO ao ler itens: {type(e).__name__}: {e}")

    print(f"\n── Combos encontrados: {len(combos)} ──")
    for i, (hc, x, t) in enumerate(combos):
        lado = "ESQUERDA (Origem)" if i == 0 else ("DIREITA (Destino)" if i == 1 else f"#{i}")
        n, atual, txt = _combo_listar(hc, pid)
        print(f"  [{lado}] hwnd={hc} x={x}  count={n} cursel={atual} texto_visível='{txt}'")

    print(f"\n── Botões (TBitBtn/TButton): {len(btns)} ──")
    for hb, x, t in btns:
        print(f"  hwnd={hb} x={x}  '{t}'")


def cmd_achar_func(hwnd, pid, alvo):
    ctrls = _coletar_controles(hwnd)
    alvo_up = alvo.upper()
    achou = False
    for i, (hlv, x) in enumerate(ctrls["listviews"]):
        lado = "ESQUERDA" if i == 0 else ("DIREITA" if i == 1 else f"#{i}")
        try:
            n, itens = _lv_listar(hlv, pid)
        except Exception as e:
            print(f"[{lado}] erro: {e}")
            continue
        for j, it in enumerate(itens):
            if alvo_up in it.upper():
                print(f"✅ Encontrado em [{lado}] índice {j}: {it}")
                achou = True
    if not achou:
        print(f"❌ '{alvo}' não encontrado em nenhuma das listas.")


def cmd_selecionar_mod(hwnd, pid, valor):
    ctrls = _coletar_controles(hwnd)
    combos = ctrls["combos"]
    if len(combos) < 2:
        print(f"❌ Esperava 2 combos, achei {len(combos)}. Abortando.")
        return
    hc_destino = combos[1][0]  # direita
    print(f"Combo DESTINO hwnd={hc_destino}. Tentando selecionar modalidade '{valor}'...")
    # Estratégia 1: CB_SETCURSEL por índice se 'valor' for número puro de índice —
    # mas aqui 'valor' é o número da modalidade (9), que não é o índice.
    # Como TComboBoxPesqRef é custom, o mais confiável costuma ser abrir o combo
    # e digitar. Aqui só TESTAMOS o count e a seleção atual; a escrita real fica
    # para o bot depois de sabermos se CB_* responde.
    n = win32gui.SendMessage(hc_destino, CB_GETCOUNT, 0, 0)
    atual = win32gui.SendMessage(hc_destino, CB_GETCURSEL, 0, 0)
    print(f"  count={n} cursel_atual={atual} texto='{_txt(hc_destino)}'")
    print("  (Teste apenas leu o combo — a seleção real será definida após confirmarmos o método.)")


def _achar_edits_busca(hwnd):
    """Retorna os TEdit 'Busca:' de cada painel, ordenados por x (esq, dir)."""
    edits = []
    def _cb(h, _):
        if _cls(h) == "TEdit":
            l, t, r, b = _rect(h)
            edits.append((h, l))
        return True
    try:
        win32gui.EnumChildWindows(hwnd, _cb, None)
    except Exception:
        pass
    edits.sort(key=lambda x: x[1])
    return edits


def cmd_testar_busca(hwnd, texto, lado):
    """
    TESTE NÃO-DESTRUTIVO: foca o campo Busca do painel indicado e digita `texto`.
    Serve para ver se a lista filtra. Nada é salvo. Requer sessão ELEVADA
    (SEFIP roda como Admin) para o envio de teclas funcionar.
    """
    from pywinauto.keyboard import send_keys

    edits = _achar_edits_busca(hwnd)
    if len(edits) < 2:
        print(f"❌ Esperava 2 campos Busca, achei {len(edits)}.")
        return
    idx = 1 if lado == "direita" else 0
    hedit = edits[idx][0]
    print(f"Focando Busca do painel {lado} (hwnd={hedit}) e digitando '{texto}'...")

    # Traz a janela Modalidades pra frente e foca o edit
    try:
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.3)
    except Exception as e:
        print(f"⚠️ SetForegroundWindow falhou: {e} (precisa rodar como Admin?)")

    try:
        win32gui.SendMessage(hedit, win32con.WM_SETFOCUS, 0, 0)
    except Exception:
        pass
    # Clique real no campo via mouse (mais confiável p/ VCL)
    l, t, r, b = _rect(hedit)
    cx, cy = (l + r) // 2, (t + b) // 2
    try:
        win32api.SetCursorPos((cx, cy))
        time.sleep(0.2)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        time.sleep(0.3)
    except Exception as e:
        print(f"⚠️ clique no campo falhou: {e}")

    try:
        send_keys("^a{BACKSPACE}")  # limpa
        time.sleep(0.2)
        send_keys(texto, with_spaces=True)
        print("✅ Texto digitado. Observe se a lista filtrou. (Nada foi salvo.)")
        print("   Se filtrou: o método 'Busca + teclado' funciona. Depois teste")
        print("   dar foco na lista e navegar com setas.")
    except Exception as e:
        print(f"❌ Falha ao digitar: {e}")


def cmd_clicar(hwnd, alvo):
    ctrls = _coletar_controles(hwnd)
    alvo_norm = alvo.strip().lower().replace("&", "")
    candidatos = []
    for hb, x, t in ctrls["bitbtns"]:
        t_norm = t.strip().lower().replace("&", "")
        if t_norm == alvo_norm:
            candidatos.append((hb, t))
    if not candidatos:
        print(f"❌ Nenhum botão com texto '{alvo}'. Botões disponíveis:")
        for hb, x, t in ctrls["bitbtns"]:
            print(f"    '{t}'")
        return
    hb, t = candidatos[0]
    print(f"▶ Clicando botão '{t}' (hwnd={hb}) via BM_CLICK...")
    win32gui.SendMessage(hb, BM_CLICK, 0, 0)
    print("✅ BM_CLICK enviado. Verifique a tela.")


def main():
    p = argparse.ArgumentParser(description="Teste da janela Modalidades do SEFIP")
    p.add_argument("--ler", action="store_true", help="Só leitura: lista itens/combos/botões (padrão)")
    p.add_argument("--achar-func", dest="achar_func", help="Procura um funcionário por nome/CPF nas listas")
    p.add_argument("--selecionar-mod", dest="selecionar_mod", help="[teste] lê o combo destino p/ modalidade")
    p.add_argument("--clicar", help="Clica um botão por texto: <<, <, >, >>, salvar, cancelar")
    p.add_argument("--testar-busca", dest="testar_busca", help="[teste não-destrutivo] digita texto no campo Busca")
    p.add_argument("--lado", default="direita", choices=["esquerda", "direita"], help="Painel do campo Busca (padrão: direita)")
    args = p.parse_args()

    hwnd = _achar_modalidades()
    if not hwnd:
        print("❌ Janela 'Modalidades' (TfrmAdmModalidades) não encontrada.")
        print("   Abra a tela via Ctrl+M no SEFIP antes de rodar este teste.")
        sys.exit(1)

    _, pid = win32process.GetWindowThreadProcessId(hwnd)

    if args.achar_func:
        cmd_achar_func(hwnd, pid, args.achar_func)
    elif args.selecionar_mod:
        cmd_selecionar_mod(hwnd, pid, args.selecionar_mod)
    elif args.testar_busca:
        cmd_testar_busca(hwnd, args.testar_busca, args.lado)
    elif args.clicar:
        cmd_clicar(hwnd, args.clicar)
    else:
        cmd_ler(hwnd, pid)


if __name__ == "__main__":
    main()
