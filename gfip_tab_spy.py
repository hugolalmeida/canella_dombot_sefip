"""
GFIP Tab Spy
============
Ferramenta de diagnóstico: mostra qual controle está com foco dentro da
janela do GFIP no Domínio Folha, a cada instante, para descobrir a ordem
exata de TAB entre os campos sem precisar adivinhar.

Como usar:
1. Abra manualmente a janela do GFIP no Domínio (Relatórios > Informativos
   > ... > GFIP), deixando-a em primeiro plano.
2. Rode este script: python gfip_tab_spy.py
3. Ele localiza a janela do GFIP e entra em loop, imprimindo o controle
   focado a cada ~0.4s. Clique no campo "Competência" para começar do
   início e vá apertando TAB manualmente — a cada TAB, a próxima linha
   impressa mostra em qual campo você caiu.
4. Ctrl+C para parar.

Cada linha impressa mostra: ctrl_id, classe, texto atual do campo (se
houver) e o texto do rótulo mais próximo (heurística por posição), o que
ajuda a identificar o campo mesmo sem saber o nome exato do controle.
"""

import time
import sys
import win32gui
import win32api
import win32process
import ctypes
from ctypes import wintypes

user32 = ctypes.windll.user32


def _find_gfip_window() -> int:
    """
    Localiza a janela do GFIP. Ela NÃO é top-level — abre como aba/filho
    embutido (classe FNWND3190) dentro da janela principal do Domínio Folha,
    então primeiro achamos o Domínio e depois procuramos o filho GFIP nele.
    """
    dominio_hwnd = [0]

    def cb_dominio(hwnd, _):
        if dominio_hwnd[0]:
            return
        try:
            if not win32gui.IsWindowVisible(hwnd):
                return
            titulo = win32gui.GetWindowText(hwnd)
            if "domínio" in titulo.lower() or "dominio" in titulo.lower():
                dominio_hwnd[0] = hwnd
        except Exception:
            pass

    win32gui.EnumWindows(cb_dominio, None)
    if not dominio_hwnd[0]:
        return 0

    result = [0]

    def cb_filho(hwnd, _):
        if result[0]:
            return
        try:
            if not win32gui.IsWindowVisible(hwnd):
                return
            titulo = win32gui.GetWindowText(hwnd)
            if "gfip" in titulo.lower():
                result[0] = hwnd
        except Exception:
            pass

    win32gui.EnumChildWindows(dominio_hwnd[0], cb_filho, None)
    if result[0]:
        return result[0]

    # Fallback: também tenta top-level (caso a versão do Domínio mude)
    def cb(hwnd, _):
        if result[0]:
            return
        try:
            if not win32gui.IsWindowVisible(hwnd):
                return
            titulo = win32gui.GetWindowText(hwnd)
            if "gfip" in titulo.lower():
                result[0] = hwnd
        except Exception:
            pass

    win32gui.EnumWindows(cb, None)
    return result[0]


def _get_focused_control_info():
    """
    Retorna (hwnd, ctrl_id, classe, texto) do controle com foco de teclado,
    usando GetGUIThreadInfo (funciona entre processos, diferente de
    GetFocus() puro que só funciona na thread chamadora).
    """
    hwnd_fg = win32gui.GetForegroundWindow()
    if not hwnd_fg:
        return None

    _, pid = win32process.GetWindowThreadProcessId(hwnd_fg)
    tid = win32process.GetWindowThreadProcessId(hwnd_fg)[0]

    class GUITHREADINFO(ctypes.Structure):
        _fields_ = [
            ("cbSize", wintypes.DWORD),
            ("flags", wintypes.DWORD),
            ("hwndActive", wintypes.HWND),
            ("hwndFocus", wintypes.HWND),
            ("hwndCapture", wintypes.HWND),
            ("hwndMenuOwner", wintypes.HWND),
            ("hwndMoveSize", wintypes.HWND),
            ("hwndCaret", wintypes.HWND),
            ("rcCaret", wintypes.RECT),
        ]

    info = GUITHREADINFO()
    info.cbSize = ctypes.sizeof(GUITHREADINFO)
    ok = user32.GetGUIThreadInfo(tid, ctypes.byref(info))
    if not ok or not info.hwndFocus:
        return None

    hwnd_focus = info.hwndFocus
    try:
        ctrl_id = win32gui.GetDlgCtrlID(hwnd_focus)
    except Exception:
        ctrl_id = None
    try:
        classe = win32gui.GetClassName(hwnd_focus)
    except Exception:
        classe = "?"
    try:
        texto = win32gui.GetWindowText(hwnd_focus)
    except Exception:
        texto = ""

    return (hwnd_focus, ctrl_id, classe, texto)


def _find_nearest_label(gfip_hwnd, target_hwnd):
    """
    Heurística simples: procura entre os irmãos 'Static' (labels) do mesmo
    diálogo aquele cujo retângulo está mais próximo (acima/à esquerda) do
    controle focado, para sugerir a que campo ele pertence.
    """
    try:
        rect_target = win32gui.GetWindowRect(target_hwnd)
    except Exception:
        return ""

    best_label = ""
    best_dist = None

    def cb(hwnd, _):
        nonlocal best_label, best_dist
        try:
            if win32gui.GetClassName(hwnd) != "Static":
                return True
            texto = win32gui.GetWindowText(hwnd).strip()
            if not texto:
                return True
            rect = win32gui.GetWindowRect(hwnd)
            # distância simples: soma de diffs de topo/esquerda
            dist = abs(rect[1] - rect_target[1]) + max(0, rect_target[0] - rect[0])
            if best_dist is None or dist < best_dist:
                best_dist = dist
                best_label = texto
        except Exception:
            pass
        return True

    try:
        win32gui.EnumChildWindows(gfip_hwnd, cb, None)
    except Exception:
        pass

    return best_label


def main():
    print("Procurando janela do GFIP (título contendo 'GFIP')...")
    gfip_hwnd = 0
    for _ in range(20):
        gfip_hwnd = _find_gfip_window()
        if gfip_hwnd:
            break
        time.sleep(0.5)

    if not gfip_hwnd:
        print("Janela do GFIP não encontrada. Abra-a manualmente no Domínio e rode de novo.")
        sys.exit(1)

    titulo = win32gui.GetWindowText(gfip_hwnd)
    print(f"Janela encontrada: '{titulo}' (hwnd={gfip_hwnd})")
    print("Clique no primeiro campo (Competência) e vá apertando TAB manualmente.")
    print("Ctrl+C para parar.\n")

    last_key = None
    try:
        while True:
            info = _get_focused_control_info()
            if info:
                hwnd_focus, ctrl_id, classe, texto = info
                label = _find_nearest_label(gfip_hwnd, hwnd_focus)
                key = (hwnd_focus, ctrl_id, classe, texto, label)
                if key != last_key:
                    print(f"FOCO -> ctrl_id={ctrl_id!r:>6}  classe={classe:<20}  "
                          f"texto={texto!r:<30}  rotulo_proximo={label!r}")
                    last_key = key
            time.sleep(0.15)
    except KeyboardInterrupt:
        print("\nEncerrado.")


if __name__ == "__main__":
    main()
