"""
Vigia do CriptoCNS
==================
Trata a janela NATIVA do Windows que o CriptoCNS (assinador da Caixa) abre
durante o uso da Conectividade Social:

    ┌─ CriptoCNS ──────────────────────────────────────────────┐
    │ Atenção! O aplicativo no endereço                        │
    │ https://conectividadesocialv2.caixa.gov.br enviou        │
    │ solicitação para enumerar certificados de assinatura     │
    │ instalados. Você concorda em realizar essa operação?     │
    │   [Aceitar]  [Recusar]     ☐ Não perguntar novamente     │
    └──────────────────────────────────────────────────────────┘

POR QUE ISSO EXISTE
-------------------
Essa janela NÃO está no DOM. Não é mat-dialog, não é cdk-overlay, não é
alert() do navegador. É um processo Win32 separado. Playwright e a extensão
do Chrome enxergam apenas o HTML da página, então nenhum dos dois consegue
"achar o handle" — o bot fica esperando para sempre um botão que, para ele,
não existe. O sintoma típico é o "Enviar" ficar desabilitado e nada mais
acontecer: a página está bloqueada aguardando a resposta do CriptoCNS.

A solução é sair do mundo do navegador e tratar a janela por Win32 — a mesma
técnica que o Bot_Sefip.py usa com os popups VCL do SEFIP.

COMO USAR
---------
    from vigia_criptocns import VigiaCriptoCNS

    with VigiaCriptoCNS(log=print) as vigia:
        ...  # fluxo Playwright normal
    print(vigia.tratadas)   # quantas janelas foram aceitas

O vigia roda em thread daemon e varre as janelas a cada `intervalo` segundos.
Precisa ser assim porque o clique em "Enviar" BLOQUEIA esperando a resposta
do CriptoCNS: se o tratamento fosse sequencial, o bot nunca chegaria nele.

Também dá para usar direto na linha de comando, para testar:

    python vigia_criptocns.py --vigiar 120

Autor: Hugo L. Almeida
"""

import argparse
import sys
import threading
import time

try:
    import win32api
    import win32con
    import win32gui
    WIN32 = True
except ImportError:
    WIN32 = False

# Constantes Win32 (mesmas do Bot_Sefip.py)
BM_CLICK = 0x00F5
BM_SETCHECK = 0x00F1
BM_GETCHECK = 0x00F0
BST_CHECKED = 1

# Identificação da janela. O título é "CriptoCNS"; o corpo cita o domínio da
# Conectividade. Conferir os dois evita clicar em janela homônima de outro
# contexto.
TITULO_CONTEM = "criptocns"
CORPO_CONTEM = ("certificados", "conectividadesocial")

# O CriptoCNS é Electron/CEF: a janela é 'Chrome_WidgetWin_1' e os únicos
# filhos são a superfície de render e a janela D3D (confirmado no dump de
# 2026-07-29, hwnd 19271670, exe criptocns.exe). Os botões "Aceitar"/"Recusar"
# são HTML DENTRO do processo — não existem como HWND. Consequência: BM_CLICK,
# busca por texto de filho e UI Automation não funcionam; resta clique real
# por coordenada.
CLASSE_ELECTRON = "chrome_widgetwin"
EXE_ESPERADO = "criptocns.exe"

# Posição do botão "Aceitar" como FRAÇÃO do retângulo da janela — nunca em
# pixels absolutos, para sobreviver a mudança de posição/resolução.
# Aferido em DUAS capturas independentes do diálogo real (2026-07-29):
#   captura A: janela 587x240 -> botão x=[0.055,0.216] y=[0.717,0.808]
#   captura B: janela 430x176 -> botão x=[0.072,0.205] y=[0.693,0.812]
# Centro convergente ~ (0.136, 0.758). A margem é folgada: o botão ocupa
# ~15% da largura, então um erro de alguns pontos percentuais ainda acerta.
FRACAO_ACEITAR = (0.136, 0.758)

# Caixa do botão em frações, usada nos testes de regressão para garantir que
# o ponto calculado cai DENTRO do botão (e não no "Recusar", à direita).
CAIXA_ACEITAR = (0.055, 0.693, 0.216, 0.812)

# Dimensões esperadas do diálogo. Se a janela vier com tamanho muito
# diferente, o layout mudou e o clique proporcional deixa de ser confiável —
# nesse caso o vigia se RECUSA a clicar e avisa, em vez de clicar às cegas.
TAMANHO_ESPERADO = (600, 250)
TOLERANCIA_TAMANHO = 0.25

TEXTO_ACEITAR = "aceitar"
TEXTO_NAO_PERGUNTAR = "não perguntar novamente"


def _norm(s):
    return (s or "").strip().lower().replace("&", "")


def _cls(h):
    try:
        return win32gui.GetClassName(h) or ""
    except Exception:
        return ""


def _txt(h):
    try:
        return win32gui.GetWindowText(h) or ""
    except Exception:
        return ""


def _filhos(hwnd):
    achados = []

    def _cb(h, _):
        achados.append(h)
        return True

    try:
        win32gui.EnumChildWindows(hwnd, _cb, None)
    except Exception:
        pass
    return achados


def _janelas_topo():
    achados = []

    def _cb(h, _):
        try:
            if win32gui.IsWindowVisible(h):
                achados.append(h)
        except Exception:
            pass
        return True

    try:
        win32gui.EnumWindows(_cb, None)
    except Exception:
        pass
    return achados


def eh_janela_criptocns(titulo, classe="", exe="", textos_filhos=()):
    """
    Decide se uma janela é o diálogo de autorização do CriptoCNS.

    Regra em três camadas, da mais forte para a mais fraca:
      1. processo == criptocns.exe  (definitivo — ninguém mais o usa);
      2. título contém "criptocns";
      3. se por acaso houver filhos com texto (versão futura em Win32
         clássico), exige-se também menção a certificados/domínio.

    Exige título E (processo OU classe Electron). Só o título não basta:
    'Serviço CriptoCNS' é OUTRA janela do mesmo exe (invisível, hwnd 66990 no
    dump) e não é o diálogo de autorização.

    Separado da API do Windows para poder ser testado sem Windows — é aqui
    que mora o risco de clicar na janela errada.
    """
    t = _norm(titulo)
    if TITULO_CONTEM not in t:
        return False
    # A janela de serviço não é o diálogo — descarta explicitamente.
    if t.startswith("servico") or t.startswith("serviço"):
        return False
    pelo_processo = _norm(exe) == EXE_ESPERADO
    pela_classe = CLASSE_ELECTRON in _norm(classe)
    if not (pelo_processo or pela_classe):
        return False
    textos = [x for x in textos_filhos if _norm(x)]
    # Se algum filho tiver texto real (não é o caso do Electron), aproveita
    # como reforço; se não tiver, o processo/classe já bastam.
    reais = [x for x in textos if _norm(x) not in ("chrome legacy window",)]
    if reais:
        corpo = " ".join(_norm(x) for x in reais)
        if not any(m in corpo for m in CORPO_CONTEM):
            return False
    return True


def coordenada_aceitar(rect, fracao=FRACAO_ACEITAR):
    """Converte o retângulo da janela na coordenada absoluta do 'Aceitar'."""
    esq, topo, dir_, base = rect
    largura, altura = dir_ - esq, base - topo
    return (int(esq + largura * fracao[0]), int(topo + altura * fracao[1]))


def tamanho_plausivel(rect, esperado=TAMANHO_ESPERADO, tol=TOLERANCIA_TAMANHO):
    """
    O diálogo tem tamanho fixo. Se vier muito diferente, o layout mudou e a
    fração deixa de apontar para o botão — melhor não clicar do que clicar
    em 'Recusar' ou fora da janela.
    """
    largura, altura = rect[2] - rect[0], rect[3] - rect[1]
    if largura <= 0 or altura <= 0:
        return False
    return (abs(largura - esperado[0]) <= esperado[0] * tol and
            abs(altura - esperado[1]) <= esperado[1] * tol)


def _achar_dialogos():
    """Retorna [(hwnd, filhos)] das janelas de autorização do CriptoCNS."""
    encontrados = []
    for h in _janelas_topo():
        titulo = _txt(h)
        if TITULO_CONTEM not in _norm(titulo):
            continue
        _, exe = _info_processo(h)
        filhos = _filhos(h)
        if eh_janela_criptocns(titulo, _cls(h), exe or "",
                               [_txt(f) for f in filhos]):
            encontrados.append((h, filhos))
    return encontrados


def _info_processo(hwnd):
    """(pid, nome do executável) da janela — identifica quem realmente a criou."""
    try:
        import win32api
        import win32process
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        # PROCESS_QUERY_INFORMATION | PROCESS_VM_READ
        h = win32api.OpenProcess(0x0400 | 0x0010, False, pid)
        try:
            exe = win32process.GetModuleFileNameEx(h, 0)
        finally:
            win32api.CloseHandle(h)
        import os
        return pid, os.path.basename(exe)
    except Exception:
        try:
            import win32process
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            return pid, "(sem permissão para ler o executável)"
        except Exception:
            return None, None


def dump_janelas(filtro=None, incluir_invisiveis=False, max_filhos=25):
    """
    Despejo CRU de todas as janelas de topo — sem nenhum filtro de conteúdo.

    Existe porque `--listar` aplica a regra de identificação inteira e, quando
    ela falha, não mostra nada: o diagnóstico fica cego. Aqui vemos classe,
    título, processo e filhos REAIS, que é o que diz se o app é Win32 clássico
    (filhos com texto) ou desenhado (Qt/WPF/Electron — poucos ou nenhum filho).
    """
    linhas = []

    def _cb(h, _):
        try:
            visivel = win32gui.IsWindowVisible(h)
            if not visivel and not incluir_invisiveis:
                return True
            titulo = _txt(h)
            classe = _cls(h)
            if filtro:
                alvo = _norm(filtro)
                if alvo not in _norm(titulo) and alvo not in _norm(classe):
                    return True
            elif not titulo:
                return True
            pid, exe = _info_processo(h)
            filhos = _filhos(h)
            linhas.append({
                "hwnd": h, "titulo": titulo, "classe": classe,
                "visivel": visivel, "pid": pid, "exe": exe,
                "rect": _rect(h),
                "filhos": [(_cls(f), _txt(f), _rect(f)) for f in filhos[:max_filhos]],
                "total_filhos": len(filhos),
            })
        except Exception:
            pass
        return True

    try:
        win32gui.EnumWindows(_cb, None)
    except Exception:
        pass
    return linhas


def _rect(h):
    try:
        return win32gui.GetWindowRect(h)
    except Exception:
        return (0, 0, 0, 0)


def _achar_filho_por_texto(filhos, alvo):
    alvo = _norm(alvo)
    for f in filhos:
        if _norm(_txt(f)) == alvo:
            return f
    for f in filhos:
        if alvo in _norm(_txt(f)):
            return f
    return 0


def _clicar_coordenada(x, y, log=print):
    """
    Clique REAL do mouse (SetCursorPos + mouse_event).

    Necessário porque o alvo é HTML dentro do Electron: não há HWND para
    receber BM_CLICK. Mesma técnica que o Bot_Sefip.py usa nos campos VCL,
    pelo mesmo motivo (o controle não responde a mensagens sintéticas).

    Efeito colateral assumido: move o cursor do usuário. Guardamos e
    restauramos a posição original para incomodar o mínimo possível.
    """
    try:
        pos_original = win32api.GetCursorPos()
    except Exception:
        pos_original = None
    try:
        win32api.SetCursorPos((x, y))
        time.sleep(0.12)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        time.sleep(0.06)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        time.sleep(0.12)
        return True
    except Exception as e:
        log(f"   ❌ Falha no clique real: {type(e).__name__}: {e}")
        return False
    finally:
        if pos_original:
            try:
                win32api.SetCursorPos(pos_original)
            except Exception:
                pass


def tratar_dialogo_electron(hwnd, log=print):
    """
    Clica "Aceitar" no diálogo Electron por coordenada proporcional.

    Antes de clicar confere o tamanho da janela: se o layout mudou, prefere
    NÃO clicar (clicar às cegas poderia acertar "Recusar", que nega o acesso
    aos certificados e derruba o envio).
    """
    rect = _rect(hwnd)
    if not tamanho_plausivel(rect):
        larg, alt = rect[2] - rect[0], rect[3] - rect[1]
        log(f"   ⚠️ Janela do CriptoCNS com tamanho inesperado ({larg}x{alt}, "
            f"esperado ~{TAMANHO_ESPERADO[0]}x{TAMANHO_ESPERADO[1]}). "
            f"NÃO vou clicar às cegas — aceite manualmente e me avise para "
            f"eu recalibrar.")
        return False

    try:
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            time.sleep(0.2)
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.25)
    except Exception:
        # Sem foco o clique ainda costuma funcionar (a janela é topmost),
        # então seguimos — só registramos.
        log("   ℹ️ Não consegui trazer a janela para frente; clicando assim mesmo.")

    x, y = coordenada_aceitar(rect)
    log(f"   🖱️ Clicando 'Aceitar' em ({x}, {y}) — janela {rect}.")
    if not _clicar_coordenada(x, y, log):
        return False

    # Confirma pelo desaparecimento da janela: é a única evidência real de
    # que o clique acertou o botão.
    for _ in range(12):
        time.sleep(0.25)
        if not win32gui.IsWindow(hwnd) or not win32gui.IsWindowVisible(hwnd):
            log("   ✅ Diálogo do CriptoCNS aceito (janela fechou).")
            return True
    log("   ⚠️ Cliquei, mas a janela continua aberta — a fração pode estar "
        "desalinhada. Aceite manualmente; rode --dump cripto para eu recalibrar.")
    return False


def tratar_dialogo(hwnd, filhos, marcar_nao_perguntar=False, log=print):
    """
    Marca (opcionalmente) "Não perguntar novamente" e clica em "Aceitar".

    A ordem importa: o checkbox precisa ser marcado ANTES do Aceitar, senão
    a preferência não é gravada e a janela volta na próxima operação.
    """
    if marcar_nao_perguntar:
        chk = _achar_filho_por_texto(filhos, TEXTO_NAO_PERGUNTAR)
        if chk:
            try:
                if win32gui.SendMessage(chk, BM_GETCHECK, 0, 0) != BST_CHECKED:
                    win32gui.SendMessage(chk, BM_SETCHECK, BST_CHECKED, 0)
                    win32gui.SendMessage(chk, BM_CLICK, 0, 0)
                    log("   ☑️ 'Não perguntar novamente' marcado.")
            except Exception as e:
                log(f"   ⚠️ Não consegui marcar o checkbox: {type(e).__name__}")

    botao = _achar_filho_por_texto(filhos, TEXTO_ACEITAR)
    if not botao:
        # Caso normal hoje: Electron, sem botão como HWND. Vai de clique real.
        return tratar_dialogo_electron(hwnd, log=log)

    try:
        # PostMessage (assíncrono): o clique dispara trabalho no processo do
        # CriptoCNS; SendMessage bloquearia esta thread até ele responder.
        win32api.PostMessage(botao, BM_CLICK, 0, 0)
        log("   ✅ 'Aceitar' clicado no CriptoCNS.")
        return True
    except Exception as e:
        log(f"   ❌ Falha ao clicar 'Aceitar': {type(e).__name__}: {e}")
        return False


class VigiaCriptoCNS:
    """
    Vigia em thread daemon. Enquanto ativo, aceita automaticamente os
    diálogos do CriptoCNS que aparecerem.

    Fora do Windows (ou sem pywin32) vira no-op silencioso, para não quebrar
    testes nem execução em outro ambiente.
    """

    def __init__(self, intervalo=0.7, marcar_nao_perguntar=False, log=print,
                 ativo=True):
        self.intervalo = intervalo
        self.marcar_nao_perguntar = marcar_nao_perguntar
        self.log = log
        self.tratadas = 0
        self._parar = threading.Event()
        self._thread = None
        self.disponivel = WIN32 and ativo

    def _loop(self):
        vistos = set()
        while not self._parar.is_set():
            try:
                for hwnd, filhos in _achar_dialogos():
                    if hwnd in vistos:
                        continue
                    self.log("🔐 Janela do CriptoCNS detectada "
                             "(pedido de acesso aos certificados).")
                    if tratar_dialogo(hwnd, filhos,
                                      self.marcar_nao_perguntar, self.log):
                        self.tratadas += 1
                        vistos.add(hwnd)
                # hwnds fechados saem do conjunto, permitindo tratar de novo
                # se o Windows reciclar o handle numa janela nova.
                vistos = {h for h in vistos if win32gui.IsWindow(h)}
            except Exception:
                pass
            self._parar.wait(self.intervalo)

    def iniciar(self):
        if not self.disponivel:
            self.log("ℹ️ Vigia do CriptoCNS inativo (fora do Windows ou sem "
                     "pywin32) — se a janela aparecer, será preciso clicar "
                     "'Aceitar' manualmente.")
            return self
        self._thread = threading.Thread(target=self._loop, daemon=True,
                                        name="vigia-criptocns")
        self._thread.start()
        self.log("👁️ Vigia do CriptoCNS ativo.")
        return self

    def parar(self):
        self._parar.set()
        if self._thread:
            self._thread.join(timeout=3)
        if self.tratadas:
            self.log(f"🔐 CriptoCNS: {self.tratadas} janela(s) aceita(s) "
                     f"automaticamente.")

    def __enter__(self):
        return self.iniciar()

    def __exit__(self, *exc):
        self.parar()
        return False


def main():
    p = argparse.ArgumentParser(description="Vigia do CriptoCNS")
    p.add_argument("--vigiar", type=int, metavar="SEGUNDOS", default=60,
                   help="Fica vigiando por N segundos (padrão: 60)")
    p.add_argument("--nao-perguntar-novamente", action="store_true",
                   help="Marca a caixa 'Não perguntar novamente' antes de aceitar")
    p.add_argument("--listar", action="store_true",
                   help="Só lista as janelas do CriptoCNS visíveis agora, sem clicar")
    p.add_argument("--dump", nargs="?", const="", metavar="FILTRO",
                   help="Despejo CRU de todas as janelas de topo (classe, título, "
                        "processo, filhos). Opcionalmente filtra por texto em "
                        "título/classe. Use quando --listar não achar nada.")
    p.add_argument("--incluir-invisiveis", action="store_true",
                   help="No --dump, inclui também janelas não visíveis")
    args = p.parse_args()

    if not WIN32:
        print("❌ Precisa de Windows + pywin32.")
        sys.exit(1)

    if args.dump is not None:
        achados = dump_janelas(filtro=args.dump or None,
                               incluir_invisiveis=args.incluir_invisiveis)
        alvo = f" contendo {args.dump!r}" if args.dump else ""
        print(f"{len(achados)} janela(s){alvo}:\n")
        for j in achados:
            vis = "" if j["visivel"] else "  [INVISÍVEL]"
            print(f"hwnd={j['hwnd']}  classe={j['classe']!r}{vis}")
            print(f"   título : {j['titulo']!r}")
            print(f"   processo: pid={j['pid']} exe={j['exe']!r}")
            print(f"   rect   : {j['rect']}   filhos: {j['total_filhos']}")
            for c, t, r in j["filhos"]:
                if t or c:
                    print(f"      {c:22} {t[:55]!r}  {r}")
            print()
        if not achados:
            print("Nada encontrado. Se a janela está na tela, tente:")
            print("   python vigia_criptocns.py --dump --incluir-invisiveis")
        sys.exit(0)

    if args.listar:
        achados = _achar_dialogos()
        if not achados:
            print("Nenhuma janela do CriptoCNS visível no momento.")
        for hwnd, filhos in achados:
            print(f"\nhwnd={hwnd} título={_txt(hwnd)!r} classe={_cls(hwnd)}")
            for f in filhos:
                t = _txt(f)
                if t:
                    print(f"   {_cls(f):18} {t[:70]!r}")
        sys.exit(0)

    print(f"Vigiando por {args.vigiar}s... (deixe o navegador disparar a operação)")
    with VigiaCriptoCNS(marcar_nao_perguntar=args.nao_perguntar_novamente) as v:
        time.sleep(args.vigiar)
    print(f"Fim. Janelas tratadas: {v.tratadas}")


if __name__ == "__main__":
    main()
