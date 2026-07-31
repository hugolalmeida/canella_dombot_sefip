"""
Bot SEFIP — Recálculo de FGTS
=============================
Automação RPA do SEFIP para o processo de recálculo de FGTS.

Fluxo manual das 6 etapas:
  1. Importar (TXT gerado pelo SEFIP) — o usuário seleciona o arquivo; o bot
     confirma o aviso de sobreposição ("todos os dados serão excluídos") com Sim.
  2. Marcar Participação — na tela Movimento, abrir Modalidades (Ctrl+M).
  3. Separar o funcionário do ajuste — na tela Modalidades: escolher modalidade
     9 (destino), mandar todos p/ direita (>>), buscar o funcionário pelo
     nome/CPF, devolvê-lo p/ esquerda (<) e Salvar. Assim o recálculo do FGTS
     ocorre só para esse funcionário.
  4. Gerar arquivo no SEFIP        — (ainda não implementado)
  5. Validar no Conectividade Social (web/externo — fora do escopo win32)
  6. Gerar Guia no SEFIP           — (ainda não implementado)

DESCOBERTAS TÉCNICAS (ver memory/sefip-automation.md):
  - SEFIP é Delphi/VCL (TfrmPrincipalSEFIP). NÃO expõe UI Automation → o
    Inspect falha (0x80070005). Mas Win32 puro lê/controla tudo — mesma técnica
    do DomBot_GFIP.py.
  - ctrl_ids VCL são INSTÁVEIS (mudam a cada recriação da janela). Localizamos
    controles por CLASSE + TEXTO + POSIÇÃO, nunca por ctrl_id fixo.
  - Os TListView de funcionários NÃO são common controls nativos (LVM_* não
    funciona). Estratégia validada: filtrar pelo campo "Busca" + selecionar por
    foco/teclado; transferir com os botões <</</>/>> via BM_CLICK.
  - Foco em campos VCL: clique de mouse REAL (SetCursorPos + mouse_event) é mais
    confiável que WM_SETFOCUS.
  - O SEFIP roda ELEVADO → este bot precisa rodar como Administrador, senão a
    UIPI bloqueia o envio de teclas/cliques.

Autor: Hugo L. Almeida
Versão: 0.1 (Etapas 1-3)
"""

import argparse
import ctypes
import os
import sys
import time
import traceback
from datetime import datetime

import win32api
import win32con
import win32gui
from pywinauto.keyboard import send_keys

try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None


def _esta_elevado() -> bool:
    """
    True se o processo atual roda com privilégios de Administrador.
    O SEFIP roda elevado; sem isso a UIPI bloqueia silenciosamente
    SetCursorPos/mouse_event/envio de teclas (falha sem exceção clara —
    o sintoma costuma ser "clicou no controle errado" ou "nada aconteceu").
    """
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


# ── Constantes Win32 ─────────────────────────────────────────────────────────

BM_CLICK = 0x00F5
WM_CLOSE = 0x0010


# ── Logger simples (console) — a GUI virá depois ────────────────────────────

class ConsoleLogger:
    def _log(self, nivel, msg):
        ts = datetime.now().strftime("%H:%M:%S")
        print(f"[{ts}] {nivel:5s} {msg}", flush=True)

    def info(self, msg):    self._log("INFO", msg)
    def warning(self, msg): self._log("AVISO", msg)
    def error(self, msg):   self._log("ERRO", msg)


# ── Helpers de janela genéricos ─────────────────────────────────────────────

def _cls(h):
    try: return win32gui.GetClassName(h) or ""
    except Exception: return ""

def _txt(h):
    try: return win32gui.GetWindowText(h) or ""
    except Exception: return ""

def _rect(h):
    try: return win32gui.GetWindowRect(h)
    except Exception: return (0, 0, 0, 0)

def _visivel(h):
    try: return win32gui.IsWindowVisible(h)
    except Exception: return False


def _enum_toplevel(pred):
    """Retorna lista de hwnds top-level visíveis que satisfazem pred(hwnd)."""
    achados = []
    def _cb(h, _):
        try:
            if _visivel(h) and pred(h):
                achados.append(h)
        except Exception:
            pass
        return True
    try:
        win32gui.EnumWindows(_cb, None)
    except Exception:
        pass
    return achados


def _enum_filhos(hwnd, pred=None):
    """Retorna lista de hwnds descendentes; se pred, filtra por pred(hwnd)."""
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


# ── Engine de automação do SEFIP ─────────────────────────────────────────────

class SefipAutomation:
    """Automação do recálculo de FGTS no SEFIP (Etapas 1-3)."""

    def __init__(self, logger=None):
        self.log = logger or ConsoleLogger()
        self.main_hwnd = 0
        self.mensagem_inconsistencia = None

    # ── Conexão ──────────────────────────────────────────────────────────────

    def conectar(self) -> bool:
        """
        Localiza a janela principal do SEFIP (TfrmPrincipalSEFIP).

        🔴 NÃO filtra por _visivel: quando o SEFIP está minimizado, a Tfrm
        principal fica com IsWindowVisible=False (e mesmo assim NÃO se
        declara IsIconic=True — quem carrega esse estado é a janela-sombra
        do processo, classe TApplication, título também "Sefip", owner da
        Tfrm via GetWindow(GW_OWNER)). Filtrar por visível faz sobrar só a
        TApplication, que casa no critério fraco de título mas NÃO tem o
        menu de verdade — testado na prática (Bot_Sefip_Protocolo.py,
        2026-07-28): rodar com o SEFIP minimizado conectava na janela
        errada. Prioriza sempre match por classe Tfrm...SEFIP (a principal
        de verdade) sobre o critério fraco de título.
        """
        self.log.info("🔎 Localizando SEFIP...")

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

        try:
            win32gui.EnumWindows(_cb, None)
        except Exception:
            pass
        janelas = por_classe or por_titulo

        if not janelas:
            self.log.error("❌ SEFIP não encontrado. Abra o sistema antes de iniciar.")
            return False

        self.main_hwnd = janelas[0]
        self.log.info(f"✅ SEFIP encontrado: '{_txt(self.main_hwnd)}' (hwnd={self.main_hwnd})")
        return True

    def _trazer_frente(self, hwnd):
        """
        Restaura e foca a janela. Também restaura a OWNER (GetWindow
        GW_OWNER) se ela estiver minimizada — no SEFIP, é a janela-sombra
        TApplication (owner da Tfrm principal) que carrega o estado
        IsIconic de verdade; restaurar só a Tfrm não traz nada de volta à
        tela (ver nota em conectar()).
        """
        try:
            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                time.sleep(0.3)
            owner = win32gui.GetWindow(hwnd, win32con.GW_OWNER)
            if owner and win32gui.IsIconic(owner):
                win32gui.ShowWindow(owner, win32con.SW_RESTORE)
                time.sleep(0.3)
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.3)
        except Exception as e:
            self.log.warning(f"⚠️ SetForegroundWindow falhou (rodando como Admin?): {e}")

    # ── Clique/foco de controles ─────────────────────────────────────────────

    def _clicar_botao_por_texto(self, hwnd_pai, texto_alvo, classes=("TBitBtn", "TButton"),
                                assincrono=False) -> bool:
        """
        Localiza um botão descendente por texto (ignora &) e clica.

        `assincrono=False` (padrão) usa SendMessage — seguro para botões que
        respondem rápido (>>, <, >, <<). `assincrono=True` usa PostMessage —
        OBRIGATÓRIO para botões que disparam processamento bloqueante no
        SEFIP (ex.: 'Salvar' da tela Modalidades, que só retorna depois do
        popup de sucesso ser fechado — testado: SendMessage aqui trava o
        processo Python por minutos até alguém clicar OK manualmente no
        popup, porque a thread de mensagens do SEFIP fica presa dentro do
        próprio clique).
        """
        alvo = _norm(texto_alvo)
        for h in _enum_filhos(hwnd_pai):
            if _cls(h) in classes and _norm(_txt(h)) == alvo:
                self.log.info(f"▶ Clicando '{_txt(h)}' (hwnd={h})")
                try:
                    if assincrono:
                        win32gui.PostMessage(h, BM_CLICK, 0, 0)
                    else:
                        win32gui.SendMessage(h, BM_CLICK, 0, 0)
                    return True
                except Exception as e:
                    self.log.error(f"❌ BM_CLICK falhou em '{texto_alvo}': {e}")
                    return False
        self.log.error(f"❌ Botão '{texto_alvo}' não encontrado em hwnd={hwnd_pai}")
        return False

    def _clicar_mouse(self, hwnd):
        """Clica com mouse real no centro de um controle (foco confiável em VCL)."""
        l, t, r, b = _rect(hwnd)
        cx, cy = (l + r) // 2, (t + b) // 2
        try:
            win32api.SetCursorPos((cx, cy))
            time.sleep(0.15)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
            time.sleep(0.2)
            return True
        except Exception as e:
            self.log.warning(f"⚠️ clique de mouse falhou: {e}")
            return False

    # ── Popups genéricos ─────────────────────────────────────────────────────

    def _achar_popup(self, titulos=None, corpo_contem=None, classes=("#32770", "TMessageForm")):
        """
        Retorna hwnd do primeiro popup visível cujo título OU corpo bate.
        titulos: lista de textos (match no título, case-insensitive, substring).
        corpo_contem: lista de textos a procurar no corpo (labels Static).
        """
        titulos = [t.lower() for t in (titulos or [])]
        corpo_contem = [c.lower() for c in (corpo_contem or [])]

        def _pred(h):
            c = _cls(h)
            if not any(k in c for k in classes) and c not in classes:
                return False
            t = _txt(h).lower()
            if titulos and any(k in t for k in titulos):
                return True
            if corpo_contem:
                corpo = " ".join(_txt(ch).lower() for ch in _enum_filhos(h))
                if any(k in corpo for k in corpo_contem):
                    return True
            if not titulos and not corpo_contem:
                return True
            return False

        achados = _enum_toplevel(_pred)
        return achados[0] if achados else 0

    def _clicar_no_popup(self, hpopup, texto_botao) -> bool:
        """Clica um botão (por texto) dentro de um popup."""
        alvo = _norm(texto_botao)
        for h in _enum_filhos(hpopup):
            if _cls(h) in ("TButton", "Button", "TBitBtn") and _norm(_txt(h)) == alvo:
                self.log.info(f"▶ Popup: clicando '{_txt(h)}'")
                try:
                    win32gui.SendMessage(h, BM_CLICK, 0, 0)
                    return True
                except Exception:
                    pass
        # Fallback: dá foco e envia o accelerator (ex.: Alt+S para Sim)
        try:
            self._trazer_frente(hpopup)
            send_keys("%" + alvo[0] if alvo else "{ENTER}")
            return True
        except Exception:
            return False

    # ══════════════════════════════════════════════════════════════════════
    # ETAPA 1 — Importação completa: menu → diálogo → sobreposição → avisos
    # ══════════════════════════════════════════════════════════════════════

    def importar_arquivo(self, caminho_arquivo: str, timeout=60) -> bool:
        """
        Executa a Etapa 1 inteira, do zero, sem qualquer ação manual prévia:
          1. Abre o menu Arquivo (Alt+A) e navega até "Importar Folha"
             (posição 6 do menu = Home + 5×DOWN + Enter — os itens são
             owner-drawn, sem texto legível via GetMenuString, então a
             navegação é por posição fixa, confirmada empiricamente).
          2. No diálogo nativo "Abrir Arquivo para Importação da Folha de
             Pagamento" (#32770), preenche o campo "Nome" (ctrl_id=1152) via
             WM_SETTEXT com o caminho completo e clica "&Abrir".
          3. Aguarda e confirma SIM no popup de sobreposição ("...Todos os
             dados serão excluídos. Confirma?").
          4. Fecha (OK) qualquer popup informativo subsequente — foram vistos
             até 2 em sequência: um 'Informação' (conclusão da importação) e,
             se a empresa/competência já tinha dados anteriores, um
             'Atenção' com OK único avisando que dados antigos foram
             substituídos (não confundir com o popup de INCONSISTÊNCIA de
             dados, que tem Sim/Não — esse é tratado à parte, ver nota).

        `caminho_arquivo`: caminho completo do .re/.txt gerado pelo GFIP.

        NOTA: se a importação detectar inconsistências nos dados (ex.:
        trabalhador já existe na base com outra modalidade — normal neste
        fluxo de recálculo), pode aparecer um popup 'Atenção' com Sim/Não
        perguntando se quer ver o relatório. Esse método responde Não
        (não é erro fatal — é o próprio motivo de existir a Etapa 2/3).
        """
        self.log.info(f"📂 Etapa 1 — importando arquivo: {caminho_arquivo}")
        self._trazer_frente(self.main_hwnd)
        time.sleep(0.3)

        # 1. Menu Arquivo → Importar Folha (posição 6, item owner-drawn)
        send_keys("%a")
        time.sleep(0.5)
        for _ in range(5):
            send_keys("{DOWN}")
            time.sleep(0.2)
        send_keys("{ENTER}")
        time.sleep(0.8)

        # 2. Diálogo nativo de seleção de arquivo
        hdialog = 0
        t0 = time.time()
        while time.time() - t0 < 10:
            achados = _enum_toplevel(
                lambda h: _cls(h) == "#32770" and "importação" in _txt(h).lower()
            )
            if achados:
                hdialog = achados[0]
                break
            time.sleep(0.3)
        if not hdialog:
            self.log.error("❌ Diálogo de importação não abriu — verifique o menu Arquivo.")
            return False

        hedit = self._achar_filho_por_ctrl_id_win32(hdialog, 1152, "Edit")
        habrir = None
        for h in _enum_filhos(hdialog):
            if _cls(h) == "Button" and _norm(_txt(h)) == "abrir":
                habrir = h
                break
        if not hedit or not habrir:
            self.log.error("❌ Campo 'Nome' ou botão 'Abrir' não encontrados no diálogo.")
            return False

        win32gui.SendMessage(hedit, win32con.WM_SETTEXT, 0, caminho_arquivo)
        time.sleep(0.3)
        win32gui.SendMessage(habrir, BM_CLICK, 0, 0)
        time.sleep(1)

        # 3. Popup de sobreposição — confirma SIM
        t0 = time.time()
        confirmou_sobreposicao = False
        while time.time() - t0 < timeout:
            hpopup = self._achar_popup(
                titulos=["informação", "informacao"],
                classes=("TMessageForm", "TfrmExibeMsg", "#32770"),
            )
            if hpopup:
                botoes = [_txt(h) for h in _enum_filhos(hpopup)
                         if _cls(h) in ("TButton", "Button")]
                if any(_norm(b) == "sim" for b in botoes):
                    self.log.info("🔔 Popup de sobreposição detectado — confirmando Sim")
                    self._clicar_no_popup(hpopup, "Sim")
                    confirmou_sobreposicao = True
                    time.sleep(1)
                    break
            time.sleep(0.4)

        if not confirmou_sobreposicao:
            self.log.warning("⚠️ Popup de sobreposição não apareceu (pode ser 1ª importação "
                             "dessa empresa/competência, sem dados a sobrescrever).")

        # 4. Fecha popups subsequentes: 'Atenção' (Sim/Não → responde Não) e
        #    'Informação' (OK) — em qualquer ordem, até a tela normal voltar.
        t0 = time.time()
        while time.time() - t0 < timeout:
            if not self._achar_popup(classes=("TMessageForm", "TfrmExibeMsg",
                                              "TfrmTelaAviso", "#32770")):
                self.log.info("✅ Importação concluída — nenhum popup pendente.")
                return True

            hatencao = self._achar_popup(titulos=["atenção", "atencao"],
                                         classes=("TMessageForm",))
            if hatencao:
                botoes = [_txt(h) for h in _enum_filhos(hatencao)
                         if _cls(h) in ("TButton", "Button")]
                if any(_norm(b) == "não" or _norm(b) == "nao" for b in botoes):
                    self.log.info("🔔 Popup 'Atenção' (inconsistência/aviso) — respondendo Não")
                    self._clicar_no_popup(hatencao, "Não")
                    time.sleep(0.6)
                    continue

            hinfo = self._achar_popup(classes=("TMessageForm", "TfrmExibeMsg", "#32770"))
            if hinfo:
                titulo = _txt(hinfo)
                self.log.info(f"🔔 Popup pós-importação (título='{titulo}') — clicando OK")
                self._clicar_no_popup(hinfo, "OK")
                time.sleep(0.6)
                continue

            time.sleep(0.4)

        self.log.warning("⚠️ Ainda há popup(s) pendente(s) após timeout — verifique manualmente.")
        return False

    def _achar_filho_por_ctrl_id_win32(self, parent_hwnd: int, ctrl_id: int, classe: str = None) -> int:
        """Igual ao helper do DomBot_GFIP: acha filho direto por ctrl_id (e classe, se dada)."""
        h = win32gui.GetWindow(parent_hwnd, 5)  # GW_CHILD
        while h:
            try:
                if win32gui.GetDlgCtrlID(h) == ctrl_id and (classe is None or _cls(h) == classe):
                    return h
            except Exception:
                pass
            h = win32gui.GetWindow(h, 2)  # GW_HWNDNEXT
        return 0

    # ══════════════════════════════════════════════════════════════════════
    # ETAPA 2 — Abrir a tela Modalidades (Ctrl+M) na aba Movimento
    # ══════════════════════════════════════════════════════════════════════

    def _achar_arvore_movimento(self) -> int:
        """
        Retorna o hwnd do TTreeView VISÍVEL da árvore de Movimento (há vários
        TTreeView, um por aba; só um está visível por vez).
        """
        candidatos = []
        for h in _enum_filhos(self.main_hwnd):
            if _cls(h) == "TTreeView" and _visivel(h):
                l, _, _, _ = _rect(h)
                candidatos.append((l, h))  # a árvore de Movimento fica à direita
        if not candidatos:
            return 0
        # Normalmente só há 1 visível; se houver mais, pega o de maior x
        candidatos.sort()
        return candidatos[-1][1]

    def selecionar_no_recolhimento(self, indice=3) -> bool:
        """
        Seleciona o nó "Recolhimento ao FGTS e Declaração à Previdê..." na
        árvore de Movimento. O Ctrl+M só abre Modalidades com esse nó
        selecionado.

        Como o TTreeView é VCL (não legível via win32, TVM_GETCOUNT=0) e a
        busca incremental por digitação NÃO funciona neste controle,
        navegamos por ÍNDICE: Home (vai ao 1º nó da árvore) + N vezes {DOWN}.

        A árvore tem nós acima do nível "Trabalhadores sem modalidade" (ex.:
        Cód. Rec., a empresa "N A SCARAMUSSA" etc.) — por isso o índice real
        até "Recolhimento ao FGTS" é maior do que o nível visível sugere.
        Confirmado empiricamente: índice 3 a partir do Home (índice 4 ficava
        um nó abaixo do alvo). Essa contagem só é válida se o nó da empresa
        (índice 1) estiver EXPANDIDO — se estiver colapsado, os nós filhos
        (Trabalhadores, Recolhimento, etc.) não existem na contagem visível e
        o índice 3 cai num nó completamente diferente (testado: abriu
        Modalidades associada ao contexto errado). Por isso, ANTES de contar
        os índices, garantimos a expansão: Home + 1×DOWN (nó da empresa,
        índice 1) + {RIGHT} (expande se colapsado; em TreeViews do Windows,
        RIGHT num nó já expandido só desce para o 1º filho — não fecha nada
        — então é seguro repetir sempre, idempotente).

        `indice`: quantos {DOWN} a partir do Home (padrão 3).
        """
        harvore = self._achar_arvore_movimento()
        if not harvore:
            self.log.warning("⚠️ Árvore de Movimento não localizada — Ctrl+M pode falhar.")
            return False

        self.log.info(f"🌳 Selecionando nó de recolhimento na árvore (índice {indice})")
        self._trazer_frente(self.main_hwnd)
        time.sleep(0.2)
        # Foca a árvore com clique de mouse (VCL responde melhor a clique real)
        self._clicar_mouse(harvore)
        time.sleep(0.3)

        # Garante que o nó da empresa (índice 1) está expandido antes de
        # contar os demais índices — senão a contagem de {DOWN} cai errada.
        send_keys("{HOME}")
        time.sleep(0.2)
        send_keys("{DOWN}")  # nó da empresa (índice 1)
        time.sleep(0.15)
        send_keys("{RIGHT}")  # expande se colapsado; idempotente se já aberto
        time.sleep(0.3)

        # Vai ao topo (1º nó raiz) e desce N vezes até o nó alvo
        send_keys("{HOME}")
        time.sleep(0.2)
        for _ in range(indice):
            send_keys("{DOWN}")
            time.sleep(0.15)
        time.sleep(0.3)
        return True

    def abrir_modalidades(self, timeout=15, selecionar_no=True,
                          no_indice=3) -> int:
        """
        Seleciona o nó de recolhimento (se selecionar_no=True) e envia Ctrl+M
        para abrir a janela Modalidades. Retorna o hwnd dela (ou 0).

        Pré-condição: o SEFIP está na tela de Movimento (TfrmMovEmpresa).
        O Ctrl+M só funciona com o nó 'Recolhimento ao FGTS' selecionado.
        """
        self.log.info("⌨️ Etapa 2 — abrindo Modalidades (Ctrl+M)...")

        if selecionar_no:
            self.selecionar_no_recolhimento(no_indice)

        self._trazer_frente(self.main_hwnd)
        time.sleep(0.3)
        send_keys("^m")

        t0 = time.time()
        while time.time() - t0 < timeout:
            h = self._achar_modalidades()
            if h:
                self.log.info(f"✅ Janela Modalidades aberta (hwnd={h})")
                return h
            time.sleep(0.4)
        self.log.error("❌ Janela Modalidades não abriu após Ctrl+M.")
        self.log.error("   Dica: confirme que o nó 'Recolhimento ao FGTS' foi "
                       "selecionado na árvore. Se o nó tem outro texto, use --no-texto.")
        return 0

    def _achar_modalidades(self) -> int:
        achados = _enum_toplevel(
            lambda h: _cls(h) == "TfrmAdmModalidades" or _txt(h) == "Modalidades"
        )
        return achados[0] if achados else 0

    # ── Coleta de controles da tela Modalidades (por posição) ────────────────

    def _controles_modalidades(self, hmodal):
        """
        Retorna dict com os controles da tela Modalidades, distinguidos por
        posição x (painel esquerdo=Origem, direito=Destino). Recoleta sempre,
        pois os hwnds mudam.
        """
        listviews, combos, edits, botoes = [], [], [], []
        for h in _enum_filhos(hmodal):
            c = _cls(h)
            l, _, _, _ = _rect(h)
            if c in ("TListView", "SysListView32"):
                listviews.append((l, h))
            elif "ComboBox" in c:
                combos.append((l, h))
            elif c == "TEdit":
                edits.append((l, h))
            elif c in ("TBitBtn", "TButton"):
                botoes.append((h, _txt(h)))
        listviews.sort(); combos.sort(); edits.sort()
        return {
            "lv_esq": listviews[0][1] if len(listviews) > 0 else 0,
            "lv_dir": listviews[1][1] if len(listviews) > 1 else 0,
            "combo_esq": combos[0][1] if len(combos) > 0 else 0,
            "combo_dir": combos[1][1] if len(combos) > 1 else 0,
            "busca_esq": edits[0][1] if len(edits) > 0 else 0,
            "busca_dir": edits[1][1] if len(edits) > 1 else 0,
            "botoes": botoes,
        }

    # ══════════════════════════════════════════════════════════════════════
    # ETAPA 3 — Separar o funcionário do ajuste
    # ══════════════════════════════════════════════════════════════════════

    def separar_funcionario(self, hmodal, funcionarios, modalidade_destino="9") -> bool:
        """
        Executa na tela Modalidades:
          1. Seleciona a modalidade de destino (combo direito) — por padrão a 9.
          2. Manda todos p/ direita (>>).
          3. Para CADA funcionário em `funcionarios` (nesta ordem): filtra no
             campo Busca da direita (por nome/CPF) e devolve p/ esquerda (<).
             O próprio filtro VCL já deixa o item correspondente selecionado/
             em foco — NÃO clicamos na lista (um clique em coordenada pegaria
             o item errado, ex.: o 1º visível em vez do filtrado). Repetir
             filtro+< para vários nomes empilha todos na esquerda, um a um —
             a lista da direita nunca é lida, então a ordem não importa e cada
             filtro busca sobre o que sobrou na direita.
        NÃO salva — o Salvar é chamado separadamente (finalizar_modalidades).

        `funcionarios`: trecho do nome/CPF de UM funcionário (string), ou uma
        lista de trechos para separar VÁRIOS funcionários de uma vez.
        """
        if isinstance(funcionarios, str):
            funcionarios = [funcionarios]

        self.log.info(f"🧩 Etapa 3 — separando {len(funcionarios)} funcionário(s): "
                     f"{funcionarios} (modalidade destino {modalidade_destino})")

        ctrls = self._controles_modalidades(hmodal)
        if not ctrls["lv_dir"] or not ctrls["busca_dir"]:
            self.log.error("❌ Não consegui localizar os painéis da tela Modalidades.")
            return False

        # 1. Selecionar modalidade de destino (combo direito)
        if not self._selecionar_modalidade_destino(ctrls["combo_dir"], modalidade_destino):
            self.log.warning("⚠️ Não confirmei a seleção da modalidade destino — "
                             "verifique se a 9 já estava selecionada.")

        # 2. Mandar todos para a direita (>>)
        self._trazer_frente(hmodal)
        if not self._clicar_botao_por_texto(hmodal, ">>"):
            return False
        time.sleep(0.5)

        # 3. Para cada funcionário: filtrar na Busca da direita + devolver (<)
        for i, funcionario in enumerate(funcionarios, start=1):
            self.log.info(f"🔎 [{i}/{len(funcionarios)}] Filtrando '{funcionario}' "
                         f"no campo Busca (direita)")
            ctrls = self._controles_modalidades(hmodal)  # recoleta (hwnds mudam)
            hbusca = ctrls["busca_dir"]
            if not hbusca:
                self.log.error(f"❌ Campo Busca (direita) não encontrado para '{funcionario}'.")
                return False
            # Reforça o foreground a CADA iteração — o clique em '<' da
            # rodada anterior não garante que a Modalidades continua em
            # foco. Além disso, Ctrl+A/Backspace via send_keys NÃO limpou o
            # campo de forma confiável ao testar com múltiplos funcionários
            # (o 2º nome ficou concatenado ao 1º: "DAMIAO JOSANDRIELE DA
            # SILVA" — o foco de teclado real não estava no campo apesar do
            # clique). Corrigido: limpar via WM_SETTEXT (mensagem direta ao
            # controle, não depende de foco do sistema) ANTES de clicar e
            # digitar o novo texto.
            self._trazer_frente(hmodal)
            time.sleep(0.15)
            win32gui.SendMessage(hbusca, win32con.WM_SETTEXT, 0, "")
            time.sleep(0.15)
            self._clicar_mouse(hbusca)
            time.sleep(0.15)
            send_keys(funcionario, with_spaces=True)
            time.sleep(0.8)  # deixa o filtro assentar e selecionar o item

            # Devolve o funcionário filtrado para a esquerda (<), direto após
            # o filtro, com o item já selecionado pela Busca.
            if not self._clicar_botao_por_texto(hmodal, "<"):
                self.log.error(f"❌ Falha ao devolver '{funcionario}' para a esquerda.")
                return False
            time.sleep(0.4)

        self.log.info(f"✅ {len(funcionarios)} funcionário(s) separado(s) "
                     "(movidos para a modalidade de origem). Revise antes de Salvar.")
        return True

    def _selecionar_modalidade_destino(self, hcombo, valor) -> bool:
        """
        Seleciona a modalidade no combo destino. Como o combo é VCL custom
        (CB_* não confiável), abrimos com clique + digitamos o número/texto.
        """
        if not hcombo:
            return False
        self.log.info(f"📋 Selecionando modalidade destino '{valor}'")
        self._clicar_mouse(hcombo)
        time.sleep(0.2)
        # Abre a lista suspensa e digita o começo do texto (ex.: '9')
        send_keys("%{DOWN}")   # Alt+Down abre o dropdown
        time.sleep(0.3)
        send_keys(str(valor))
        time.sleep(0.3)
        send_keys("{ENTER}")
        time.sleep(0.3)
        return True

    def finalizar_modalidades(self, hmodal, salvar=True, timeout=25) -> bool:
        """
        Clica Salvar (ou Cancelar) na tela Modalidades. Ao salvar, o SEFIP
        mostra um popup 'Informação' — 'Alterações efetuadas com sucesso!'
        (TMessageForm, botão OK) que o bot aguarda e fecha automaticamente.

        O texto do corpo desse popup NÃO é lido de forma confiável via
        EnumChildWindows (é owner-draw em alguns casos, como já visto no
        popup 'Atenção' de inconsistências da importação) — por isso o
        match usa APENAS o título 'Informação' (curto e estável), sem
        depender de corpo_contem. O timeout foi ampliado (25s) e o polling
        agora loga um heartbeat a cada ~5s para diferenciar "ainda esperando"
        de "travou".
        """
        botao = "Salvar" if salvar else "Cancelar"
        self.log.info(f"💾 Finalizando Modalidades — {botao}")
        self._trazer_frente(hmodal)
        # assincrono=True no Salvar: SendMessage travaria o processo Python
        # até o popup de sucesso ser fechado (o SEFIP processa de forma
        # bloqueante e prende a thread de mensagens dentro do próprio clique).
        ok = self._clicar_botao_por_texto(hmodal, botao, assincrono=salvar)
        if not ok:
            return False

        if not salvar:
            return True

        # Aguarda QUALQUER popup #32770/TMessageForm surgir e fecha via OK.
        # Não filtramos por corpo (não confiável); título "Informação" já
        # é suficiente para identificar o popup de sucesso nesse fluxo.
        t0 = time.time()
        confirmado = False
        ultimo_heartbeat = t0
        while time.time() - t0 < timeout:
            hpopup = self._achar_popup(classes=("#32770", "TMessageForm"))
            if hpopup:
                titulo = _txt(hpopup)
                self.log.info(f"🔔 Popup pós-Salvar detectado (título='{titulo}') — clicando OK")
                self._clicar_no_popup(hpopup, "OK")
                confirmado = True
                break
            if time.time() - ultimo_heartbeat > 5:
                self.log.info(f"⏳ Aguardando popup pós-Salvar... ({int(time.time()-t0)}s)")
                ultimo_heartbeat = time.time()
            time.sleep(0.4)

        if not confirmado:
            self.log.warning("⚠️ Nenhum popup de confirmação apareceu após Salvar "
                             "(timeout). Verifique a tela manualmente.")

        # Verifica se a janela Modalidades fechou (indício final de sucesso)
        t0 = time.time()
        while time.time() - t0 < 8:
            if not self._achar_modalidades():
                self.log.info("✅ Janela Modalidades fechou. Gravação concluída.")
                return True
            time.sleep(0.4)
        self.log.warning("⚠️ Janela Modalidades ainda aberta — confira manualmente.")
        return confirmado

    # ══════════════════════════════════════════════════════════════════════
    # ETAPA 4 — Executar (gerar/fechar movimento) e Geração de Arquivo
    # ══════════════════════════════════════════════════════════════════════

    def selecionar_no_inicio(self) -> bool:
        """
        Seleciona o primeiro nó (topo/raiz) da árvore de Movimento — pré-
        requisito para a tela 'Abertura de Movimento' com o botão Executar.
        Diferente do índice 3 usado para Ctrl+M/Modalidades.
        """
        harvore = self._achar_arvore_movimento()
        if not harvore:
            self.log.warning("⚠️ Árvore de Movimento não localizada.")
            return False
        self.log.info("🌳 Selecionando nó inicial (índice 0) da árvore")
        self._trazer_frente(self.main_hwnd)
        time.sleep(0.2)
        self._clicar_mouse(harvore)
        time.sleep(0.3)
        send_keys("{HOME}")
        time.sleep(0.3)
        return True

    def _achar_tela_abertura(self) -> int:
        """Retorna hwnd do TfrmManterAbertura (tela com o botão Executar), ou 0."""
        for h in _enum_filhos(self.main_hwnd):
            if _cls(h) == "TfrmManterAbertura":
                return h
        return 0

    def executar_movimento(self, timeout_processamento=180, ler_inconsistencia=True) -> str:
        """
        Clica no botão 'Executar' da tela Abertura de Movimento e aguarda
        o resultado. O clique é SÍNCRONO/BLOQUEANTE no SEFIP (o processamento
        pode levar dezenas de segundos a alguns minutos) — por isso usamos
        PostMessage (assíncrono) em vez de SendMessage, e faz polling do
        resultado em vez de esperar o retorno da mensagem.

        `ler_inconsistencia`: se True (padrão) e o Executar detectar
        inconsistência, abre o relatório (Sim), gera o PDF e tenta extrair a
        mensagem de erro (ver _ler_relatorio_inconsistencia), deixando o
        resultado em self.mensagem_inconsistencia. Se False, mantém o
        comportamento antigo (responde Não, não vê o relatório).

        Retorna:
          'sucesso'      — popup 'Informação'+OK único; abriu tela de Geração
                           de Arquivo (TfrmGravarDisco).
          'inconsistencia' — popup 'Atenção'+Sim/Não (ex.: FAP inválido);
                           usuário precisa corrigir o cadastro manualmente.
          'data_invalida' — popup 'Confirmação' avisando que a data de
                           atraso está fora da vigência do edital de
                           índices; usuário precisa regerar o .re com uma
                           data dentro da janela informada pelo SEFIP.
          'timeout'      — não detectou nenhum popup no tempo esperado.
        """
        htela = self._achar_tela_abertura()
        if not htela:
            self.log.error("❌ Tela 'Abertura de Movimento' não encontrada.")
            return "erro"

        hexecutar = None
        for h in _enum_filhos(htela):
            if _cls(h) == "TBitBtn" and _norm(_txt(h)) == "executar":
                hexecutar = h
                break
        if not hexecutar:
            self.log.error("❌ Botão 'Executar' não encontrado.")
            return "erro"

        self.log.info("▶ Clicando 'Executar' (processamento pode demorar)...")
        self._trazer_frente(self.main_hwnd)
        time.sleep(0.3)
        # PostMessage (assíncrono) — NÃO usar SendMessage aqui: o SEFIP
        # processa o fechamento do movimento de forma bloqueante e
        # SendMessage ficaria esperando minutos pelo retorno.
        win32gui.PostMessage(hexecutar, BM_CLICK, 0, 0)

        t0 = time.time()
        ultimo_heartbeat = t0
        while time.time() - t0 < timeout_processamento:
            # Sucesso: popup 'Informação' com botão OK único, seguido da
            # tela de Geração de Arquivo (TfrmGravarDisco). A classe varia
            # entre TfrmExibeMsg e TMessageForm (confirmado na prática).
            hinfo = self._achar_popup(titulos=["informação", "informacao"],
                                      classes=("TfrmExibeMsg", "TMessageForm", "#32770"))
            if hinfo:
                self.log.info("🔔 Popup 'Informação' detectado — Executar concluído com sucesso")
                self._clicar_no_popup(hinfo, "OK")
                time.sleep(1)
                return "sucesso"

            # Data de atraso fora da vigência do edital: popup 'Confirmação'
            # (não 'Atenção'!) perguntando se quer carregar a tabela de
            # índices mesmo assim. Testado na prática: mesmo respondendo Sim,
            # o processamento segue e AINDA ASSIM termina em inconsistência
            # (o SEFIP não convalida a data fora da janela) — por isso aqui
            # respondemos Não e abortamos direto, sem tentar prosseguir.
            # O corpo do popup não é lido de forma confiável (mesma limitação
            # de sempre); o título "Confirmação" já é suficiente para
            # distinguir esse caso de "Informação"/"Atenção".
            hconfirmacao = self._achar_popup(titulos=["confirmação", "confirmacao"],
                                             classes=("TfrmExibeMsg", "TMessageForm", "#32770"))
            if hconfirmacao:
                self.log.error("🔔 Popup 'Confirmação' detectado — provável data de atraso fora "
                              "da vigência do edital de índices. Respondendo Não e abortando: "
                              "regere o .re com uma data válida (o SEFIP informa a janela exata "
                              "no texto do popup — confira manualmente).")
                self._clicar_no_popup(hconfirmacao, "Não")
                time.sleep(1)
                return "data_invalida"

            # Inconsistência: popup 'Atenção' com Sim/Não (pergunta se quer
            # ver o relatório de inconsistências).
            hatencao = self._achar_popup(titulos=["atenção", "atencao"],
                                         classes=("TfrmExibeMsg", "TMessageForm", "#32770"))
            if hatencao:
                if ler_inconsistencia:
                    self.log.warning("🔔 Popup 'Atenção' detectado — inconsistência no "
                                     "cadastro. Abrindo relatório (Sim) para identificar a causa.")
                    self._clicar_no_popup(hatencao, "Sim")
                    time.sleep(1.5)
                    self.mensagem_inconsistencia = self._ler_relatorio_inconsistencia()
                    if self.mensagem_inconsistencia:
                        self.log.error(f"📄 Inconsistência (extraída do PDF): "
                                      f"{self.mensagem_inconsistencia}")
                    else:
                        self.log.warning("⚠️ Não consegui extrair o texto do relatório de "
                                         "inconsistência — confira manualmente o PDF/tela.")
                else:
                    self.log.warning("🔔 Popup 'Atenção' detectado — inconsistência no cadastro "
                                     "(ex.: FAP inválido). Fechando sem ver relatório (Não).")
                    self._clicar_no_popup(hatencao, "Não")
                time.sleep(1)
                return "inconsistencia"

            if time.time() - ultimo_heartbeat > 10:
                self.log.info(f"⏳ Aguardando processamento do Executar... ({int(time.time()-t0)}s)")
                ultimo_heartbeat = time.time()
            time.sleep(0.5)

        self.log.warning("⚠️ Timeout aguardando resultado do Executar.")
        return "timeout"

    def _achar_tela_inconsistencia(self) -> int:
        achados = _enum_toplevel(
            lambda h: _cls(h) == "TfrmImprimeRelInconsistencia" or
                     "inconsist" in _txt(h).lower())
        return achados[0] if achados else 0

    def _ler_relatorio_inconsistencia(self, timeout=20) -> str:
        """
        Na tela 'Impressão de Relatório' (TfrmImprimeRelInconsistencia,
        botões Visualizar/Imprimir/Gerar PDF/Fechar): clica 'Gerar PDF',
        salva num caminho temporário previsível, lê o texto do PDF (via
        pypdf) e fecha a tela. Contorna a limitação antiga de que o
        TQRPreview (relatório gráfico) não é legível via win32 — o PDF tem
        o mesmo conteúdo, só que como texto extraível.

        Retorna a mensagem de inconsistência extraída (string), ou "" se
        não foi possível gerar/ler o PDF.
        """
        htela = self._achar_tela_inconsistencia()
        if not htela:
            self.log.warning("⚠️ Tela de relatório de inconsistência não encontrada.")
            return ""

        hgerar = None
        for h in _enum_filhos(htela):
            if _cls(h) in ("TBitBtn", "TButton") and "pdf" in _norm(_txt(h)):
                hgerar = h
                break
        if not hgerar:
            self.log.warning("⚠️ Botão 'Gerar PDF' não encontrado na tela de inconsistência.")
            return ""

        self.log.info("📄 Clicando 'Gerar PDF' para capturar a inconsistência...")
        self._trazer_frente(htela)
        time.sleep(0.3)
        win32gui.PostMessage(hgerar, BM_CLICK, 0, 0)
        time.sleep(1.5)

        # Diálogo clássico "Salvar como" (Nome=ctrl_id 1152, botão=ctrl_id 1),
        # mesmo padrão já usado em outras telas do SEFIP.
        caminho_pdf = os.path.join(
            os.environ.get("TEMP", "."), f"sefip_inconsistencia_{int(time.time())}.pdf")
        t0 = time.time()
        hdialog = 0
        while time.time() - t0 < timeout:
            hdialog = None
            for h in _enum_toplevel(lambda h: _cls(h) == "#32770"):
                hdialog = h
                break
            if hdialog:
                break
            time.sleep(0.3)
        if not hdialog:
            self.log.warning("⚠️ Diálogo 'Salvar PDF como' não apareceu.")
            return ""

        hedit = self._achar_filho_por_ctrl_id_win32(hdialog, 1152, "Edit")
        hsalvar = None
        for h in _enum_filhos(hdialog):
            if _cls(h) == "Button" and win32gui.GetDlgCtrlID(h) == 1:
                hsalvar = h
                break
        if not hedit or not hsalvar:
            self.log.warning("⚠️ Campo Nome ou botão Salvar não encontrados no diálogo do PDF.")
            return ""

        self._trazer_frente(hdialog)
        time.sleep(0.3)
        win32gui.SendMessage(hedit, win32con.WM_SETTEXT, 0, caminho_pdf)
        time.sleep(0.3)
        win32gui.PostMessage(hsalvar, BM_CLICK, 0, 0)
        time.sleep(1.5)

        # Aguarda o arquivo aparecer em disco (o SEFIP grava de forma
        # assíncrona em relação ao clique) e fecha popups residuais (OK).
        t0 = time.time()
        while time.time() - t0 < timeout:
            if os.path.exists(caminho_pdf) and os.path.getsize(caminho_pdf) > 0:
                break
            hpopup = self._achar_popup(classes=("TfrmExibeMsg", "TMessageForm", "#32770"))
            if hpopup:
                self._clicar_no_popup(hpopup, "OK")
                time.sleep(0.5)
            time.sleep(0.4)

        texto = ""
        if os.path.exists(caminho_pdf):
            texto = self._extrair_erro_pdf_inconsistencia(caminho_pdf)
            try:
                os.remove(caminho_pdf)
            except Exception:
                pass
        else:
            self.log.warning(f"⚠️ PDF não encontrado em '{caminho_pdf}' após Gerar PDF.")

        # Fecha a tela de relatório (botão 'Fechar'), liberando o SEFIP.
        htela = self._achar_tela_inconsistencia()
        if htela:
            self._clicar_botao_por_texto(htela, "Fechar", classes=("TBitBtn", "TButton"),
                                         assincrono=True)
            time.sleep(1)

        return texto

    def _extrair_erro_pdf_inconsistencia(self, caminho: str) -> str:
        """
        Extrai a mensagem de erro do "Relatório de Inconsistências do
        Fechamento" gerado pelo SEFIP. Formato real confirmado (2026-07-31,
        empresa GELOARTE, cód. 101177 — FAP inválido):

            CONTEÚDO DO CAMPO
            CÓDIGO - DESCRIÇÃO DO ERRO
            TRABALHADOR
            03.081.396/0001-37
            0,00
            101177-ALÍQUOTA FAP INVÁLIDA. O VALOR DEVE ESTAR ENTRE 0,5 E 2,00.

        A linha do erro em si é sempre "<código numérico>-<descrição em
        maiúsculas>", isolada numa linha própria — é isso que a regex busca,
        não a posição fixa (o relatório pode ter mais ou menos linhas de
        cabeçalho dependendo da versão/idioma do SEFIP).
        """
        if PdfReader is None:
            self.log.warning("⚠️ Biblioteca pypdf não disponível — não consigo ler o PDF.")
            return ""
        try:
            leitor = PdfReader(caminho)
            texto_bruto = "\n".join(p.extract_text() or "" for p in leitor.pages)
        except Exception as e:
            self.log.warning(f"⚠️ Falha ao ler o PDF de inconsistência: {type(e).__name__}: {e}")
            return ""

        import re
        # Código: 4-6 dígitos. Descrição: resto da linha (maiúsculas/acentos/
        # pontuação), até a quebra de linha do PDF.
        erros = re.findall(r"^(\d{4,6})-(.+)$", texto_bruto, flags=re.MULTILINE)
        if erros:
            linhas = [f"{codigo} - {descricao.strip()}" for codigo, descricao in erros]
            return " | ".join(linhas)

        # Fallback: não achou o padrão esperado — devolve o texto inteiro
        # normalizado, para não perder a informação (mesmo sem conseguir
        # isolar só a linha do erro).
        self.log.warning("⚠️ Não encontrei o padrão 'código-descrição' esperado no PDF — "
                         "retornando o texto completo extraído.")
        return " ".join(texto_bruto.split())

    def _achar_tela_geracao_arquivo(self) -> int:
        achados = _enum_toplevel(lambda h: _cls(h) == "TfrmGravarDisco")
        return achados[0] if achados else 0

    def salvar_arquivo_saida(self, timeout=30) -> bool:
        """
        Na tela 'SEFIP - Geração de Arquivo de Saída' (TfrmGravarDisco),
        aceita o caminho PADRÃO já preenchido (não mexe no diálogo nativo
        'Procurar Pasta' — esse diálogo não aceita digitação/paste, só
        navegação visual, e uma tentativa de ler sua árvore via
        VirtualAllocEx cross-process derrubou o SEFIP) e clica Salvar.

        O caminho padrão observado é 'C:\\Program Files (x86)\\CAIXA\\SEFIP'.
        Após salvar, o arquivo final (nome fixo gerado pelo SEFIP, ex.:
        'K6C8TpMY5cI00001.SFP') deve ser localizado nessa pasta e movido
        para o destino desejado por quem chama este método (ver
        mover_arquivo_gerado).
        """
        htela = self._achar_tela_geracao_arquivo()
        if not htela:
            self.log.error("❌ Tela 'Geração de Arquivo de Saída' não encontrada.")
            return False

        hsalvar = None
        for h in _enum_filhos(htela):
            if _cls(h) == "TBitBtn" and _norm(_txt(h)) == "salvar":
                hsalvar = h
                break
        if not hsalvar:
            self.log.error("❌ Botão 'Salvar' não encontrado na tela de Geração de Arquivo.")
            return False

        self.log.info("💾 Aceitando caminho padrão e clicando Salvar...")
        self._trazer_frente(htela)
        time.sleep(0.3)
        # PostMessage (assíncrono): o Salvar dispara uma sequência de até 3
        # popups "Informação" (confirmado na prática — classes variam entre
        # TfrmExibeMsg e TMessageForm, todos com botão OK único) ANTES da
        # janela fechar. SendMessage travaria esperando o 1º popup ser
        # fechado, então não dá pra usar aqui.
        win32gui.PostMessage(hsalvar, BM_CLICK, 0, 0)
        time.sleep(0.5)

        # Fecha (via OK) todos os popups "Informação" que aparecerem, até a
        # janela de Geração de Arquivo fechar ou o timeout estourar.
        t0 = time.time()
        while time.time() - t0 < timeout:
            if not self._achar_tela_geracao_arquivo():
                self.log.info("✅ Arquivo salvo — janela de Geração de Arquivo fechou.")
                return True
            hpopup = self._achar_popup(classes=("TfrmExibeMsg", "TMessageForm", "#32770"))
            if hpopup:
                self.log.info(f"🔔 Popup pós-Salvar detectado (título='{_txt(hpopup)}') — clicando OK")
                self._clicar_no_popup(hpopup, "OK")
                time.sleep(0.6)
                continue
            time.sleep(0.4)
        self.log.warning("⚠️ Janela de Geração de Arquivo ainda aberta após timeout.")
        return False


CAMINHO_PADRAO_SEFIP = r"C:\Program Files (x86)\CAIXA\SEFIP"


def mover_arquivo_gerado(pasta_destino: str, logger=None, pasta_origem: str = CAMINHO_PADRAO_SEFIP) -> str:
    """
    Localiza o(s) arquivo(s) .SFP mais recente(s) na pasta padrão do SEFIP
    (CAMINHO_PADRAO_SEFIP) e move para `pasta_destino`. O nome do arquivo é
    gerado automaticamente pelo SEFIP (código alfanumérico + extensão .SFP)
    e não é previsível — por isso localizamos pelo mais recentemente
    modificado, não por nome.

    Retorna o caminho final do arquivo movido, ou None se nada foi encontrado.
    """
    import glob
    import os
    import shutil

    log = logger or ConsoleLogger()

    candidatos = glob.glob(os.path.join(pasta_origem, "*.SFP")) + \
                 glob.glob(os.path.join(pasta_origem, "*.sfp"))
    if not candidatos:
        log.error(f"❌ Nenhum arquivo .SFP encontrado em '{pasta_origem}'.")
        return None

    mais_recente = max(candidatos, key=os.path.getmtime)
    os.makedirs(pasta_destino, exist_ok=True)
    destino_final = os.path.join(pasta_destino, os.path.basename(mais_recente))
    shutil.move(mais_recente, destino_final)
    log.info(f"📁 Arquivo movido: '{mais_recente}' → '{destino_final}'")
    return destino_final


# ── Fluxo de teste (CLI) ─────────────────────────────────────────────────────

def run_teste(args):
    """
    Executa as Etapas 1-3 de forma controlada, para teste manual.
    Cada etapa pode ser ligada/desligada por flag para testar isolado.

    IMPORTANTE: rode em PowerShell como Administrador (SEFIP é elevado).
    """
    bot = SefipAutomation()

    if not bot.conectar():
        sys.exit(1)

    # Etapa 1 — importação completa (menu → diálogo → sobreposição → avisos)
    if args.etapa1:
        if not args.arquivo:
            bot.log.error("Etapa 1 requer --arquivo 'CAMINHO\\Sefip.re'.")
            sys.exit(1)
        if not bot.importar_arquivo(args.arquivo):
            bot.log.error("❌ Falha na importação — verifique a tela manualmente.")
            sys.exit(1)

    # Etapa 2 — abrir Modalidades
    hmodal = 0
    if args.etapa2:
        hmodal = bot.abrir_modalidades(selecionar_no=not args.sem_selecionar_no,
                                       no_indice=args.no_indice)
        if not hmodal:
            bot.log.error("Abortando: Modalidades não abriu.")
            sys.exit(1)
    else:
        # Se pulou etapa 2, tenta achar Modalidades já aberta
        hmodal = bot._achar_modalidades()

    # Etapa 3 — separar funcionário
    if args.etapa3:
        if not hmodal:
            bot.log.error("Etapa 3 requer a tela Modalidades aberta.")
            sys.exit(1)
        if not args.funcionario:
            bot.log.error("Etapa 3 requer --funcionario 'NOME OU CPF' [...].")
            sys.exit(1)
        ok = bot.separar_funcionario(hmodal, args.funcionario, args.modalidade)
        if ok and args.salvar:
            bot.finalizar_modalidades(hmodal, salvar=True)
        elif ok:
            bot.log.info("ℹ️ --salvar NÃO informado: parei antes de gravar. "
                         "Revise a tela e salve manualmente, ou rode com --salvar.")

    # Etapa 4 — Executar movimento e gerar arquivo de saída
    if args.etapa4:
        if not args.dir_saida:
            bot.log.error("Etapa 4 requer --dir-saida 'PASTA DESTINO FINAL'.")
            sys.exit(1)

        bot.selecionar_no_inicio()
        resultado = bot.executar_movimento()

        if resultado == "inconsistencia":
            if bot.mensagem_inconsistencia:
                bot.log.error(f"❌ Inconsistência no cadastro: {bot.mensagem_inconsistencia}")
            else:
                bot.log.error("❌ Inconsistência no cadastro (ex.: FAP inválido) — não foi "
                              "possível extrair a mensagem exata do PDF, confira manualmente.")
            bot.log.error("   Corrija o cadastro e rode a Etapa 4 novamente.")
            sys.exit(1)
        elif resultado == "data_invalida":
            bot.log.error("❌ Data de atraso fora da vigência do edital de índices. "
                          "Confira a janela de datas exibida no popup (na tela do SEFIP) "
                          "e regere o .re com --data-atraso dentro dela.")
            sys.exit(1)
        elif resultado != "sucesso":
            bot.log.error(f"❌ Executar não teve sucesso (resultado='{resultado}').")
            sys.exit(1)

        if not bot.salvar_arquivo_saida():
            bot.log.error("❌ Falha ao salvar o arquivo de saída.")
            sys.exit(1)

        destino = mover_arquivo_gerado(args.dir_saida, logger=bot.log)
        if not destino:
            sys.exit(1)
        bot.log.info(f"✅ Etapa 4 concluída — arquivo final: {destino}")

    bot.log.info("🏁 Fluxo de teste concluído.")


def main():
    p = argparse.ArgumentParser(description="Bot SEFIP — recálculo de FGTS (Etapas 1-4)")
    p.add_argument("--etapa1", action="store_true",
                   help="Importar arquivo completo (menu Arquivo → diálogo → sobreposição → avisos)")
    p.add_argument("--arquivo", help="Caminho completo do .re/.txt a importar (Etapa 1)")
    p.add_argument("--etapa2", action="store_true", help="Abrir Modalidades (Ctrl+M)")
    p.add_argument("--etapa3", action="store_true", help="Separar funcionário na tela Modalidades")
    p.add_argument("--etapa4", action="store_true",
                   help="Executar movimento + salvar arquivo (caminho padrão) + mover para --dir-saida")
    p.add_argument("--funcionario", nargs="+",
                   help="Nome(s)/CPF(s) do(s) funcionário(s) a separar (Etapa 3). "
                        "Aceita múltiplos valores: --funcionario \"FULANO\" \"CICLANO\"")
    p.add_argument("--modalidade", default="9", help="Modalidade de destino (padrão: 9)")
    p.add_argument("--no-indice", dest="no_indice", type=int, default=3,
                   help="Quantos {DOWN} a partir do topo da árvore até o nó 'Recolhimento ao FGTS' (padrão: 3)")
    p.add_argument("--sem-selecionar-no", dest="sem_selecionar_no", action="store_true",
                   help="Não tenta selecionar o nó na árvore antes do Ctrl+M (se você já selecionou manual)")
    p.add_argument("--salvar", action="store_true", help="Clicar Salvar ao fim da Etapa 3 (GRAVA de verdade)")
    p.add_argument("--dir-saida", dest="dir_saida",
                   help="Pasta final para onde mover o arquivo .SFP gerado (Etapa 4)")
    p.add_argument("--tudo", action="store_true", help="Atalho: liga etapa1+etapa2+etapa3+etapa4")
    args = p.parse_args()

    if args.tudo:
        args.etapa1 = args.etapa2 = args.etapa3 = args.etapa4 = True

    if not (args.etapa1 or args.etapa2 or args.etapa3 or args.etapa4):
        p.print_help()
        print("\nExemplos:")
        print('  # só testar abrir Modalidades e separar (sem salvar):')
        print('  python Bot_Sefip.py --etapa2 --etapa3 --funcionario "DAMIAO JOSE"')
        print('  # separar VÁRIOS funcionários de uma vez:')
        print('  python Bot_Sefip.py --etapa2 --etapa3 --funcionario "DAMIAO JOSE" "CARLOS AUGUSTO" --salvar')
        print('  # etapa 4: executar + salvar + mover arquivo para a pasta da empresa:')
        print('  python Bot_Sefip.py --etapa4 --dir-saida "C:\\...\\results\\404"')
        print('  # fluxo COMPLETO 1-4, do zero (SEFIP já aberto na empresa/competência certa):')
        print('  python Bot_Sefip.py --tudo --arquivo "C:\\...\\results\\404\\Sefip.re" '
              '--funcionario "DAMIAO JOSE" --salvar --dir-saida "C:\\...\\results\\404"')
        sys.exit(0)

    if not _esta_elevado():
        print("❌ Este processo NÃO está rodando como Administrador.")
        print("   O SEFIP roda elevado — sem admin, a UIPI bloqueia o envio de")
        print("   teclas/cliques de forma SILENCIOSA (o bot parece 'clicar no")
        print("   controle errado' ou travar sem erro claro).")
        print("   Abra um PowerShell 'Executar como Administrador' e rode de novo.")
        sys.exit(1)

    try:
        run_teste(args)
    except Exception as e:
        print(f"Erro crítico: {e}")
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
