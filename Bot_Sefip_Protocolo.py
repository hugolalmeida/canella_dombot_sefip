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
    WIN32 = True
except ImportError:
    WIN32 = False

SUBPASTA_PROTOCOLO = "protocolo"


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
    """Mesma heurística do Bot_Sefip.py: classe/título TfrmPrincipalSEFIP."""
    achados = []

    def _cb(h, _):
        try:
            if win32gui.IsWindowVisible(h):
                c, t = _cls(h), _txt(h)
                if (c.startswith("Tfrm") and "SEFIP" in c.upper()) or \
                   (t.upper().startswith("SEFIP") and c.startswith("T")):
                    achados.append(h)
        except Exception:
            pass
        return True

    win32gui.EnumWindows(_cb, None)
    return achados


def _dump_menu(hmenu, profundidade=0, caminho="", saida=None):
    """
    Enumera recursivamente a barra de menus do SEFIP.

    É a forma mais confiável de descobrir o caminho da validação de protocolo:
    em vez de tentar Alt+letra às cegas, lemos os rótulos reais e os IDs de
    comando (que podem ser acionados por WM_COMMAND).
    """
    saida = saida if saida is not None else []
    try:
        total = win32gui.GetMenuItemCount(hmenu)
    except Exception:
        return saida

    for i in range(total):
        try:
            rotulo = win32gui.GetMenuString(hmenu, i, win32con.MF_BYPOSITION)
        except Exception:
            rotulo = ""
        rotulo = (rotulo or "").replace("&", "").strip()
        try:
            sub = win32gui.GetSubMenu(hmenu, i)
        except Exception:
            sub = 0
        try:
            id_cmd = win32gui.GetMenuItemID(hmenu, i)
        except Exception:
            id_cmd = -1

        completo = f"{caminho} > {rotulo}" if caminho else rotulo
        if rotulo:
            saida.append({"caminho": completo, "nivel": profundidade,
                          "id": None if sub else id_cmd, "tem_submenu": bool(sub)})
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
        for item in _dump_menu(hmenu):
            prefixo = "  " * item["nivel"]
            marca = "▸" if item["tem_submenu"] else " "
            id_txt = "" if item["id"] in (None, -1, 0) else f"   [id={item['id']}]"
            print(f"  {prefixo}{marca} {item['caminho'].split(' > ')[-1]}{id_txt}")

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


def main():
    p = argparse.ArgumentParser(
        description="Bot SEFIP — Etapa 6 (validação do protocolo)")
    p.add_argument("--spy", action="store_true",
                   help="Mapeia menus/janelas do SEFIP (rodar com o SEFIP aberto, como Admin)")
    p.add_argument("--conferir", metavar="PASTA_EMPRESA",
                   help="Confere os artefatos do envio sem tocar no SEFIP")
    args = p.parse_args()

    if args.spy:
        sys.exit(spy())
    if args.conferir:
        sys.exit(_imprimir_conferencia(args.conferir))

    p.print_help()
    print("\nExemplos:")
    print('  python Bot_Sefip_Protocolo.py --conferir ".\\results\\404"')
    print("  python Bot_Sefip_Protocolo.py --spy")


if __name__ == "__main__":
    main()
