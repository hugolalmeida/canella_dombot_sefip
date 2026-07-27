"""
Conectividade Social — Spy de exploração (Playwright)
======================================================
Ferramenta de mapeamento do site da Conectividade Social, análoga ao
sefip_spy.py mas para navegador em vez de win32.

Abre o Google Chrome REAL (não o Chromium isolado do Playwright) usando
--channel=chrome, o que permite acesso aos certificados digitais A1
instalados no repositório do Windows/Chrome (necessário para autenticação).

Uso:
    python conectividade_spy.py

Abre o navegador e mantém a sessão ativa para você navegar manualmente
(escolher o certificado, fazer login). Depois de você navegar até a tela
que queremos mapear, rode em outro terminal (ou pressione Enter aqui) para
capturar o HTML/estrutura da página atual.

Autor: Hugo L. Almeida
"""

import sys
import time

from playwright.sync_api import sync_playwright

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

URL_INICIAL = "https://conectividadesocialv2.caixa.gov.br/sicns/"


def dump_estrutura(page, max_chars=8000):
    """Imprime um resumo da estrutura da página atual: título, URL, e os
    elementos interativos visíveis (links, botões, inputs) com seus
    seletores mais estáveis (id, texto, aria-label)."""
    print("=" * 78)
    print(f"URL atual: {page.url}")
    print(f"Título: {page.title()}")
    print("=" * 78)

    # Inclui componentes Angular Material (mat-select, mat-form-field etc.)
    # além dos elementos HTML nativos — o Material não usa <select> real.
    elementos = page.eval_on_selector_all(
        "a, button, input, select, textarea, "
        "mat-select, mat-form-field, [role='button'], [role='link'], "
        "[role='combobox'], [role='listbox'], [formcontrolname]",
        """els => els.map(el => ({
            tag: el.tagName.toLowerCase(),
            id: el.id || null,
            name: el.getAttribute('name') || null,
            type: el.getAttribute('type') || null,
            text: (el.innerText || el.value || '').trim().slice(0, 60),
            ariaLabel: el.getAttribute('aria-label') || null,
            placeholder: el.getAttribute('placeholder') || null,
            formControlName: el.getAttribute('formcontrolname') || null,
            role: el.getAttribute('role') || null,
            visible: !!(el.offsetWidth || el.offsetHeight || el.getClientRects().length),
            classes: el.className && typeof el.className === 'string' ? el.className.slice(0, 80) : null
        })).filter(e => e.visible)"""
    )

    print(f"\n{len(elementos)} elemento(s) interativo(s) visível(is):\n")
    for el in elementos:
        partes = [f"<{el['tag']}>"]
        if el["id"]:
            partes.append(f"id={el['id']!r}")
        if el["name"]:
            partes.append(f"name={el['name']!r}")
        if el["type"]:
            partes.append(f"type={el['type']!r}")
        if el["text"]:
            partes.append(f"text={el['text']!r}")
        if el["ariaLabel"]:
            partes.append(f"aria-label={el['ariaLabel']!r}")
        if el["placeholder"]:
            partes.append(f"placeholder={el['placeholder']!r}")
        if el.get("formControlName"):
            partes.append(f"formControlName={el['formControlName']!r}")
        if el.get("role"):
            partes.append(f"role={el['role']!r}")
        print("  " + "  ".join(partes))


def main():
    print("Abrindo Google Chrome (canal 'chrome' — acessa certificados A1 instalados)...")
    with sync_playwright() as p:
        # persistent_context com user_data_dir separado: usa o motor do
        # Chrome instalado, mas com um perfil PRÓPRIO (não mistura com seu
        # perfil pessoal). Certificados A1 instalados no Windows (não só no
        # perfil do Chrome) ainda ficam acessíveis via CryptoAPI do SO.
        user_data_dir = r"C:\Users\Canella e Santos\AppData\Local\Temp\claude_conectividade_profile"
        context = p.chromium.launch_persistent_context(
            user_data_dir,
            channel="chrome",
            headless=False,
            viewport={"width": 1400, "height": 900},
        )
        page = context.pages[0] if context.pages else context.new_page()

        print(f"Navegando para {URL_INICIAL} ...")
        page.goto(URL_INICIAL, wait_until="domcontentloaded")
        time.sleep(2)

        print("\n" + "=" * 78)
        print("Navegador aberto. Navegue manualmente (clique em Caixa Postal,")
        print("escolha o certificado, faça login) até a tela que quer mapear.")
        print("Pressione ENTER aqui a qualquer momento para capturar a estrutura")
        print("da página atual. Digite 'sair' + ENTER para encerrar.")
        print("=" * 78)

        while True:
            cmd = input("\n[Enter=capturar / 'sair'=encerrar] > ").strip().lower()
            if cmd == "sair":
                break
            try:
                dump_estrutura(page)
            except Exception as e:
                print(f"Erro ao capturar estrutura: {e}")

        context.close()


if __name__ == "__main__":
    main()
