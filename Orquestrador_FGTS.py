r"""
Orquestrador — Recálculo de FGTS (fluxo completo, 6 etapas)
============================================================
Encadeia os 4 bots do projeto num único processo, dividido em 2 fases
por causa de uma pausa manual inevitável: a Conectividade Social exige
login por certificado digital, que só o usuário pode fazer (o navegador
mostra o popup nativo do Windows para escolher o certificado).

FASE 1 (--fase1): gera o arquivo, processa no SEFIP e envia à Conectividade.
    1. DomBot_GFIP.py     — gera o .re/.txt do GFIP (Domínio Folha).
    2. Bot_Sefip.py --tudo — importa, separa o funcionário na Modalidades,
       executa o movimento e gera o .SFP.
    3. Bot_Conectividade.py --confirmar-envio — sobe o .SFP, envia (o
       usuário só precisa escolher o certificado no popup do Windows/
       Chrome quando ele aparecer) e baixa os artefatos do protocolo em
       <pasta_empresa>/protocolo/.

FASE 2 (--fase2): valida o protocolo de volta no SEFIP e fecha o ciclo.
    4. Bot_Sefip_Protocolo.py --validar-protocolo — Relatórios > GRF,
       valida o XML de protocolo (localizado automaticamente na pasta),
       gera a guia em PDF.
    5. (opcional, --com-backup) Bot_Sefip_Protocolo.py --backup.

Por que 2 fases e não um script contínuo: entre a Fase 1 e a Fase 2 o
usuário revisa o envio na Conectividade (mesmo com --confirmar-envio, o
clique em Enviar e a eventual seleção de certificado são pontos de
atenção do usuário) — rodar a Fase 2 automaticamente em seguida arriscaria
validar um protocolo que ainda não terminou de ser processado pela Caixa.
Rodar as fases como 2 comandos separados dá ao usuário o controle de
quando prosseguir.

ORGANIZAÇÃO DE PASTAS — por competência, não só por empresa:
  Uma mesma empresa pode passar pelo recálculo de FGTS várias vezes, uma
  por competência (às vezes um funcionário só, às vezes vários) — os
  funcionários entram como parâmetro de cada rodada (--funcionario), não
  como parte do caminho. Sem separar por competência, uma rodada nova
  sobrescreveria os arquivos (.SFP, XML de protocolo, PDF da guia) da
  rodada anterior da mesma empresa. Por isso o orquestrador recebe a
  pasta RAIZ (--dir-raiz, ex.: .\results) e monta sozinho:
      {dir-raiz}/{empresa}/{MM-AAAA}/
  (ou {MM-AAAA_inicial}_a_{MM-AAAA_final} se for um intervalo de
  competências) — mesma convenção usada pelo DomBot_GFIP.py.

PRÉ-REQUISITOS (iguais aos dos bots individuais):
  - SEFIP aberto na empresa/competência certa antes da Fase 1 (nenhum bot
    troca de empresa sozinho).
  - Sessão elevada (Administrador) — obrigatório para Bot_Sefip.py e
    Bot_Sefip_Protocolo.py (SEFIP roda elevado).
  - Data de atraso dentro da vigência do edital de índices vigente no
    SEFIP (o próprio SEFIP informa a janela no popup de erro, se errar).

Exemplos:
  # Fase 1 completa, do zero (grava em .\results\404\11-2022\):
  python Orquestrador_FGTS.py --fase1 --dir-raiz ".\results" ^
      --empresa 404 --empresa-nome "N A SCARAMUSSA" ^
      --comp-inicial 11/2022 --comp-final 11/2022 --data-atraso 20/08/2026 ^
      --funcionario "ALTAIR MACHADO" --estado "Rio de Janeiro" ^
      --base-arrecadacao "Volta Redonda / RJ"

  # Fase 2, depois que a Conectividade terminou de processar:
  python Orquestrador_FGTS.py --fase2 --dir-raiz ".\results" --empresa 404 ^
      --comp-inicial 11/2022 --comp-final 11/2022 --com-backup

Autor: Hugo L. Almeida
"""

import argparse
import glob
import json
import os
import subprocess
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

DIR_SCRIPT = os.path.dirname(os.path.abspath(__file__))


def _nome_pasta_competencia(comp_inicial: str, comp_final: str) -> str:
    """Mesma convenção do DomBot_GFIP.py — ver _nome_pasta_competencia lá."""
    ini = (comp_inicial or "").replace("/", "-")
    fim = (comp_final or "").replace("/", "-")
    if not ini:
        return "sem-competencia"
    if fim and fim != ini:
        return f"{ini}_a_{fim}"
    return ini


def _rodar(args_cli, descricao):
    """
    Executa um dos bots como subprocesso, com stdout/stderr herdados (o
    usuário acompanha o log de cada bot em tempo real, igual a rodá-los
    manualmente). Levanta SystemExit se o bot retornar código != 0 — o
    orquestrador para no primeiro erro, não tenta adivinhar como seguir.
    """
    print("=" * 78)
    print(f"▶ {descricao}")
    print("   " + " ".join(f'"{a}"' if " " in a else a for a in args_cli))
    print("=" * 78)
    resultado = subprocess.run([sys.executable] + args_cli, cwd=DIR_SCRIPT)
    if resultado.returncode != 0:
        print(f"\n❌ '{descricao}' falhou (código {resultado.returncode}). "
              "Orquestrador interrompido — corrija manualmente e rode de novo.")
        sys.exit(resultado.returncode)
    print(f"✅ {descricao} — concluído.\n")


def _pasta_empresa_competencia(args) -> str:
    return os.path.join(args.dir_raiz, args.empresa,
                        _nome_pasta_competencia(args.comp_inicial, args.comp_final))


def fase1(args):
    print("\n########## FASE 1 — Gerar arquivo, processar no SEFIP e enviar ##########\n")

    dir_saida = _pasta_empresa_competencia(args)
    print(f"📁 Pasta desta rodada (empresa {args.empresa}, competência "
          f"{args.comp_inicial}"
          f"{' a ' + args.comp_final if args.comp_final != args.comp_inicial else ''}): "
          f"{dir_saida}\n")

    _rodar([
        os.path.join(DIR_SCRIPT, "DomBot_GFIP.py"),
        "--empresa", args.empresa,
        "--empresa-nome", args.empresa_nome,
        "--comp-inicial", args.comp_inicial,
        "--comp-final", args.comp_final,
        "--data-atraso", args.data_atraso,
        "--dir-saida", args.dir_raiz,
    ] + (["--cod-recolhimento", args.cod_recolhimento] if args.cod_recolhimento else [])
      + (["--responsavel", args.responsavel] if args.responsavel else []),
        "1/5 — DomBot_GFIP (gera o .re do GFIP)")

    caminho_re = os.path.join(dir_saida, "Sefip.re")
    if not os.path.exists(caminho_re):
        print(f"❌ Esperava encontrar '{caminho_re}' após o DomBot_GFIP, mas não achei. "
              "Confira o nome do arquivo gerado e ajuste manualmente antes de prosseguir.")
        sys.exit(1)

    _rodar([
        os.path.join(DIR_SCRIPT, "Bot_Sefip.py"),
        "--tudo",
        "--arquivo", caminho_re,
        "--funcionario", *args.funcionario,
        "--salvar",
        "--dir-saida", dir_saida,
    ], "2/5 — Bot_Sefip (Etapas 1-4: importar, Modalidades, Executar, gerar .SFP)")

    sfps = sorted(glob.glob(os.path.join(dir_saida, "*.SFP")) +
                  glob.glob(os.path.join(dir_saida, "*.sfp")),
                  key=os.path.getmtime, reverse=True)
    if not sfps:
        print(f"❌ Nenhum .SFP encontrado em '{dir_saida}' após o Bot_Sefip. "
              "Confira a Etapa 4 manualmente antes de prosseguir.")
        sys.exit(1)
    caminho_sfp = sfps[0]
    print(f"📎 Arquivo .SFP a enviar: {caminho_sfp}")

    _rodar([
        os.path.join(DIR_SCRIPT, "Bot_Conectividade.py"),
        "--arquivo", caminho_sfp,
        "--nome-mensagem", args.nome_mensagem,
        "--estado", args.estado,
        "--base-arrecadacao", args.base_arrecadacao,
        "--confirmar-envio",
    ], "3/5 — Bot_Conectividade (envio à Conectividade Social — selecione o "
       "certificado digital quando o popup do Windows/Chrome aparecer)")

    print("\n########## FASE 1 concluída. ##########")
    print(f"Artefatos do protocolo devem estar em: {os.path.join(dir_saida, 'protocolo')}")
    print("Quando quiser fechar o ciclo, rode a Fase 2:")
    print(f'  python Orquestrador_FGTS.py --fase2 --dir-raiz "{args.dir_raiz}" '
          f'--empresa {args.empresa} --comp-inicial {args.comp_inicial} '
          f'--comp-final {args.comp_final}'
          + (" --com-backup" if args.com_backup else ""))


def fase2(args):
    print("\n########## FASE 2 — Validar protocolo no SEFIP e gerar guia ##########\n")

    dir_saida = _pasta_empresa_competencia(args)
    print(f"📁 Pasta desta rodada: {dir_saida}\n")

    pasta_protocolo = os.path.join(dir_saida, "protocolo")
    xmls = sorted(glob.glob(os.path.join(pasta_protocolo, "*.xml")),
                  key=os.path.getmtime, reverse=True)
    if not xmls:
        print(f"❌ Nenhum XML de protocolo encontrado em '{pasta_protocolo}'. "
              "A Fase 1 (Etapa 5 — Conectividade) já terminou de fato?")
        sys.exit(1)
    caminho_xml = xmls[0]

    sfps = sorted(glob.glob(os.path.join(dir_saida, "*.SFP")) +
                  glob.glob(os.path.join(dir_saida, "*.sfp")),
                  key=os.path.getmtime, reverse=True)
    if not sfps:
        print(f"❌ Nenhum .SFP encontrado em '{dir_saida}'.")
        sys.exit(1)
    caminho_sfp = sfps[0]

    manifesto = {}
    caminho_manifesto = os.path.join(pasta_protocolo, "protocolo.json")
    if os.path.exists(caminho_manifesto):
        try:
            with open(caminho_manifesto, encoding="utf-8") as f:
                manifesto = json.load(f)
        except Exception:
            pass

    nome_pdf = args.nome_pdf or f"Guia_{os.path.splitext(os.path.basename(caminho_sfp))[0]}.pdf"
    caminho_pdf = os.path.join(dir_saida, nome_pdf)

    print(f"XML do protocolo : {caminho_xml}")
    print(f"Arquivo .SFP     : {caminho_sfp}")
    print(f"Protocolo        : {manifesto.get('numero_protocolo', '(sem manifesto)')}")
    print(f"PDF de saída     : {caminho_pdf}\n")

    _rodar([
        os.path.join(DIR_SCRIPT, "Bot_Sefip_Protocolo.py"),
        "--validar-protocolo",
        "--xml", caminho_xml,
        "--sfp", caminho_sfp,
        "--pdf-saida", caminho_pdf,
    ], "4/5 — Bot_Sefip_Protocolo (Relatórios > GRF: valida protocolo, gera guia em PDF)")

    if args.com_backup:
        caminho_backup = args.dir_backup or os.path.join(dir_saida, "backup_sefip")
        _rodar([
            os.path.join(DIR_SCRIPT, "Bot_Sefip_Protocolo.py"),
            "--backup",
            "--dir-backup", caminho_backup,
        ], "5/5 — Bot_Sefip_Protocolo (Ferramentas > Fazer backup)")

    print("\n########## FASE 2 concluída — ciclo do recálculo de FGTS fechado. ##########")
    print("Lembrete: 'Ferramentas > Limpar Base de Dados' é uma ação irreversível e "
          "NÃO é rodada automaticamente pelo orquestrador — use "
          "'python Bot_Sefip_Protocolo.py --limpar-base --confirmar-limpeza' manualmente "
          "quando tiver certeza de que quer resetar o SEFIP para o próximo processo.")


def main():
    p = argparse.ArgumentParser(
        description="Orquestrador do recálculo de FGTS (DomBot_GFIP → Bot_Sefip → "
                     "Bot_Conectividade → Bot_Sefip_Protocolo)")
    fase_grupo = p.add_mutually_exclusive_group(required=True)
    fase_grupo.add_argument("--fase1", action="store_true",
                             help="Gera o arquivo, processa no SEFIP e envia à Conectividade")
    fase_grupo.add_argument("--fase2", action="store_true",
                             help="Valida o protocolo de volta no SEFIP e gera a guia")

    p.add_argument("--dir-raiz", dest="dir_raiz", required=True,
                   help="Pasta raiz de saída (ex.: .\\results) — o orquestrador monta "
                        "sozinho {dir-raiz}\\{empresa}\\{MM-AAAA}\\ dentro dela")
    p.add_argument("--empresa", required=True, help="Número da empresa")
    p.add_argument("--comp-inicial", dest="comp_inicial", required=True,
                   help="Competência inicial MM/AAAA")
    p.add_argument("--comp-final", dest="comp_final", required=True,
                   help="Competência final MM/AAAA (igual à inicial se for uma só)")

    # Fase 1
    p.add_argument("--empresa-nome", dest="empresa_nome", help="[Fase 1] Nome da empresa")
    p.add_argument("--cod-recolhimento", dest="cod_recolhimento",
                   help="[Fase 1] Código de Recolhimento (opcional)")
    p.add_argument("--data-atraso", dest="data_atraso",
                   help="[Fase 1] DD/MM/AAAA dentro da vigência do edital de índices vigente")
    p.add_argument("--responsavel", dest="responsavel", default="1",
                   help="[Fase 1] Código do responsável a preencher na extração do GFIP "
                        "(campo PBEDIT190 do Domínio). PRECISA bater com o CNPJ do certificado "
                        "usado depois na Conectividade Social — valor errado gera .SFP recusado "
                        "no envio ('inscrição do responsável diferente da empresa logada'). "
                        "Padrão: '1'.")
    p.add_argument("--funcionario", nargs="+",
                   help="[Fase 1] Nome(s)/CPF(s) do(s) funcionário(s) a separar na Modalidades "
                        "(essa rodada da competência é só para esses funcionários)")
    p.add_argument("--nome-mensagem", dest="nome_mensagem",
                   help="[Fase 1] Texto do campo 'Nome da Mensagem' na Conectividade")
    p.add_argument("--estado", help="[Fase 1] Estado na Conectividade (ex.: 'Rio de Janeiro')")
    p.add_argument("--base-arrecadacao", dest="base_arrecadacao",
                   help="[Fase 1] Base de Arrecadação (ex.: 'Volta Redonda / RJ')")

    # Fase 2
    p.add_argument("--nome-pdf", dest="nome_pdf",
                   help="[Fase 2] Nome do PDF da guia (padrão: Guia_<NRA>.pdf)")
    p.add_argument("--com-backup", dest="com_backup", action="store_true",
                   help="[Fase 2] Também roda Ferramentas > Fazer backup ao final")
    p.add_argument("--dir-backup", dest="dir_backup",
                   help="[Fase 2] Caminho do arquivo de backup (padrão: <pasta da rodada>\\backup_sefip)")

    args = p.parse_args()

    # Sempre absolutizar dir_raiz: os bots chamados (DomBot_GFIP, Bot_Sefip,
    # etc.) rodam como processos Win32 com diretório de trabalho próprio
    # (ex.: o Domínio Folha não necessariamente compartilha o cwd do
    # orquestrador) — um caminho relativo tipo ".\results_teste" passado
    # a eles resolve para o lugar ERRADO. Testado na prática (2026-07-28):
    # isso causou "é necessário informar uma pasta de destino válida" no
    # Domínio, seguido de falso positivo "sem dados para gerar a GFIP".
    args.dir_raiz = os.path.abspath(args.dir_raiz)

    if args.fase1:
        obrigatorios = ["empresa_nome", "data_atraso", "funcionario",
                         "nome_mensagem", "estado", "base_arrecadacao"]
        faltando = [f"--{o.replace('_', '-')}" for o in obrigatorios if not getattr(args, o)]
        if faltando:
            print(f"❌ --fase1 requer também: {', '.join(faltando)}")
            sys.exit(1)
        fase1(args)
    else:
        fase2(args)


if __name__ == "__main__":
    main()
