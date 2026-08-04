r"""
gerar_rodadas.py — monta um loop_rodadas.json com UMA entrada por mês entre
comp-inicial e comp-final, para o mesmo funcionário/empresa.

Por quê uma rodada por mês: a tela de extração do GFIP no Domínio Folha só
processa UMA competência por vez — o campo "até" não existe nessa tela
(ver DomBot_GFIP.preencher_gfip, docstring). Um recálculo de FGTS de vários
anos (ex.: 6 anos = 72 meses) precisa de 72 rodadas separadas, cada uma
virando seu próprio .SFP e protocolo — não um único intervalo.

Uso:
  python gerar_rodadas.py --empresa 25 --empresa-nome "GELO ARTE" ^
      --funcionario "JOSE CARLOS FERREIRA" --data-atraso 20/08/2026 ^
      --comp-inicial 01/2018 --comp-final 12/2023 ^
      --responsavel 3 --estado "Rio de Janeiro" ^
      --base-arrecadacao "Volta Redonda / RJ" ^
      --saida loop_rodadas.json

Flags opcionais: --cod-recolhimento, --com-backup (default: true).
Se --saida já existir, PEDE confirmação antes de sobrescrever (a menos que
--forcar seja passado).
"""

import argparse
import json
import os
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def _gerar_competencias(comp_inicial: str, comp_final: str):
    mi, ai = (int(x) for x in comp_inicial.split("/"))
    mf, af = (int(x) for x in comp_final.split("/"))
    if (af, mf) < (ai, mi):
        raise ValueError(f"comp-final ({comp_final}) é anterior a comp-inicial ({comp_inicial})")
    m, a = mi, ai
    while (a, m) <= (af, mf):
        yield f"{m:02d}/{a:04d}"
        m += 1
        if m == 13:
            m = 1
            a += 1


def main():
    p = argparse.ArgumentParser(description="Gera loop_rodadas.json com uma rodada por mês, mesmo funcionário/empresa")
    p.add_argument("--empresa", required=True)
    p.add_argument("--empresa-nome", dest="empresa_nome", required=True)
    p.add_argument("--funcionario", nargs="+", required=True,
                   help="Nome(s)/CPF(s) do(s) funcionário(s) — mesmo(s) em todas as rodadas geradas")
    p.add_argument("--data-atraso", dest="data_atraso", required=True,
                   help="DD/MM/AAAA — mesma data de atraso em todas as rodadas")
    p.add_argument("--comp-inicial", dest="comp_inicial", required=True, help="MM/AAAA")
    p.add_argument("--comp-final", dest="comp_final", required=True, help="MM/AAAA")
    p.add_argument("--responsavel", default="1")
    p.add_argument("--cod-recolhimento", dest="cod_recolhimento", default=None)
    p.add_argument("--estado", required=True)
    p.add_argument("--base-arrecadacao", dest="base_arrecadacao", required=True)
    p.add_argument("--com-backup", dest="com_backup", action="store_true", default=True)
    p.add_argument("--sem-backup", dest="com_backup", action="store_false")
    p.add_argument("--saida", default="loop_rodadas.json")
    p.add_argument("--forcar", action="store_true", help="Sobrescreve --saida sem perguntar")
    args = p.parse_args()

    competencias = list(_gerar_competencias(args.comp_inicial, args.comp_final))

    rodadas = []
    for comp in competencias:
        rodada = {
            "empresa": args.empresa,
            "empresa_nome": args.empresa_nome,
            "comp_inicial": comp,
            "comp_final": comp,
            "data_atraso": args.data_atraso,
            "responsavel": args.responsavel,
            "funcionario": args.funcionario,
            "nome_mensagem": f"GFIP {args.empresa} - {comp.replace('/', '-')}",
            "estado": args.estado,
            "base_arrecadacao": args.base_arrecadacao,
            "com_backup": args.com_backup,
        }
        if args.cod_recolhimento:
            rodada["cod_recolhimento"] = args.cod_recolhimento
        rodadas.append(rodada)

    if os.path.exists(args.saida) and not args.forcar:
        resp = input(f"⚠️ '{args.saida}' já existe ({os.path.getsize(args.saida)} bytes). "
                     f"Sobrescrever com {len(rodadas)} rodada(s)? [s/N] ").strip().lower()
        if resp != "s":
            print("Cancelado — nada foi escrito.")
            sys.exit(1)

    with open(args.saida, "w", encoding="utf-8") as f:
        json.dump(rodadas, f, ensure_ascii=False, indent=2)

    print(f"✅ {len(rodadas)} rodada(s) escrita(s) em '{args.saida}' "
          f"({competencias[0]} até {competencias[-1]}).")


if __name__ == "__main__":
    main()
