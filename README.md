# DomBot GFIP

Automação RPA para exportação do relatório GFIP no sistema Domínio Folha.

## O que faz

Para cada empresa listada numa planilha Excel, o bot:

1. Conecta na janela já aberta do Domínio Folha (o sistema precisa estar aberto antes de iniciar).
2. Troca para a empresa (`F8` + código da empresa).
3. Abre o relatório GFIP via menu: `Alt+R` (Relatórios) → `I` (Informativos) → `→` → `→` → `Enter`.
4. Preenche, via TAB, os campos da janela na ordem: Competência de → Competência até → Código de Recolhimento → Caminho de destino.
5. Confirma (`Enter`) — o Domínio grava o arquivo TXT diretamente na pasta informada, sem diálogo "Salvar como".
6. Fecha as janelas/abas abertas e segue para a próxima empresa.

Como o Domínio grava o GFIP sempre com um **nome de arquivo fixo** (não é possível renomear), cada empresa é salva em sua própria subpasta:

```
{pasta_de_saida}/{numero_da_empresa}/
```

Casos especiais tratados sem contar como erro:
- **Sem dados para emitir** — empresa sem movimento no período (`sem_dados`).
- **Bloqueada pelo módulo Honorários** — empresa ignorada (`honorarios`).

## Planilha de entrada

Arquivo Excel com as colunas:

| Nº   | Empresa                      | Cod_Recolhimento |
|------|-------------------------------|------------------|
| 1028 | SABOR E LENHA PIZZARIA LTDA  | 115              |
| 1042 | EMPRESA EXEMPLO LTDA          | 115              |

Ao rodar sem argumentos, o script gera um exemplo (`empresas_gfip_exemplo.xlsx`) na própria pasta caso ele não exista.

## Modos de execução

### GUI (padrão)

```
python DomBot_GFIP.py
```

Abre a interface gráfica: seleção da planilha, linha inicial, competências, pasta de saída, logs em tempo real e preview da planilha.

### CLI (uma empresa por vez, para uso via subprocess/servidor)

```
python DomBot_GFIP.py --empresa 1028 --empresa-nome "SABOR E LENHA PIZZARIA LTDA" ^
  --comp-inicial 01/2025 --comp-final 01/2025 --cod-recolhimento 115 --dir-saida C:\saida
```

Emite uma linha JSON por evento no stdout (`{"tipo": ..., "msg": ..., "ts": ...}`), pensado para ser consumido por um servidor via WebSocket. Sai com código `0` em sucesso e `1` em falha.

### Worker (consulta um servidor por pendências em loop)

```
python DomBot_GFIP.py --worker --servidor http://localhost:8000 --intervalo 10
```

Consulta `GET /api/bot/pendentes`, processa um item por vez e reporta o resultado via `POST /api/bot/{execucao_id}/iniciar` / `.../finalizar` / `.../log`.

## Requisitos

- Windows, com o Domínio Folha já aberto e logado.
- Python 3 com: `pywinauto`, `pywin32` (`win32gui`, `win32con`, `win32api`).
- Para o modo GUI: `customtkinter`, `pandas`, `Pillow`, `openpyxl`.

## Logs

Logs de sucesso/erro são gravados em `logs/gfip_success_AAAA-MM-DD.log` e `logs/gfip_error_AAAA-MM-DD.log`, além do painel de logs da GUI (exportável para `.txt`).

## Pontos a validar contra o sistema real

Como o script foi adaptado a partir do `DomBot_DIRF.py` sem acesso direto à janela do GFIP no Domínio, os pontos abaixo são suposições que podem exigir ajuste após o primeiro teste real:

1. **Texto de identificação da janela** — hoje o bot procura por um título contendo `"GFIP"`. Se o título real for diferente, ajustar `title_re` e `_find_window` em `abrir_extrator_gfip` / `preencher_gfip` / `salvar_gfip` / `_fechar_todas_abas`.
2. **Ordem dos campos via TAB** — assume-se a sequência Competência de → até → Código de Recolhimento → Caminho, sem campos extras. Se houver algum campo adicional na tela, a sequência de `{TAB}` em `preencher_gfip` precisa ser corrigida.
3. **Botão de confirmação** — hoje usa `{ENTER}` simples. Se o botão OK exigir clique direto (como no DIRF, via `ctrl_id=1000`), trocar por `win32gui.SendMessage`/`PostMessage` com o ID correto.

Recomenda-se rodar primeiro em modo CLI contra uma empresa de teste, observando os logs, antes de usar em produção.
