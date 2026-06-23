# SPED — Notas e Itens de Entrada por fornecedor

Variante do conversor que, a partir de uma **tabela de fornecedores selecionados**, extrai do SPED Fiscal as **notas de entrada (C100)** e os **itens/produtos (C170)** desses fornecedores, identifica o posto e confere a soma dos itens contra o total da nota.

## O que faz

- Pede: a pasta-raiz dos SPED (`.txt`), a **tabela de fornecedores** (`.csv` com coluna `CNPJBase` ou `CNPJ`), a pasta de saída e, se necessário, a tabela `PostosR7.csv`.
- Casa o fornecedor pela **raiz do CNPJ (8 dígitos)** — pega todas as filiais.
- Lê notas `C100` (entrada, NF-e, ignorando canceladas/denegadas/inutilizadas) e seus itens `C170`; a descrição do produto vem do cadastro `0200`.
- Identifica **Cod/Posto** pelo CNPJ do declarante (`PostosR7.csv`).
- **Conferência:** soma o `VL_ITEM` dos itens e compara com o `VL_DOC` da nota → `OK` / `DIVERGE` / `SEM ITENS`.
- Processamento em paralelo; deduplica por CNPJ + chave (retificadora mais recente).

## Saídas (na pasta escolhida)

| Arquivo | Conteúdo |
|---------|----------|
| `notas_entrada.csv` | Uma linha por nota (Cod, Posto, soma dos itens, diferença, confere) |
| `itens_entrada.csv` | Uma linha por item/produto |
| `entrada_fornecedores.xlsx` | Guias **Notas_Entrada** e **Itens_Entrada** |
| `log_processamento.csv` | Auditoria por arquivo |

## Requisitos

- Windows com Python 3.x. Para o XLSX: **openpyxl** (o `.bat` instala `openpyxl pandas lxml xlsxwriter`).

## Como usar

1. Duplo clique em **`executar_forn_item.bat`**.
2. Selecione: pasta dos SPED, tabela de fornecedores, pasta de saída (e `PostosR7.csv` se pedido).

Por linha de comando:

```
python sped_para_csv_forn_Item.py <pasta_speds> <Fornecedores.csv> [pasta_saida] [PostosR7.csv]
```

Versão atual: **2.1.0**.
