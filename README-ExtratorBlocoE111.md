# SPED-ExtratorBlocoE111

Extrator do **Bloco E** do SPED Fiscal (EFD ICMS/IPI): lê os ajustes da apuração
do ICMS (registro **E111**) e seus registros-filho (**E112** e **E113**),
identifica o posto pelo CNPJ e gera dois modelos de planilha Excel.

## O que faz

- Pergunta, em janelas de seleção: a pasta-raiz com os SPED (`.txt`), o arquivo
  de parâmetros `PostosR7.csv` e a pasta de saída
- Varre a pasta-raiz **e todas as subpastas**, processando em paralelo (um
  processo por núcleo da CPU) — dimensionado para milhares de arquivos
- Lê o registro `0000` (empresa, CNPJ, IE, UF, período) e percorre o Bloco E
  coletando cada `E111` com seus filhos `E112`/`E113`
- Identifica o posto cruzando o CNPJ do declarante com `PostosR7.csv`
  (`Cod;Posto;CNPJ`)
- Deduplica por **CNPJ + período**, mantendo a **retificadora mais recente**
  (maior nº de sequência no nome do arquivo; empate → data de modificação)
- Detecta o encoding automaticamente (UTF-8 ou Latin-1)
- Nunca para por causa de um arquivo com problema: registra no log e continua

## Saídas (na pasta escolhida)

| Arquivo | Conteúdo |
|---------|----------|
| `E111_Consolidado.xlsx` | Abas **Resumo** (totais gerais e por posto, com fórmulas SUMIFS/COUNTIFS), **E111** (1 linha por ajuste), **E112** e **E113** (ligadas ao E111 pela coluna `ID_E111`) |
| `E111_por_Posto.xlsx` | Aba **Índice** + **1 aba por posto**; dentro de cada aba os dados são agrupados por período e, em cada período, os blocos **E111**, **E112** e **E113** na sequência, com "Total do período" |
| `log_E111.csv` | Auditoria: 1 linha por arquivo (empresa, período, qtd de E111, status/erro) |

Valores em formato numérico (`#,##0.00`) e datas em `DD/MM/AAAA`.

## Requisitos

- Windows com Python 3.x
- Biblioteca **openpyxl** (`pip install openpyxl`) — o `.bat` instala automaticamente se faltar

## Como usar

1. Dê duplo clique em **`executar_ExtratorBlocoE111.bat`**
2. Selecione, nas três janelas: a pasta dos SPED, o `PostosR7.csv` e a pasta de saída
3. Ao final, abra os arquivos `E111_Consolidado.xlsx` e `E111_por_Posto.xlsx`

Por linha de comando:

```
python SPED-ExtratorBlocoE111.py  <pasta_speds>  <PostosR7.csv>  <pasta_saida>
```

## Versionamento

Use **`versionar_ExtratorBlocoE111.bat`** para commitar e marcar a versão
(`extrator-e111-v1.0.0`). Versão atual: **1.0.0**.
