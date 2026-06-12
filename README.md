# sped-para-csv

Extrator de notas fiscais de entrada de arquivos SPED Fiscal (EFD ICMS/IPI), consolidadas em CSV.

## O que faz

- Pergunta, em janelas de seleção, a pasta-raiz com os arquivos SPED (`.txt`) e a pasta onde salvar os CSVs
- Varre a pasta-raiz **e todas as subpastas**, processando os arquivos em paralelo (um processo por núcleo da CPU) — dimensionado para milhares de arquivos
- Extrai as notas de entrada: registro `C100` com `IND_OPER=0` e `COD_MOD=55` (NF-e), ignorando canceladas/denegadas/inutilizadas (`COD_SIT` 02, 03, 04, 05)
- Deduplica por CNPJ do declarante + chave da NF-e, mantendo o arquivo mais recente (cobre SPED original + retificadora)
- Nunca interrompe por causa de um arquivo com problema: registra o erro no log e continua
- Detecta automaticamente o encoding de cada arquivo (UTF-8 ou Latin-1)

## Saídas

| Arquivo | Conteúdo |
|---------|----------|
| `notas_entrada.csv` | Uma linha por nota: Razão Social (SPED), CNPJ da Sociedade (SPED), CNPJ do Fornecedor, Razão Social do Fornecedor, Data de Emissão, Chave XML, Valor Total da Nota |
| `log_processamento.csv` | Uma linha por arquivo processado: empresa, período, nº de notas, status/erro |

CSV no padrão Excel-BR: separador `;`, UTF-8 com BOM, decimal com vírgula e chave de 44 dígitos protegida contra notação científica.

## Requisitos

- Windows com Python 3.x instalado (somente biblioteca padrão — sem dependências externas)

## Como usar

1. Dê duplo clique em **`executar_sped_csv.bat`**
2. Na primeira janela, selecione a pasta-raiz com os arquivos SPED
3. Na segunda janela, selecione a pasta onde salvar os CSVs
4. Acompanhe o progresso no terminal

Ou pelo terminal:

```bat
python sped_para_csv.py                    :: abre as janelas de seleção
python sped_para_csv.py C:\speds           :: saída na própria pasta de entrada
python sped_para_csv.py C:\speds C:\saida  :: entrada e saída indicadas
```

## Estrutura de arquivos

```
SPEDNFE/
├── executar_sped_csv.bat   # Script principal — duplo clique
├── sped_para_csv.py        # Extrator SPED → CSV
├── versionar.bat           # Commit + push da versão atual
└── *.txt                   # Arquivos SPED (não versionados)
```

## Versão

**v1.0.0** — junho/2026

- v1.0.0: versão inicial — extração paralela de notas de entrada com deduplicação e log de auditoria
