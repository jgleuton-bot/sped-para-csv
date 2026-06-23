# SPED-GIA-ConciliaRessarcimento

Concilia o **Ressarcimento de ICMS-ST** entre o **SPED Fiscal** e a **GIA-SP**,
por posto e por período, usando o **Código do Visto Eletrônico** como chave.

## O que faz

- Pede, em três janelas: a pasta (e subpastas) do **SPED** (`.txt`), a pasta (e
  subpastas) da **GIA** (`.prf`) e a pasta de **saída**.
- SPED: lê o registro **E111** com `COD_AJ_APUR = SP020799` (crédito de
  ressarcimento); o visto vem na `DESCR_COMPL_AJ`. Processamento em paralelo.
- GIA (`.prf`): lê o registro tipo **20**, sub-item **007.99**; valor com 3 casas
  decimais (÷1000); visto na descrição após "n.". **Período e CNPJ são lidos do
  registro 05 do próprio arquivo** — não dependem do nome da pasta (há GIAs de um
  ano dentro da pasta de outro). Quando há mais de uma GIA por posto/período
  (backup, substitutiva, retificada), usa a **mais recente**.
- O posto (código + nome) é identificado pela pasta `NNN - NOME` e pelo CNPJ.

## Saídas (na pasta escolhida)

| Arquivo | Conteúdo |
|---------|----------|
| `Conciliacao_SPEDxGIA.xlsx` | Aba **Resumo** (por posto, com situação colorida), aba **Método** e **1 aba por posto** com os blocos por período comparando SPED × GIA e subtotais |
| `log_conciliacao.csv` | Auditoria: 1 linha por arquivo lido (SPED e GIA) |

## Status por lançamento

- **OK** — visto e valor idênticos nos dois.
- **VISTO DIVERGENTE** — o valor confere, mas o Código do Visto Eletrônico difere.
- **VALOR DIVERGENTE** — visto igual, valor diferente.
- **SÓ NO SPED / SÓ NA GIA** — lançamento existe em um lado e não no outro.

Na aba **Resumo**, priorize as situações **VERIFICAR** (diferença de valor/quantidade).

## Requisitos

- Windows com Python 3.x + **openpyxl** (`pip install openpyxl`; o `.bat` instala se faltar).

## Como usar

1. Duplo clique em **`executar_ConciliaRessarcimento.bat`**.
2. Selecione as três pastas (SPED, GIA, saída).
3. Abra o `Conciliacao_SPEDxGIA.xlsx` gerado.

Por linha de comando:

```
python SPED-GIA-ConciliaRessarcimento.py <pasta_sped> <pasta_gia> <pasta_saida>
```

Versão atual: **1.0.0**.
