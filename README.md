# Ferramentas SPED Fiscal

Conjunto de utilitários em Python para extração e conciliação de dados do **SPED Fiscal** (EFD ICMS/IPI) e da **GIA-SP**, voltados a varejo de combustíveis (vários postos).

Cada ferramenta fica na sua própria pasta, com README e launcher próprios:

| Pasta | Ferramenta | O que faz |
|-------|-----------|-----------|
| [`conversor-csv/`](conversor-csv/) | Conversor SPED → CSV | Extrai notas fiscais de entrada (C100/NF-e) e consolida em CSV. |
| [`extrator-e111/`](extrator-e111/) | Extrator do Bloco E | Extrai E111/E112/E113 e gera Excel consolidado e por posto. |
| [`gia-concilia/`](gia-concilia/) | Conciliação SPED × GIA | Concilia o ressarcimento de ICMS-ST (SP020799 × 007.99) por posto/período. |

Cada pasta traz o `.py`, um `README.md` e um `.bat` de execução (duplo clique no Windows; instala `openpyxl` quando necessário).

**Requisitos:** Windows com Python 3.x. As ferramentas de Excel usam `openpyxl`.
