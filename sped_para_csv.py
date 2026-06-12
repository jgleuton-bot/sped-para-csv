"""
Extrator de Notas de Entrada de SPED Fiscal (EFD ICMS/IPI) para CSV  —  v1.0.0
===============================================================================
Varre uma pasta (e subpastas) com arquivos SPED .txt, extrai as notas fiscais
de entrada (registro C100, modelo 55) e consolida tudo em um unico CSV.

Saidas (na pasta escolhida):
  - notas_entrada.csv       -> uma linha por nota (deduplicada)
  - log_processamento.csv   -> uma linha por arquivo processado (auditoria)

Colunas do notas_entrada.csv:
  Razao Social (SPED); CNPJ da Sociedade (SPED); CNPJ do Fornecedor;
  Razao Social do Fornecedor; Data de Emissao; Chave XML; Valor Total da Nota

Regras:
  - Considera apenas IND_OPER=0 (entrada) e COD_MOD=55 (NF-e)
  - Ignora notas canceladas/denegadas/inutilizadas (COD_SIT 02, 03, 04, 05)
  - Deduplica por (CNPJ declarante + chave NF-e), mantendo o arquivo mais
    recente (data de modificacao) — cobre SPED original + retificadora
  - Nunca para por causa de um arquivo com problema: registra o erro no log
    e continua

Uso:
    python sped_para_csv.py                    -> abre janelas para escolher as pastas
    python sped_para_csv.py C:/speds           -> entrada indicada, saida na mesma pasta
    python sped_para_csv.py C:/speds C:/saida  -> entrada e saida indicadas
"""

__version__ = "1.0.0"

import csv
import io
import os
import sys
import time
from concurrent.futures import ProcessPoolExecutor, as_completed
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

CODSIT_EXCLUIR = {"02", "03", "04", "05"}  # canceladas, denegadas, inutilizadas


def fmt_cnpj(c):
    c = (c or "").strip()
    if len(c) == 14:
        return f"{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:]}"
    if len(c) == 11:
        return f"{c[:3]}.{c[3:6]}.{c[6:9]}-{c[9:]}"
    return c


def fmt_data(d):
    d = (d or "").strip()
    return f"{d[0:2]}/{d[2:4]}/{d[4:8]}" if len(d) == 8 else d


def fmt_valor(v):
    """'19040,8' ou '53855' -> '19040,80' / '53855,00' (decimal com virgula)."""
    v = (v or "").strip()
    if not v:
        return "0,00"
    try:
        f = float(v.replace(".", "").replace(",", ".")) if "," in v else float(v)
        return f"{f:.2f}".replace(".", ",")
    except ValueError:
        return v


def processar_arquivo(caminho_str):
    """Processa um arquivo SPED. Roda em processo separado (worker)."""
    p = Path(caminho_str)
    res = {
        "arquivo": str(p),
        "mtime": 0.0,
        "empresa": "",
        "cnpj": "",
        "periodo": "",
        "notas": [],
        "status": "",
    }
    try:
        res["mtime"] = p.stat().st_mtime
        raw = p.read_bytes()
        try:
            texto = raw.decode("utf-8")
        except UnicodeDecodeError:
            texto = raw.decode("latin-1")

        if not texto.startswith("|0000|"):
            res["status"] = "ignorado: nao e arquivo SPED"
            return res

        parts = {}
        pendentes = []  # C100 pode aparecer antes? (0150 vem antes no leiaute)
        for ln in texto.splitlines():
            if not ln.startswith("|"):
                continue
            f = ln.split("|")
            reg = f[1]
            if reg == "0000":
                # |0000|COD_VER|COD_FIN|DT_INI|DT_FIN|NOME|CNPJ|...
                res["empresa"] = f[6].strip()
                res["cnpj"] = f[7].strip()
                res["periodo"] = f"{fmt_data(f[4])} a {fmt_data(f[5])}"
            elif reg == "0150":
                # |0150|COD_PART|NOME|COD_PAIS|CNPJ|CPF|IE|...
                if len(f) > 6:
                    parts[f[2].strip()] = (f[3].strip(), f[5].strip() or f[6].strip())
            elif reg == "C100":
                # |C100|IND_OPER|IND_EMIT|COD_PART|COD_MOD|COD_SIT|SER|NUM|CHV|DT_DOC|DT_E_S|VL_DOC|
                if len(f) < 13:
                    continue
                if f[2].strip() != "0" or f[5].strip() != "55":
                    continue
                if f[6].strip() in CODSIT_EXCLUIR:
                    continue
                chave = f[9].strip()
                if len(chave) != 44:
                    continue
                pendentes.append((f[4].strip(), chave, f[10].strip(), f[12].strip()))

        for cod_part, chave, dt_doc, vl in pendentes:
            nome_forn, doc_forn = parts.get(cod_part, ("", ""))
            res["notas"].append(
                [
                    res["empresa"],
                    fmt_cnpj(res["cnpj"]),
                    fmt_cnpj(doc_forn),
                    nome_forn,
                    fmt_data(dt_doc),
                    chave,
                    fmt_valor(vl),
                ]
            )
        res["status"] = "ok"
    except Exception as e:
        res["status"] = f"ERRO: {e}"
    return res


def escolher_pastas():
    """Abre janelas para escolher a pasta dos SPEDs e a pasta de saida."""
    try:
        import tkinter as tk
        from tkinter import filedialog
    except ImportError:
        ent = input("Pasta-raiz com os arquivos SPED (.txt): ").strip(' "')
        sai = input("Pasta para salvar os CSVs: ").strip(' "')
        if not ent or not sai:
            print("Pasta nao informada. Operacao cancelada.")
            sys.exit(1)
        return Path(ent), Path(sai)

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    ent = filedialog.askdirectory(
        parent=root, title="Selecione a pasta-raiz com os arquivos SPED (.txt)"
    )
    if not ent:
        print("Nenhuma pasta selecionada. Operacao cancelada.")
        root.destroy()
        sys.exit(1)
    sai = filedialog.askdirectory(
        parent=root,
        title="Selecione a pasta onde salvar os CSVs gerados",
        initialdir=ent,
    )
    if not sai:
        print("Nenhuma pasta de saida selecionada. Operacao cancelada.")
        root.destroy()
        sys.exit(1)
    root.destroy()
    return Path(ent), Path(sai)


def main():
    if len(sys.argv) > 2:
        pasta_ent, pasta_sai = Path(sys.argv[1]), Path(sys.argv[2])
    elif len(sys.argv) > 1:
        pasta_ent = Path(sys.argv[1])
        pasta_sai = pasta_ent
    else:
        pasta_ent, pasta_sai = escolher_pastas()

    if not pasta_ent.exists():
        print(f"Pasta nao encontrada: {pasta_ent}")
        sys.exit(1)
    pasta_sai.mkdir(parents=True, exist_ok=True)

    print("=" * 70)
    print("  Extrator de Notas de Entrada — SPED Fiscal -> CSV")
    print("=" * 70)
    print(f"  Entrada : {pasta_ent}  (inclui subpastas)")
    print(f"  Saida   : {pasta_sai}")

    arquivos = sorted(str(p) for p in pasta_ent.rglob("*.txt"))
    print(f"  Arquivos .txt encontrados: {len(arquivos)}")
    print("=" * 70 + "\n")
    if not arquivos:
        print("Nenhum arquivo .txt encontrado.")
        sys.exit(0)

    t0 = time.time()
    workers = max(1, (os.cpu_count() or 2) - 1)
    print(f"Processando com {workers} processos em paralelo...\n")

    melhores = {}  # (cnpj declarante, chave) -> (mtime, linha)
    log = []
    feitos = 0
    passo = max(1, len(arquivos) // 50)

    with ProcessPoolExecutor(max_workers=workers) as pool:
        futuros = {pool.submit(processar_arquivo, a): a for a in arquivos}
        for fut in as_completed(futuros):
            r = fut.result()
            log.append(
                [
                    Path(r["arquivo"]).name,
                    r["empresa"],
                    fmt_cnpj(r["cnpj"]),
                    r["periodo"],
                    len(r["notas"]),
                    r["status"],
                ]
            )
            for linha in r["notas"]:
                ch = (r["cnpj"], linha[5])
                atual = melhores.get(ch)
                if atual is None or r["mtime"] > atual[0]:
                    melhores[ch] = (r["mtime"], linha)
            feitos += 1
            if feitos % passo == 0 or feitos == len(arquivos):
                pct = 100 * feitos / len(arquivos)
                print(f"  [{feitos:>6}/{len(arquivos)}]  {pct:5.1f}%  "
                      f"notas unicas: {len(melhores)}")

    linhas = [v[1] for v in melhores.values()]
    # ordena por empresa, data (AAAA-MM-DD), fornecedor
    linhas.sort(key=lambda x: (x[0], x[4][6:10] + x[4][3:5] + x[4][0:2], x[3]))

    csv_notas = pasta_sai / "notas_entrada.csv"
    with open(csv_notas, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(
            [
                "Razao Social (SPED)",
                "CNPJ da Sociedade (SPED)",
                "CNPJ do Fornecedor",
                "Razao Social do Fornecedor",
                "Data de Emissao",
                "Chave XML",
                "Valor Total da Nota",
            ]
        )
        for linha in linhas:
            linha = list(linha)
            linha[5] = f'="{linha[5]}"'  # protege a chave no Excel
            w.writerow(linha)

    csv_log = pasta_sai / "log_processamento.csv"
    log.sort(key=lambda x: x[0])
    with open(csv_log, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["Arquivo", "Razao Social", "CNPJ", "Periodo",
                    "Notas de Entrada", "Status"])
        w.writerows(log)

    erros = sum(1 for l in log if str(l[5]).startswith("ERRO"))
    ignorados = sum(1 for l in log if str(l[5]).startswith("ignorado"))
    total_notas = sum(l[4] for l in log if isinstance(l[4], int))

    print("\n" + "=" * 70)
    print(f"  Concluido em {time.time() - t0:.1f}s")
    print(f"  Arquivos OK: {len(log) - erros - ignorados}   "
          f"Ignorados: {ignorados}   Erros: {erros}")
    print(f"  Notas lidas: {total_notas}   Notas unicas (apos dedup): {len(linhas)}")
    print(f"  -> {csv_notas}")
    print(f"  -> {csv_log}")
    print("=" * 70)


if __name__ == "__main__":
    import multiprocessing

    multiprocessing.freeze_support()
    main()
