"""Compare real output files with generated results column by column."""
import csv
from pathlib import Path

PAIRS = [
    ("CATR", "load/ESPF_CATR_SicReport.csv", "results/frbdfruba_catr.csv"),
    ("ENGAG", "load/ESPF_engag_SicReport.csv", "results/frbdfruba_engag.csv"),
]


def read_csv(path: str):
    with open(path, "r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        return reader.fieldnames, list(reader)


for label, real_path, gen_path in PAIRS:
    print(f"\n{'='*100}")
    print(f"COMPARING {label}")
    print(f"  Real: {real_path}")
    print(f"  Generated: {gen_path}")

    real_cols, real_rows = read_csv(real_path)
    gen_cols, gen_rows = read_csv(gen_path)

    print(f"\n  Row count: real={len(real_rows)} generated={len(gen_rows)}")

    # Column comparison
    real_set = set(real_cols)
    gen_set = set(gen_cols)
    print(f"\n  Columns real: {len(real_cols)}")
    print(f"  Columns generated: {len(gen_cols)}")
    print(f"  In real but NOT in generated: {sorted(real_set - gen_set)}")
    print(f"  In generated but NOT in real: {sorted(gen_set - real_set)}")
    print(f"  Common columns: {len(real_set & gen_set)}")

    # Value comparison for first N rows on common columns
    common_cols = [c for c in real_cols if c in gen_set]
    n = min(5, len(real_rows), len(gen_rows))
    diffs_summary = {}
    for i in range(n):
        rr = real_rows[i]
        gr = gen_rows[i]
        for col in common_cols:
            rv = rr.get(col, "")
            gv = gr.get(col, "")
            if rv != gv:
                if col not in diffs_summary:
                    diffs_summary[col] = []
                diffs_summary[col].append((i + 1, rv, gv))

    if not diffs_summary:
        print(f"\n  First {n} rows: ALL MATCH on common columns")
    else:
        print(f"\n  VALUE DIFFERENCES (first {n} rows, common columns):")
        for col, examples in sorted(diffs_summary.items()):
            print(f"\n    Column: {col}")
            for row_num, rv, gv in examples:
                print(f"      row {row_num}: real={repr(rv)[:80]}  gen={repr(gv)[:80]}")

    # Column order comparison
    real_order = list(real_cols)
    gen_order = [c for c in gen_cols if c in real_set]
    real_filtered = [c for c in real_order if c in gen_set]
    if real_filtered == gen_order:
        print(f"\n  Column order: MATCHES (for common columns)")
    else:
        print(f"\n  Column order MISMATCH:")
        print(f"    Real order:      {real_filtered}")
        print(f"    Generated order: {gen_order}")
