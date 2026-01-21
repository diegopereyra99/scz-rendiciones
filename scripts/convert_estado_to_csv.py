#!/usr/bin/env python3
"""Convierte un estado.json a CSV con transacciones y fila de totales."""

import argparse
import json
from pathlib import Path

import pandas as pd


def load_estado(path: Path) -> dict:
    raw = json.loads(path.read_text(encoding="utf-8"))
    if isinstance(raw, list) and raw:
        raw = raw[0]
    if not isinstance(raw, dict) or "data" not in raw:
        raise ValueError("JSON inesperado: no se encontró la clave 'data'")
    return raw["data"]


def build_transactions_dataframe(data: dict) -> pd.DataFrame:
    cols = ["fecha", "detalle", "importe_origen", "importe_uyu", "importe_usd"]
    df = pd.DataFrame(data.get("transacciones", []), columns=cols)
    idx_series = pd.Series(range(1, len(df) + 1), dtype="Int64")
    df.insert(0, "idx", idx_series)
    return df


def append_total_row(df: pd.DataFrame, data: dict) -> pd.DataFrame:
    total_row = {
        "idx": pd.NA,
        "fecha": data.get("fecha_emision"),
        "detalle": "TOTAL",
        "importe_origen": None,
        "importe_uyu": data.get("total_uyu"),
        "importe_usd": data.get("total_usd"),
    }
    out = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)
    if "idx" in out.columns:
        out["idx"] = out["idx"].astype("Int64")
    return out


def _safe_sum(series: pd.Series) -> float:
    numeric = pd.to_numeric(series, errors="coerce")
    return float(numeric.sum(skipna=True))


def check_totals(df: pd.DataFrame, data: dict, tolerance: float = 0.01) -> None:
    calc_uyu = _safe_sum(df["importe_uyu"]) if "importe_uyu" in df else 0.0
    calc_usd = _safe_sum(df["importe_usd"]) if "importe_usd" in df else 0.0
    reported_uyu = data.get("total_uyu")
    reported_usd = data.get("total_usd")

    if reported_uyu is not None:
        diff_uyu = abs(calc_uyu - float(reported_uyu))
        status = "OK" if diff_uyu <= tolerance else "MISMATCH"
        print(f"[{status}] total_uyu calculado={calc_uyu:.2f} reportado={reported_uyu}")
    else:
        print(f"[INFO] total_uyu no reportado; calculado={calc_uyu:.2f}")

    if reported_usd is not None:
        diff_usd = abs(calc_usd - float(reported_usd))
        status = "OK" if diff_usd <= tolerance else "MISMATCH"
        print(f"[{status}] total_usd calculado={calc_usd:.2f} reportado={reported_usd}")
    else:
        print(f"[INFO] total_usd no reportado; calculado={calc_usd:.2f}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Convierte estado.json a CSV con una fila de totales.",
    )
    parser.add_argument(
        "input",
        nargs="?",
        default="estado.json",
        help="Ruta al JSON de estado (por defecto: estado.json)",
    )
    parser.add_argument(
        "-o",
        "--output",
        help="Ruta de salida CSV (por defecto, mismo nombre con .csv)",
    )
    args = parser.parse_args()

    in_path = Path(args.input)
    if not in_path.exists():
        raise FileNotFoundError(f"No se encontró el archivo: {in_path}")

    data = load_estado(in_path)
    tx_df = build_transactions_dataframe(data)
    check_totals(tx_df, data)
    df = append_total_row(tx_df, data)

    out_path = Path(args.output) if args.output else in_path.with_suffix(".csv")
    df.to_csv(out_path, index=False)
    print(f"CSV generado en: {out_path}")


if __name__ == "__main__":
    main()
