#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from docflow.sdk import DocflowClient, profiles
from docflow.sdk.config import SdkConfig

DEFAULT_PROFILE = "lineas_gastos/v0"
DEFAULT_FILE = "data/sep/RENDREQ-0372_MPLOPEZ 10.25/RENDREQ-0372_MPLOPEZ 10.25_page_8.pdf"
DEFAULT_OUT = "sample_output_sdk.json"


def _build_prompt(base_prompt: str | None, statement_text: str) -> str:
    parts = []
    if base_prompt:
        parts.append(base_prompt.strip())
    parts.append("ESTADO DE CUENTA PARSEADO:\n" + statement_text.strip())
    parts.append(
        "Usa esta info para completar el campo 'Estado de cuenta' de cada item. "
        "Si no aparece pone, estado unmatched y justifica."
    )
    return "\n\n".join(p for p in parts if p)


def _to_jsonable(obj):
    if hasattr(obj, "to_dict"):
        return obj.to_dict()
    if isinstance(obj, list):
        return [_to_jsonable(x) for x in obj]
    if isinstance(obj, dict):
        return {k: _to_jsonable(v) for k, v in obj.items()}
    return obj


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Run DocFlow SDK to match CLI envelope for a profile."
    )
    parser.add_argument("--profile", default=DEFAULT_PROFILE)
    parser.add_argument("--file", default=DEFAULT_FILE)
    parser.add_argument("--out", default=DEFAULT_OUT)
    parser.add_argument("--profile-dir", default=None)
    parser.add_argument("--statement-path", default=None)
    parser.add_argument("--statement-text", default=None)
    parser.add_argument("--multi", default="per_file", choices=["per_file", "aggregate", "both"])
    args = parser.parse_args()

    repo_root = Path(__file__).resolve().parent
    profile_dir = Path(args.profile_dir) if args.profile_dir else repo_root

    file_path = Path(args.file)
    if not file_path.exists():
        print(f"ERROR: file not found: {file_path}", file=sys.stderr)
        return 2

    statement_text = None
    if args.statement_text:
        statement_text = args.statement_text
    elif args.statement_path:
        statement_text = Path(args.statement_path).read_text(encoding="utf-8")

    config = SdkConfig(profile_dir=profile_dir)
    client = DocflowClient(mode="local", config=config)

    try:
        if statement_text:
            profile = profiles.load_profile(args.profile, config)
            profile.prompt = _build_prompt(profile.prompt, statement_text)
            result = client._execute(
                schema=profile.schema,
                files=[str(file_path)],
                profile_name=profile.name,
                profile=profile,
                multi_mode=args.multi,
            )
        else:
            result = client.run_profile(args.profile, [str(file_path)], multi_mode=args.multi)
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1

    payload = _to_jsonable(result)
    out_path = Path(args.out)
    out_path.write_text(json.dumps(payload, ensure_ascii=True, indent=2), encoding="utf-8")
    print(f"Wrote {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
