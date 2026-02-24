#!/usr/bin/env python3
from __future__ import annotations

import argparse
import sys
from pathlib import Path


TEXT_EXTENSIONS = {".md", ".ps1", ".py", ".bas", ".cls", ".frm"}
MOJIBAKE_HINTS = (
    "Ã",
    "Â",
    "Ã‚Â",
    "â€”",
    "â€“",
    "â€œ",
    "â€",
    "â€˜",
    "â€™",
    "â€¦",
)


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Corrige mojibake comun en archivos de texto (latin1->utf8)."
    )
    parser.add_argument(
        "target",
        nargs="?",
        default="dist/public_release",
        help="Directorio a recorrer (default: dist/public_release).",
    )
    parser.add_argument(
        "--max-iterations",
        type=int,
        default=2,
        help="Cantidad maxima de pasadas de correccion por archivo (default: 2).",
    )
    return parser.parse_args(argv)


def is_text_candidate(path: Path) -> bool:
    return path.is_file() and path.suffix.lower() in TEXT_EXTENSIONS


def mojibake_score(text: str) -> int:
    return sum(text.count(token) for token in MOJIBAKE_HINTS)


def has_mojibake(text: str) -> bool:
    return mojibake_score(text) > 0


def decode_bytes(raw: bytes) -> tuple[str, str]:
    for enc in ("utf-8", "cp1252", "latin-1"):
        try:
            return raw.decode(enc), enc
        except UnicodeDecodeError:
            continue
    # latin-1 siempre deberia funcionar, pero queda fallback defensivo
    return raw.decode("latin-1", errors="replace"), "latin-1"


def try_fix_once(text: str) -> str:
    try:
        return text.encode("latin-1").decode("utf-8")
    except (UnicodeEncodeError, UnicodeDecodeError):
        return text


def fix_mojibake_text(text: str, max_iterations: int) -> tuple[str, int]:
    current = text
    applied = 0

    for _ in range(max_iterations):
        score_before = mojibake_score(current)
        if score_before == 0:
            break

        candidate = try_fix_once(current)
        if candidate == current:
            break

        score_after = mojibake_score(candidate)
        if score_after <= score_before:
            current = candidate
            applied += 1
            continue

        break

    return current, applied


def process_file(path: Path, max_iterations: int) -> tuple[bool, str, int]:
    raw = path.read_bytes()
    text, detected_encoding = decode_bytes(raw)

    if not has_mojibake(text):
        return False, detected_encoding, 0

    fixed_text, iterations = fix_mojibake_text(text, max_iterations=max_iterations)
    if iterations <= 0 or fixed_text == text:
        return False, detected_encoding, 0

    path.write_bytes(fixed_text.encode("utf-8"))
    return True, detected_encoding, iterations


def main(argv: list[str]) -> int:
    args = parse_args(argv)
    target_dir = Path(args.target).resolve()

    if not target_dir.exists() or not target_dir.is_dir():
        print(f"[ERROR] Directorio invalido: {target_dir}", file=sys.stderr)
        return 1

    scanned = 0
    changed = 0

    for path in sorted(target_dir.rglob("*")):
        if not is_text_candidate(path):
            continue

        scanned += 1
        was_changed, detected_encoding, iterations = process_file(
            path, max_iterations=max(1, args.max_iterations)
        )
        if was_changed:
            changed += 1
            rel = path.relative_to(target_dir)
            print(f"[FIXED] {rel} (iter={iterations}, src-encoding={detected_encoding})")

    print(f"[SUMMARY] scanned={scanned} changed={changed} target={target_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))

