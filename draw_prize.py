"""
Weighted random prize draw. Decrements the chosen SKU's Qty by 1 and saves the list.

Uses a local tab-separated file (SKU\\tQty\\timg), default: Input List.xlsx next to this script.
The img column is optional: a picture path (relative to the list file or absolute) or http(s) URL for the UI.
Saving can fail if another program (e.g. Excel) holds an exclusive lock on the file.
"""

from __future__ import annotations

import argparse
import os
import random
import sys
import time
from dataclasses import dataclass
from pathlib import Path


def default_list_path() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent / "Input List.xlsx"
    return Path(__file__).resolve().parent / "Input List.xlsx"


class PrizeDrawError(RuntimeError):
    """Recoverable errors for library / UI callers (CLI maps these to stderr + exit)."""


def load_rows(path: Path) -> tuple[str, list[tuple[str, int, str]]]:
    text = path.read_text(encoding="utf-8")
    lines = [ln for ln in text.splitlines() if ln.strip() != ""]
    if not lines:
        raise PrizeDrawError("Input file is empty.")
    header = lines[0]
    rows: list[tuple[str, int, str]] = []
    for line in lines[1:]:
        parts = line.split("\t")
        if len(parts) < 2:
            continue
        sku = parts[0].strip()
        try:
            qty = int(parts[1].strip())
        except ValueError as e:
            raise PrizeDrawError(f"Bad quantity for line {line!r}: {e}") from e
        img = parts[2].strip().strip('"').strip("'") if len(parts) > 2 else ""
        rows.append((sku, qty, img))
    return header, rows


def save_rows(path: Path, header: str, rows: list[tuple[str, int, str]], newline: str) -> None:
    body = newline.join(f"{sku}\t{qty}\t{img}" for sku, qty, img in rows)
    content = header + newline + body + newline
    tmp = path.with_name(path.name + ".tmp")
    tmp.write_text(content, encoding="utf-8")
    last: PermissionError | OSError | None = None
    for _ in range(30):
        try:
            os.replace(tmp, path)
            return
        except PermissionError as e:
            last = e
            time.sleep(0.1)
        except OSError as e:
            if getattr(e, "winerror", None) == 5:
                last = e
                time.sleep(0.1)
            else:
                tmp.unlink(missing_ok=True)
                raise
    tmp.unlink(missing_ok=True)
    raise PrizeDrawError(
        "Could not save the list file (access denied). "
        "Close it in Excel or any other program that has it open, then run the draw again. "
        f"File: {path}"
    ) from last


def pick_sku(rows: list[tuple[str, int, str]]) -> int:
    pool = [(i, q) for i, (_, q, _) in enumerate(rows) if q > 0]
    if not pool:
        raise PrizeDrawError("No prizes left (all quantities are 0).")
    indices, weights = zip(*pool)
    chosen_i = random.choices(indices, weights=weights, k=1)[0]
    return chosen_i


# Cap for building a literal unit pool on the wheel (each remaining item = one tile chance).
_WHEEL_EXPANDED_POOL_CAP = 10_000


def weighted_sku_entries(rows: list[tuple[str, int, str]]) -> tuple[list[str], list[int]]:
    """Unique SKUs with positive qty and their weights (same weights as ``pick_sku``)."""
    skus: list[str] = []
    weights: list[int] = []
    for sku, q, _ in rows:
        s = (sku or "").strip()
        if s and q > 0:
            skus.append(s)
            weights.append(int(q))
    return skus, weights


def expanded_unit_pool(
    rows: list[tuple[str, int, str]],
    *,
    max_units: int = _WHEEL_EXPANDED_POOL_CAP,
) -> list[str]:
    """One list entry per remaining unit (SKU repeated by its Qty)."""
    pool: list[str] = []
    for sku, q, _ in rows:
        s = (sku or "").strip()
        if not s or q <= 0:
            continue
        for _ in range(int(q)):
            pool.append(s)
            if len(pool) >= max_units:
                return pool
    return pool


def sample_skus_weighted(rows: list[tuple[str, int, str]], count: int) -> list[str]:
    """Sample ``count`` SKUs; each draw is independent with odds proportional to Qty."""
    skus, weights = weighted_sku_entries(rows)
    if not skus or count <= 0:
        return []
    return random.choices(skus, weights=weights, k=count)


def sample_wheel_strip_labels(rows: list[tuple[str, int, str]], length: int) -> list[str]:
    """
    Labels for wheel strip cells (idle / loading / filler tiles).
    Uses a literal unit pool when total qty is modest so tile frequency matches stock.
    """
    if length <= 0:
        return []
    total = sum(q for _, q, _ in rows if q > 0)
    if total <= 0:
        return []
    if total <= _WHEEL_EXPANDED_POOL_CAP:
        expanded = expanded_unit_pool(rows)
        if expanded:
            return random.choices(expanded, k=length)
    return sample_skus_weighted(rows, length)


def build_wheel_spin_strip(
    rows: list[tuple[str, int, str]],
    length: int,
    *,
    winner: str,
    win_idx: int,
) -> list[str]:
    """Build a strip of ``length`` cells; ``win_idx`` shows ``winner``, others are qty-weighted."""
    fillers = sample_wheel_strip_labels(rows, max(0, length - 1))
    strip: list[str] = []
    fi = 0
    for i in range(length):
        if i == win_idx:
            strip.append(winner)
        elif fi < len(fillers):
            strip.append(fillers[fi])
            fi += 1
        else:
            strip.append(winner)
    return strip


@dataclass
class SpinResult:
    """One weighted pick against a snapshot of rows."""

    index: int
    sku: str
    qty: int
    img: str = ""


class FileDrawSession:
    """Local tab-separated list: refresh, pick, commit one decrement."""

    def __init__(self, path: Path) -> None:
        self.path = path.resolve()
        self.header = ""
        self.newline = "\n"
        self.rows: list[tuple[str, int, str]] = []
        self.refresh()

    def refresh(self) -> None:
        if not self.path.is_file():
            raise PrizeDrawError(f"File not found: {self.path}")
        raw = self.path.read_bytes()
        self.newline = "\r\n" if b"\r\n" in raw else "\n"
        self.header, self.rows = load_rows(self.path)

    def pick(self) -> SpinResult:
        self.refresh()
        i = pick_sku(self.rows)
        sku, qty, img = self.rows[i]
        return SpinResult(i, sku, qty, img)

    def commit(self, result: SpinResult) -> None:
        self.refresh()
        if result.index >= len(self.rows):
            raise PrizeDrawError("The list changed since your spin. Try again.")
        sku, qty, img = self.rows[result.index]
        if sku != result.sku or qty != result.qty:
            raise PrizeDrawError("The list changed since your spin. Try again.")
        self.rows[result.index] = (sku, max(0, qty - 1), img)
        save_rows(self.path, self.header, self.rows, self.newline)

    def revert_decrement(self, result: SpinResult) -> None:
        """Restore +1 qty for a row that was decremented by ``commit`` with this same ``result`` snapshot."""
        self.refresh()
        if result.index >= len(self.rows):
            raise PrizeDrawError("The list changed — cannot undo (row missing).")
        sku, qty, img = self.rows[result.index]
        if sku != result.sku:
            raise PrizeDrawError("The list changed — cannot undo (SKU mismatch).")
        if qty != result.qty - 1:
            raise PrizeDrawError(
                f"The list changed — cannot undo (expected qty {result.qty - 1} after draw, found {qty})."
            )
        self.rows[result.index] = (sku, result.qty, img)
        save_rows(self.path, self.header, self.rows, self.newline)


def open_draw_session(list_path: Path) -> FileDrawSession:
    return FileDrawSession(list_path.resolve())


def main() -> None:
    parser = argparse.ArgumentParser(description="Draw one weighted prize and decrement Qty.")
    parser.add_argument(
        "-f",
        "--file",
        type=Path,
        default=default_list_path(),
        help="Path to tab-separated SKU / Qty / img list (default: Input List.xlsx next to this script).",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Show the pick but do not persist changes.",
    )
    args = parser.parse_args()
    path: Path = args.file
    if not path.is_file():
        raise SystemExit(f"File not found: {path}")

    raw = path.read_bytes()
    newline = "\r\n" if b"\r\n" in raw else "\n"

    try:
        header, rows = load_rows(path)
        i = pick_sku(rows)
    except PrizeDrawError as e:
        raise SystemExit(str(e)) from e
    sku, qty, img = rows[i]
    extra = f"  img={img!r}" if img else ""
    print(f"Prize: {sku}  (was Qty {qty}){extra}")

    if args.dry_run:
        print("Dry run: file not updated.")
        return

    rows[i] = (sku, max(0, qty - 1), img)
    try:
        save_rows(path, header, rows, newline)
    except PrizeDrawError as e:
        raise SystemExit(str(e)) from e
    print(f"Updated: {sku} Qty -> {rows[i][1]}")


if __name__ == "__main__":
    main()
