"""
Interactive prize draw: compact chroma overlay, HTML wheel for OBS, prize board window, and Spin & controls window.
"""

from __future__ import annotations

import mimetypes
import os
import random
import sys
import re
import ssl
import threading
import time
from dataclasses import dataclass
from datetime import datetime
from io import BytesIO
from typing import Literal
import urllib.request
import tkinter as tk
import tkinter.font as tkfont
import webbrowser
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk
from urllib.parse import quote, urlparse, unquote
from urllib.request import url2pathname

import draw_prize

try:
    from openpyxl import Workbook, load_workbook
except ImportError:
    Workbook = None  # type: ignore[assignment,misc]
    load_workbook = None  # type: ignore[assignment,misc]


WHEEL_FRAME_BG = "#1a1a2e"
# Solid dark green behind wheel chrome for OBS Chroma Key. Avoid this exact color in prize art.
OBS_CHROMA_KEY = "#196619"
WHEEL_BG = "#16213e"
WHEEL_CELL_BG = "#0f3460"
WHEEL_CELL_BORDER = "#533483"
WHEEL_FG = "#eaeaea"
WHEEL_WIN_BG = "#e94560"
WHEEL_WIN_FG = "#ffffff"
WHEEL_POINTER = "#ffd93d"
WHEEL_ACCENT = "#ff6b6b"
WHEEL_TITLE = "#ffd93d"
WHEEL_MUTED = "#4a4e69"

# Horizontal prize strip (canvas): slate / indigo — used only for wheel drawing (chroma canvas bg stays OBS_CHROMA_KEY).
WHEEL_STRIP_SLOT_BG = "#1a2744"
WHEEL_STRIP_SLOT_BD = "#334155"
WHEEL_STRIP_FG = "#e2e8f0"
WHEEL_STRIP_WIN_BG = "#4f46e5"
WHEEL_STRIP_WIN_FG = "#f8fafc"
WHEEL_STRIP_WIN_RING = "#a5b4fc"
WHEEL_STRIP_POINTER = "#fbbf24"
WHEEL_STRIP_POINTER_EDGE = "#0f172a"
WHEEL_STRIP_TRACK = "#273549"
# Wheel strip vertical scale (cell height + pads + overlay window vs original design).
WHEEL_VERTICAL_SCALE = 1.2

WHEEL_STRIP_PAD_TOP = int(round(28 * WHEEL_VERTICAL_SCALE))
WHEEL_STRIP_PAD_BOTTOM = int(round(14 * WHEEL_VERTICAL_SCALE))
WHEEL_STRIP_CELL_GAP = 2.5
# HTML / OBS Browser Source: strip only — base scale vs. Tk, then width/pad tweaks for overlay layout.
HTML_WHEEL_DISPLAY_SCALE = 1.25
HTML_WHEEL_STRIP_WIDTH_MUL = 0.8  # strip cells 20% narrower than (Tk × DISPLAY_SCALE)
HTML_WHEEL_STRIP_PAD_V_MUL = 1.1  # vertical strip padding +10% vs. (Tk pad × DISPLAY_SCALE)
# HTML strip row height: divide snapshot pad top/bottom by this so cells are ~30% taller (same scroll math).
HTML_WHEEL_STRIP_CELL_HEIGHT_MUL = 1.3
BTN_SPIN = "#2980b9"
BTN_SUPER = "#8e44ad"
BTN_REROLL = "#16a085"
BTN_KEEP = "#c0392b"
BTN_SKIP = "#dc7633"
BTN_UNDO = "#1f618d"
BTN_DISABLED = "#3d3d54"
DRAG_BAR = "#f39c12"

# Prize board storage grid (row-major: `cols` tiles per row, then wrap).
INV_GRID_GAP = 2
INV_SLOT_INSET = 2
INV_SLOT_EDGE = "#7a9fd6"
INV_SLOT_FACE = "#0c1220"
INV_QTY_BADGE_BG = "#ffd93d"
INV_QTY_BADGE_FG = "#1a1a2e"

# Prize board grid: columns = items per row; rows = target row count for sizing tiles to window height.
PRIZE_BOARD_SLOT_MIN_PX = 44
PRIZE_BOARD_SLOT_MAX_PX = 160
PRIZE_BOARD_GRID_COLS_DEFAULT = 6
PRIZE_BOARD_ROWS_FIT_DEFAULT = 4

# Prize board Toplevel: readable on a monitor (not keyed for OBS chroma).
PRIZE_BOARD_WINDOW_BG = WHEEL_FRAME_BG
PRIZE_BOARD_CONTENT_BG = WHEEL_BG
PRIZE_BOARD_HEADER_BG = "#152238"
PRIZE_BOARD_RIM = WHEEL_CELL_BORDER

# Spin & controls Toplevel: same on-monitor palette; cards sit slightly inset.
SPIN_CONTROLS_CARD_BG = "#131c2e"
SPIN_CONTROLS_CARD_BORDER = WHEEL_CELL_BORDER
# Default / minimum size — window is user-resizable (drag edges or corners to expand).
SPIN_CONTROLS_WIN_WIDTH = 720
SPIN_CONTROLS_WIN_HEIGHT = 920
SPIN_CONTROLS_WIN_MIN_WIDTH = 560
SPIN_CONTROLS_WIN_MIN_HEIGHT = 520


def script_dir() -> Path:
    """Folder for user data (winner_sessions, default browse dir): next to the .exe when frozen."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def resource_root() -> Path:
    """Read-only bundled assets (e.g. web/); PyInstaller one-file extract dir when frozen."""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS).resolve()
    return Path(__file__).resolve().parent


def _pillow_runtime_available() -> bool:
    try:
        import PIL.Image  # noqa: F401
        import PIL.ImageTk  # noqa: F401

        return True
    except ImportError:
        return False


def _looks_like_html_head(data: bytes) -> bool:
    head = data[:800].lstrip().lower()
    return head.startswith((b"<", b"<!doctype")) or b"<html" in head[:1200]


def _normalize_image_ref(ref: str) -> str:
    ref = (ref or "").strip().strip("\ufeff").strip('"').strip("'")
    if not ref:
        return ""
    if ref.lower().startswith("file:"):
        parsed = urlparse(ref.replace("\\", "/"))
        raw_path = unquote(parsed.path or "")
        if parsed.netloc and os.name == "nt":
            ref = f"\\\\{parsed.netloc}{raw_path.replace('/', os.sep)}"
        else:
            try:
                ref = url2pathname(raw_path)
            except (ValueError, OSError):
                pass
        return ref.strip()
    m = re.search(r"https?://[^\s]+", ref, flags=re.IGNORECASE)
    if m:
        return m.group(0).rstrip(").,]>'\"")
    return ref


def _path_for_project_join(ref: str) -> str:
    ref = ref.strip()
    if not ref:
        return ref
    if ref.startswith("\\\\") or (len(ref) >= 2 and ref[1] == ":"):
        return ref
    return ref.lstrip("/\\")


class PrizeListOverlay(tk.Toplevel):
    """Separate window: PC-style remaining prizes grid, tuned for viewing on a monitor (not chroma/OBS)."""

    def __init__(self, app: "DrawPrizeApp") -> None:
        super().__init__(app)
        self._app = app
        self.title("Energy Break — Prize board")
        self.minsize(560, 480)
        self.geometry("760x680")
        self.configure(bg=PRIZE_BOARD_WINDOW_BG)
        try:
            self.attributes("-topmost", False)
        except tk.TclError:
            pass
        try:
            self.resizable(True, True)
        except tk.TclError:
            pass

        app.update_idletasks()
        px = app.winfo_rootx() + app.winfo_width() + 16
        py = app.winfo_rooty()
        self.geometry(f"760x680+{px}+{py}")

        self.protocol("WM_DELETE_WINDOW", self._close_overlay)

        outer = tk.Frame(
            self,
            bg=PRIZE_BOARD_WINDOW_BG,
            highlightbackground=PRIZE_BOARD_RIM,
            highlightthickness=1,
        )
        self._void_outer_frame = outer
        outer.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        inner = tk.Frame(outer, bg=PRIZE_BOARD_CONTENT_BG)
        inner.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

        head = tk.Frame(inner, bg=PRIZE_BOARD_HEADER_BG, padx=14, pady=12)
        head.pack(fill=tk.X)
        head_top = tk.Frame(head, bg=PRIZE_BOARD_HEADER_BG)
        head_top.pack(fill=tk.X)
        tk.Label(
            head_top,
            text="Still in draw",
            font=("Segoe UI", 16, "bold"),
            bg=PRIZE_BOARD_HEADER_BG,
            fg=WHEEL_TITLE,
            anchor=tk.W,
        ).pack(side=tk.LEFT)
        app._prize_board_totals_label = tk.Label(
            head_top,
            text="—",
            font=("Segoe UI", 12, "bold"),
            bg=PRIZE_BOARD_HEADER_BG,
            fg=WHEEL_POINTER,
            anchor=tk.E,
        )
        app._prize_board_totals_label.pack(side=tk.RIGHT)
        tk.Label(
            head,
            text="Prizes from your list with quantity remaining (each tile shows stock on hand).",
            font=("Segoe UI", 9),
            bg=PRIZE_BOARD_HEADER_BG,
            fg=WHEEL_MUTED,
            anchor=tk.W,
            justify=tk.LEFT,
        ).pack(fill=tk.X, pady=(8, 0))

        scroll_wrap = tk.Frame(inner, bg=PRIZE_BOARD_CONTENT_BG)
        scroll_wrap.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 14))
        app.remaining_scroll = None  # type: ignore[assignment]
        app.remaining_canvas = None  # type: ignore[assignment]
        app._remaining_canvas_win = None  # type: ignore[assignment]
        app.remaining_inner = tk.Frame(scroll_wrap, bg=PRIZE_BOARD_CONTENT_BG)
        app.remaining_inner.pack(fill=tk.BOTH, expand=True)
        app.remaining_inner.bind("<Configure>", lambda _e: app._schedule_prize_board_layout_refresh())

    def _close_overlay(self) -> None:
        app = self._app
        app._cancel_prize_board_resize_refresh()
        app._prize_win = None
        app._prize_board_totals_label = None  # type: ignore[assignment]
        app.remaining_scroll = None  # type: ignore[assignment]
        app.remaining_canvas = None  # type: ignore[assignment]
        app.remaining_inner = None  # type: ignore[assignment]
        app._remaining_canvas_win = None  # type: ignore[assignment]
        try:
            self.destroy()
        except tk.TclError:
            pass


@dataclass
class _UndoSpinSnapshot:
    """Single-step undo: spin, super keep, skip, or fill-skip (update empty winner row)."""

    kind: Literal["spin", "skip_log", "skip_counter_only", "backfill"]
    list_path: Path
    result: draw_prize.SpinResult | None
    spot_written: int
    winner_row_added: bool


class SpinControlsWindow(tk.Toplevel):
    """Draw / Super / Reroll / Keep, edit spin, fill skipped, HTML wheel, session, prize board — resizable, scrollable."""

    def __init__(self, app: "DrawPrizeApp") -> None:
        super().__init__(app)
        self._app = app
        self.title("Energy Break — Spin & controls")
        win_bg = PRIZE_BOARD_WINDOW_BG
        self.configure(bg=win_bg)
        try:
            self.attributes("-topmost", False)
        except tk.TclError:
            pass
        try:
            self.resizable(True, True)
        except tk.TclError:
            pass
        self.minsize(SPIN_CONTROLS_WIN_MIN_WIDTH, SPIN_CONTROLS_WIN_MIN_HEIGHT)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        app.update_idletasks()
        w, h = SPIN_CONTROLS_WIN_WIDTH, SPIN_CONTROLS_WIN_HEIGHT
        px = max(0, app.winfo_rootx() + 24)
        py = max(0, app.winfo_rooty() + app.winfo_height() + 16)
        self.geometry(f"{w}x{h}+{px}+{py}")

        outer = tk.Frame(
            self,
            bg=win_bg,
            highlightbackground=SPIN_CONTROLS_CARD_BORDER,
            highlightthickness=1,
        )
        outer.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)
        shell = tk.Frame(outer, bg=PRIZE_BOARD_CONTENT_BG)
        shell.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

        header = tk.Frame(shell, bg=PRIZE_BOARD_HEADER_BG)
        header.pack(fill=tk.X)
        tk.Label(
            header,
            text="Spin & controls",
            font=("Segoe UI", 17, "bold"),
            bg=PRIZE_BOARD_HEADER_BG,
            fg=WHEEL_TITLE,
            anchor=tk.W,
        ).pack(fill=tk.X, padx=18, pady=(14, 4))
        tk.Label(
            header,
            text="Draw, edit spin, fill skipped spots, HTML wheel for OBS, new wheel, prize board (scroll below).",
            font=("Segoe UI", 9),
            bg=PRIZE_BOARD_HEADER_BG,
            fg=WHEEL_MUTED,
            anchor=tk.W,
            justify=tk.LEFT,
        ).pack(fill=tk.X, padx=18, pady=(0, 14))

        scroll_outer = tk.Frame(shell, bg=PRIZE_BOARD_CONTENT_BG)
        scroll_outer.pack(fill=tk.BOTH, expand=True)
        scroll_canvas = tk.Canvas(
            scroll_outer,
            bg=PRIZE_BOARD_CONTENT_BG,
            highlightthickness=0,
            bd=0,
        )
        vscroll = tk.Scrollbar(
            scroll_outer,
            orient=tk.VERTICAL,
            command=scroll_canvas.yview,
            bg=PRIZE_BOARD_CONTENT_BG,
            troughcolor=SPIN_CONTROLS_CARD_BG,
            activebackground=WHEEL_CELL_BG,
        )
        scroll_canvas.configure(yscrollcommand=vscroll.set)
        try:
            scroll_canvas.configure(takefocus=True)
        except tk.TclError:
            pass
        vscroll.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 2))
        scroll_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._spin_controls_scroll_canvas = scroll_canvas

        body = tk.Frame(scroll_canvas, bg=PRIZE_BOARD_CONTENT_BG)
        body_win = scroll_canvas.create_window((0, 0), window=body, anchor=tk.NW)

        def _sync_scroll_body_width(_event: object | None = None) -> None:
            try:
                cw = max(120, int(scroll_canvas.winfo_width()) - 2)
            except tk.TclError:
                return
            try:
                scroll_canvas.itemconfigure(body_win, width=cw)
            except tk.TclError:
                return

        def _on_scroll_body_configure(_event: tk.Event) -> None:
            scroll_canvas.configure(scrollregion=scroll_canvas.bbox("all"))
            _sync_scroll_body_width()

        def _on_scroll_canvas_configure(_event: tk.Event) -> None:
            _sync_scroll_body_width()

        body.bind("<Configure>", _on_scroll_body_configure, add="+")
        scroll_canvas.bind("<Configure>", _on_scroll_canvas_configure, add="+")

        def _wheel_spin_controls(event: tk.Event) -> None:
            if getattr(event, "delta", 0):
                scroll_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        def _wheel_spin_controls_up(_event: tk.Event) -> None:
            scroll_canvas.yview_scroll(-1, "units")

        def _wheel_spin_controls_down(_event: tk.Event) -> None:
            scroll_canvas.yview_scroll(1, "units")

        def _bind_spin_controls_mousewheel(_event: tk.Event) -> None:
            scroll_canvas.bind_all("<MouseWheel>", _wheel_spin_controls)
            scroll_canvas.bind_all("<Button-4>", _wheel_spin_controls_up)
            scroll_canvas.bind_all("<Button-5>", _wheel_spin_controls_down)

        def _unbind_spin_controls_mousewheel(_event: tk.Event) -> None:
            try:
                scroll_canvas.unbind_all("<MouseWheel>")
                scroll_canvas.unbind_all("<Button-4>")
                scroll_canvas.unbind_all("<Button-5>")
            except tk.TclError:
                pass

        scroll_outer.bind("<Enter>", _bind_spin_controls_mousewheel)
        scroll_outer.bind("<Leave>", _unbind_spin_controls_mousewheel)

        inner = tk.Frame(body, bg=PRIZE_BOARD_CONTENT_BG)
        inner.pack(fill=tk.BOTH, expand=True, padx=18, pady=(0, 18))
        app._build_spin_controls_inner(inner)

        def _initial_scroll_sync() -> None:
            _sync_scroll_body_width()
            try:
                scroll_canvas.configure(scrollregion=scroll_canvas.bbox("all"))
            except tk.TclError:
                pass

        app.after_idle(_initial_scroll_sync)

        self.bind("<Button-1>", app._on_ui_click_defocus_wheel_name_entry, add="+")

        app._spin_controls_wheel_link_reset()
        self.bind("<Configure>", self._on_spin_controls_configure)

    def _on_spin_controls_configure(self, event: tk.Event) -> None:
        if event.widget is not self:
            return
        self._app._schedule_spin_controls_wheel_sync()

    def _on_close(self) -> None:
        self._app._cancel_spin_controls_wheel_sync()
        self._app._wheel_control_link_ref = None
        self._app._controls_win = None
        try:
            c = getattr(self, "_spin_controls_scroll_canvas", None)
            if c is not None:
                c.unbind_all("<MouseWheel>")
                c.unbind_all("<Button-4>")
                c.unbind_all("<Button-5>")
        except tk.TclError:
            pass
        try:
            self.destroy()
        except tk.TclError:
            pass


class DrawPrizeApp(tk.Tk):
    WHEEL_CELL = int(round(118 * 1.5 * 0.8 * 0.9 * WHEEL_VERTICAL_SCALE))
    WHEEL_STRIP_LEN = 44
    # Home / idle horizontal strip: logical preview length; physical strip is doubled for seamless carousel.
    WHEEL_IDLE_STRIP_LEN = 11
    WHEEL_IDLE_CAROUSEL_COPIES = 2
    # Short strip while the worker loads picks (before the full spin strip is built).
    WHEEL_LOADING_STRIP_LEN = 7
    # Total horizontal clear pixels between adjacent wheel item rectangles (scaled with cell).
    WHEEL_CELL_INTER_GAP = int(round(18 * 1.5 * 0.8 * 0.9 * 1.1))
    # Strip spin: duration + ease control how the wheel slows into the winner (quintic felt sticky at the end).
    WHEEL_SPIN_DURATION_SEC = 2.15 * 1.2
    WHEEL_SPIN_EASE_OUT_POWER = 3
    WHEEL_SPIN_TICK_MS = 16
    # Idle home strip: slow scroll left (loops when it reaches end of usable range).
    WHEEL_IDLE_DRIFT_MS = 45
    WHEEL_IDLE_DRIFT_PX_PER_SEC = 33.0
    # Resizing "Spin & controls" window stretches wheel strip (linked delta from sizes at open).
    WHEEL_CANVAS_H_LINK_MIN = int(round(96 * WHEEL_VERTICAL_SCALE))
    WHEEL_CANVAS_H_LINK_MAX = int(round(540 * WHEEL_VERTICAL_SCALE))
    # Main overlay: winner session listbox shows at most this many rows; scroll for the rest.
    OVERLAY_WINNER_LIST_VISIBLE_ROWS = 25

    @classmethod
    def _wheel_cell_xmargins(cls) -> tuple[int, int]:
        g = cls.WHEEL_CELL_INTER_GAP
        return g // 2, g - g // 2

    @classmethod
    def _wheel_slot_center_offset(cls) -> float:
        """X offset from i*cw - scroll to the visual center of slot i (matches Tk (x0+x1)/2)."""
        cw = float(cls.WHEEL_CELL)
        pl, pr = cls._wheel_cell_xmargins()
        return (float(pl) + cw - float(pr)) / 2.0

    @classmethod
    def _wheel_scroll_to_center_index(cls, index: float, viewport_w: float) -> float:
        """Scroll value so slot ``index`` is centered under x = viewport_w/2."""
        cw = float(cls.WHEEL_CELL)
        return index * cw + cls._wheel_slot_center_offset() - viewport_w / 2.0

    def __init__(self) -> None:
        super().__init__()
        self._void_bg_widgets: list[tk.Misc] = []
        self.title("Energy Break — Overlay")
        self.geometry(f"480x{int(round(280 * WHEEL_VERTICAL_SCALE))}")
        self.minsize(380, 200)
        self._prize_win: PrizeListOverlay | None = None
        self._controls_win: SpinControlsWindow | None = None
        self._spin_controls_wheel_sync_after: str | int | None = None
        self._wheel_control_link_ref: tuple[int, int] | None = None
        self._prize_board_totals_label: tk.Label | None = None
        self.configure(bg=self._void_bg())
        try:
            self.attributes("-topmost", False)
        except tk.TclError:
            pass
        self.overrideredirect(True)

        self._drag_offset: tuple[int, int] = (0, 0)
        self._session: draw_prize.FileDrawSession | None = None
        self._session_target = ""
        self._anim_after: str | int | None = None
        self._pulse_after: str | int | None = None
        self._pending_super: tuple[draw_prize.FileDrawSession, draw_prize.SpinResult] | None = None
        self._super_reroll_used = False
        self._busy = False
        self._obs_last_result_spot: int | None = None
        self._obs_last_result_sku: str = ""
        self._backfill_target_spot: int | None = None
        self._winner_log_path: Path | None = None
        self._winner_next_spin = 1
        self._undo_spin_snap: _UndoSpinSnapshot | None = None
        self._winner_overlay_poll_after: str | int | None = None
        self._winner_overlay_poll_started = False
        self._winner_overlay_last_sig: tuple[str, float] | None = None

        self._title_font = tkfont.Font(self, family="Segoe UI", size=13, weight="bold")
        self._wheel_font = tkfont.Font(self, family="Segoe UI", size=10, weight="normal")
        self._wheel_font_bold = tkfont.Font(self, family="Segoe UI", size=10, weight="bold")
        self._btn_font = tkfont.Font(self, family="Segoe UI", size=12, weight="bold")

        self._wheel_strip: list[str] = []
        self._wheel_win_idx = 0
        self._wheel_target_scroll = 0.0
        self._wheel_scroll = 0.0
        self._wheel_idle_offset = 0.0
        self._wheel_idle_drift_after: str | int | None = None
        self._wheel_pulse_highlight = False
        self._customer_empty_wheel = False
        self._wheel_sku_to_img: dict[str, str] = {}
        self._wheel_image_cache: dict[str, object] = {}
        self._wheel_http_server: object | None = None
        self._wheel_http_port: int = 0
        self._wheel_html_snapshot_cache: dict = {}
        self._wheel_banner_title: str | None = None
        self._wheel_banner_subtitle: str | None = None
        self._wheel_banner_is_error: bool = False
        self._inventory_image_cache: dict[tuple[str, int], object] = {}
        self._prize_board_grid_cols = tk.IntVar(value=PRIZE_BOARD_GRID_COLS_DEFAULT)
        self._prize_board_tile_rows_fit = tk.IntVar(value=PRIZE_BOARD_ROWS_FIT_DEFAULT)
        self._new_wheel_name_var = tk.StringVar(value="")
        self._prize_board_resize_after: str | int | None = None
        self._prize_board_last_layout: tuple[int, int, int] | None = None

        self._prize_board_grid_cols.trace_add("write", self._on_prize_board_grid_var_write)
        self._prize_board_tile_rows_fit.trace_add("write", self._on_prize_board_grid_var_write)

        self._build_drag_bar()

        self._setup_expanded = tk.BooleanVar(value=False)
        self._log_visible = tk.BooleanVar(value=False)

        self._setup_toggle_frame = tk.Frame(self, bg=self._void_bg())
        self._setup_toggle_frame.pack(fill=tk.X, padx=8, pady=(0, 4))
        self._register_void_bg(self._setup_toggle_frame)
        self._chk_setup_expand = tk.Checkbutton(
            self._setup_toggle_frame,
            text="Show setup (list path, dry run, log)",
            variable=self._setup_expanded,
            command=self._toggle_setup,
            bg=self._void_bg(),
            fg=WHEEL_FG,
            selectcolor=WHEEL_CELL_BG,
            activebackground=self._void_bg(),
            activeforeground=WHEEL_FG,
            font=("Segoe UI", 10),
        )
        self._chk_setup_expand.pack(side=tk.LEFT)
        self._register_void_bg(self._chk_setup_expand)

        self.setup_frame = tk.Frame(self, bg=self._void_bg())
        self._register_void_bg(self.setup_frame)
        self._build_setup_form(self.setup_frame)
        self._init_winner_session_log()

        self._wheel_host = tk.Frame(self, bg=self._void_bg())
        self._register_void_bg(self._wheel_host)
        self._wheel_host.pack(fill=tk.X, expand=False, padx=8, pady=(4, 8))
        self._build_wheel_area(self._wheel_host)
        self._build_prize_list_window()

        self.bind("<Button-1>", self._on_ui_click_defocus_wheel_name_entry, add="+")

        self._update_spin_counter_label()
        self._log(
            f"Chroma (wheel only): empty chrome uses solid {OBS_CHROMA_KEY} — key that out in OBS if you capture the wheel. "
            "Do not use the same green in prize images.\n"
        )
        self._log(
            "Spin & controls opens automatically in a separate window. "
            "Prize board: use the button on the wheel bar. "
            "HTML wheel for OBS: Spin & controls → Open HTML wheel. "
            "Drag the orange bar to move, ✕ to exit.\n"
        )
        self._reset_wheel_idle()
        self._refresh_prizes_label()
        self._sync_obs_overlay()
        self._wheel_html_server_start()
        self._wheel_publish_html_snapshot()
        self.after_idle(self._show_spin_controls)

    def _on_prize_board_grid_var_write(self, *_args: object) -> None:
        self._schedule_prize_board_layout_refresh()

    def _void_bg(self) -> str:
        """Background for overlay chrome keyed out in OBS (solid green, not OS transparency)."""
        return OBS_CHROMA_KEY

    def _prize_board_surface_bg(self) -> str:
        """Panel background for the on-monitor prize board (not chroma green)."""
        return PRIZE_BOARD_CONTENT_BG

    def _prize_board_cols_rows_clamped(self) -> tuple[int, int]:
        """User grid: columns (tiles per row) and rows (used only to fit tile height to the panel)."""
        try:
            c = int(self._prize_board_grid_cols.get())
        except (tk.TclError, ValueError, TypeError):
            c = PRIZE_BOARD_GRID_COLS_DEFAULT
        c = max(2, min(24, c))
        try:
            r = int(self._prize_board_tile_rows_fit.get())
        except (tk.TclError, ValueError, TypeError):
            r = PRIZE_BOARD_ROWS_FIT_DEFAULT
        r = max(2, min(20, r))
        return c, r

    def _prize_board_slot_thumb_for_inner(self, ri: tk.Frame) -> tuple[int, int, int]:
        """Square slot size and thumbnail side from inner frame size; returns (slot_px, thumb_px, cols)."""
        cols, rows_fit = self._prize_board_cols_rows_clamped()
        try:
            ri.update_idletasks()
        except tk.TclError:
            pass
        g = INV_GRID_GAP
        edge = 16
        w = max(260, int(ri.winfo_width()))
        h = max(140, int(ri.winfo_height()))
        slot_w = (w - edge * 2 - cols * (2 * g)) // max(1, cols)
        slot_h = (h - edge * 2 - rows_fit * (2 * g)) // max(1, rows_fit)
        slot = min(slot_w, slot_h)
        slot = max(PRIZE_BOARD_SLOT_MIN_PX, min(PRIZE_BOARD_SLOT_MAX_PX, slot))
        thumb = max(22, min(int(round(slot * 0.62)), slot - 16))
        return slot, thumb, cols

    def _schedule_prize_board_layout_refresh(self, _ev: object | None = None) -> None:
        """Debounce grid rebuild when the prize board is resized or grid prefs change."""
        aid = getattr(self, "_prize_board_resize_after", None)
        if aid is not None:
            try:
                self.after_cancel(aid)
            except tk.TclError:
                pass
        self._prize_board_resize_after = self.after(100, self._prize_board_layout_refresh_cb)

    def _prize_board_layout_refresh_cb(self) -> None:
        self._prize_board_resize_after = None
        self._refresh_remaining_skus_panel()

    def _cancel_prize_board_resize_refresh(self) -> None:
        aid = getattr(self, "_prize_board_resize_after", None)
        if aid is not None:
            try:
                self.after_cancel(aid)
            except tk.TclError:
                pass
            self._prize_board_resize_after = None

    def _register_void_bg(self, w: tk.Misc) -> None:
        if w not in self._void_bg_widgets:
            self._void_bg_widgets.append(w)

    def _sync_obs_overlay(self) -> None:
        void = self._void_bg()
        try:
            self.configure(bg=void)
        except tk.TclError:
            pass
        for w in self._void_bg_widgets:
            try:
                if isinstance(w, tk.Checkbutton):
                    w.configure(bg=void, activebackground=void)
                else:
                    w.configure(bg=void)
            except tk.TclError:
                pass
        if os.name == "nt":
            try:
                self.attributes("-transparentcolor", "")
            except tk.TclError:
                pass
            pw = getattr(self, "_prize_win", None)
            if pw is not None:
                try:
                    if pw.winfo_exists():
                        pw.attributes("-transparentcolor", "")
                except tk.TclError:
                    pass

    def _open_html_wheel_in_browser(self) -> None:
        """Open the HTML wheel page (same feed as OBS Browser Source)."""
        port = int(getattr(self, "_wheel_http_port", 0) or 0)
        if port <= 0:
            messagebox.showinfo(
                "HTML wheel",
                "The local wheel server did not start. Ensure web/wheel_spin.html exists next to the app.",
            )
            return
        url = f"http://127.0.0.1:{port}/"
        try:
            webbrowser.open(url)
        except OSError as e:
            messagebox.showerror("HTML wheel", f"Could not open browser:\n{e}\n\nURL:\n{url}")

    def _wheel_shutdown_html_server(self) -> None:
        httpd = getattr(self, "_wheel_http_server", None)
        if httpd is None:
            return
        self._wheel_http_server = None
        self._wheel_http_port = 0
        try:
            httpd.shutdown()
        except Exception:
            pass

    def _wheel_html_server_start(self) -> None:
        self._wheel_shutdown_html_server()
        html_path = resource_root() / "web" / "wheel_spin.html"
        try:
            from wheel_html_server import start_wheel_html_server as _wheel_http_boot
        except ImportError:
            self._log("HTML wheel: could not import wheel_html_server.py (browser mirror disabled).\n")
            return

        def _snap() -> dict:
            return dict(getattr(self, "_wheel_html_snapshot_cache", {}))

        def _img(sku: str) -> tuple[bytes, str] | None:
            return self._wheel_html_image_bytes(sku)

        httpd, port = _wheel_http_boot(html_path, _snap, _img)
        if httpd is None or not port:
            self._log("HTML wheel: web/wheel_spin.html not found — browser / OBS mirror disabled.\n")
            return
        self._wheel_http_server = httpd
        self._wheel_http_port = int(port)
        self._log(
            f"HTML wheel (OBS): http://127.0.0.1:{port}/  — transparent page; strip width "
            f"{int(round(100 * HTML_WHEEL_DISPLAY_SCALE * HTML_WHEEL_STRIP_WIDTH_MUL))}% of app cells, "
            f"strip vertical pad +{int(round(100 * (HTML_WHEEL_STRIP_PAD_V_MUL - 1)))}% (after display scale).\n"
        )

    def _wheel_html_image_bytes(self, sku: str) -> tuple[bytes, str] | None:
        if sku not in self._wheel_strip and sku not in self._wheel_sku_to_img:
            return None
        ref = self._wheel_sku_to_img.get(sku, "")
        refn = _normalize_image_ref(ref)
        if not refn or refn.lower().startswith(("http://", "https://")):
            return None
        path = self._resolve_prize_image_path(refn)
        if path is None or not path.is_file():
            return None
        try:
            blob = path.read_bytes()
        except OSError:
            return None
        ct, _enc = mimetypes.guess_type(str(path))
        return blob, ct or "application/octet-stream"

    def _clear_obs_wheel_result(self) -> None:
        self._obs_last_result_spot = None
        self._obs_last_result_sku = ""

    def _set_obs_wheel_result(self, spot: int, sku: str) -> None:
        self._obs_last_result_spot = int(spot)
        self._obs_last_result_sku = (sku or "").strip()

    def _wheel_obs_overlay_lines(
        self,
        spot_num: int | None,
        landed: bool,
        winning_sku: str,
        strip_len: int,
    ) -> tuple[str, str]:
        """Primary + secondary lines for the HTML / OBS spot chip (not the Tk wheel status)."""
        spot_x = f"Spot {spot_num}" if spot_num is not None else "Spot —"
        WSL = int(type(self).WHEEL_STRIP_LEN)
        WIL = int(type(self).WHEEL_IDLE_STRIP_LEN)
        anim = self._anim_after is not None
        pending_super = self._pending_super is not None
        busy = self._busy
        last_spot = self._obs_last_result_spot
        last_sku = (self._obs_last_result_sku or "").strip()
        win_sku = (winning_sku or "").strip()

        if strip_len == 0:
            return spot_x, ""

        bf = getattr(self, "_backfill_target_spot", None)
        if anim or (
            busy
            and not pending_super
            and strip_len not in (WSL, WIL, WIL * int(type(self).WHEEL_IDLE_CAROUSEL_COPIES))
        ):
            if bf is not None:
                return f"Spinning to fill Spot {bf} (was skipped)", ""
            return f"Spinning for {spot_x}", ""

        if pending_super:
            sub = win_sku
            if not sub and self._pending_super is not None:
                sub = (self._pending_super[1].sku or "").strip()
            line2 = f"Landed: {sub}" if sub else "Choose REROLL or KEEP"
            if self._super_reroll_used:
                return f"Super — {spot_x} · final pick", line2
            return f"Super — {spot_x}", line2

        if strip_len == WSL and landed:
            if last_spot is not None:
                return f"Winner Spot {last_spot}", last_sku or win_sku
            if busy and spot_num is not None:
                if bf is not None:
                    return f"Winner Spot {bf}", win_sku or last_sku
                return f"Winner {spot_x}", win_sku or last_sku
            if spot_num is not None and win_sku:
                return f"Winner {spot_x}", win_sku
            if spot_num is not None:
                return f"Winner {spot_x}", ""
            return spot_x, win_sku

        if strip_len == WIL or strip_len == WIL * int(type(self).WHEEL_IDLE_CAROUSEL_COPIES):
            return f"Next draw: {spot_x}", ""

        if strip_len == int(type(self).WHEEL_LOADING_STRIP_LEN):
            return f"Spinning for {spot_x}", ""

        return spot_x, ""

    def _wheel_prizes_left_label_text(self) -> str:
        """Short line for wheel chrome + HTML mirror (same quantity source as the main Total prizes label)."""
        n = self._total_qty_remaining()
        if n is None:
            return "Total prizes left: —"
        return f"Total prizes left: {n}"

    def _wheel_build_html_snapshot(self) -> dict:
        colors = {
            "slotBg": WHEEL_STRIP_SLOT_BG,
            "slotBd": WHEEL_STRIP_SLOT_BD,
            "fg": WHEEL_STRIP_FG,
            "winBg": WHEEL_STRIP_WIN_BG,
            "winFg": WHEEL_STRIP_WIN_FG,
            "winRing": WHEEL_STRIP_WIN_RING,
            "pointer": WHEEL_STRIP_POINTER,
            "pointerEdge": WHEEL_STRIP_POINTER_EDGE,
            "track": WHEEL_STRIP_TRACK,
            "title": WHEEL_TITLE,
            "accent": WHEEL_ACCENT,
        }
        spot = "Spot 1"
        st = ""
        if hasattr(self, "spin_counter_label"):
            try:
                spot = str(self.spin_counter_label.cget("text") or spot)
            except tk.TclError:
                pass
        if hasattr(self, "wheel_status"):
            try:
                st = str(self.wheel_status.cget("text") or "")
            except tk.TclError:
                pass
        spot_num = self._parse_spot_number_from_label(spot)
        if getattr(self, "_anim_after", None) is not None:
            spot_assignment = ""
        else:
            spot_assignment = (
                self._winner_assignment_display_for_spot(spot_num)
                if spot_num is not None
                else ""
            )
        base: dict = {
            "chroma": OBS_CHROMA_KEY,
            "colors": colors,
            "spotLabel": spot,
            "spotNumber": spot_num,
            "spotAssignment": spot_assignment,
            "statusText": st,
            "prizesLeftLabel": self._wheel_prizes_left_label_text(),
        }
        if self._wheel_banner_title and not self._wheel_strip:
            oh, os_ = self._wheel_obs_overlay_lines(spot_num, False, "", 0)
            base["obsHeadline"] = oh or spot
            base["obsSubhead"] = os_ or (spot_assignment or "")
            base.update(
                {
                    "mode": "banner",
                    "bannerTitle": self._wheel_banner_title,
                    "bannerSubtitle": self._wheel_banner_subtitle or "",
                    "bannerAccent": "error" if self._wheel_banner_is_error else "muted",
                }
            )
            return base

        c = self.wheel_canvas
        try:
            h = int(c.cget("height"))
            w = max(int(c.winfo_width()), 2)
        except tk.TclError:
            h, w = 120, 520
        cx = w / 2.0
        scroll_model = float(self._wheel_scroll)
        scroll_vis = self._wheel_scroll_for_strip_render()
        cw = float(self.WHEEL_CELL)
        pl, pr = self._wheel_cell_xmargins()
        gap = float(WHEEL_STRIP_CELL_GAP)
        hs = float(HTML_WHEEL_DISPLAY_SCALE)
        hw = hs * float(HTML_WHEEL_STRIP_WIDTH_MUL)
        pv = hs * float(HTML_WHEEL_STRIP_PAD_V_MUL)
        hch = float(HTML_WHEEL_STRIP_CELL_HEIGHT_MUL)
        landed = abs(scroll_model - float(self._wheel_target_scroll)) < 1.5
        slot_mid = self._wheel_slot_center_offset()
        cells: list[dict] = []
        winning_sku = ""
        idle_carousel = len(self._wheel_strip) == int(type(self).WHEEL_IDLE_STRIP_LEN) * int(
            type(self).WHEEL_IDLE_CAROUSEL_COPIES
        )
        for i, raw in enumerate(self._wheel_strip):
            cell_center_x = i * cw + slot_mid - scroll_vis
            if idle_carousel:
                win_cell = False
            elif len(self._wheel_strip) == self.WHEEL_IDLE_STRIP_LEN:
                win_cell = landed and abs(cell_center_x - cx) < cw * 0.38
            else:
                win_cell = landed and i == self._wheel_win_idx and abs(cell_center_x - cx) < cw * 0.38
            if win_cell:
                winning_sku = self._wheel_label_text(raw)
            ref = self._wheel_sku_to_img.get(raw, "")
            refn = _normalize_image_ref(ref)
            if refn and refn.lower().startswith(("http://", "https://")):
                img_url = refn
            elif refn:
                img_url = f"/wheel/img?sku={quote(raw, safe='')}"
            else:
                img_url = ""
            cells.append(
                {
                    "sku": self._wheel_label_text(raw),
                    "img": img_url,
                    "win": bool(win_cell),
                }
            )
        oh, os_ = self._wheel_obs_overlay_lines(spot_num, landed, winning_sku, len(self._wheel_strip))
        base["obsHeadline"] = oh or spot
        base["obsSubhead"] = os_
        base.update(
            {
                "mode": "strip",
                "viewportW": w,
                "canvasH": h,
                "cellW": int(round(self.WHEEL_CELL * hw)),
                "scroll": scroll_vis * hw,
                "padTop": max(18, int(round(WHEEL_STRIP_PAD_TOP * pv / hch))),
                "padBottom": max(8, int(round(WHEEL_STRIP_PAD_BOTTOM * pv / hch))),
                "gap": gap * hw,
                "pl": int(round(pl * hw)),
                "pr": int(round(pr * hw)),
                "pulse": bool(self._wheel_pulse_highlight),
                "cells": cells,
            }
        )
        return base

    def _wheel_html_snapshot_spin_frame(self, prev: dict) -> dict | None:
        """During strip spin animation, reuse cell image/SKU payload and only update scroll + win flags."""
        if not self._wheel_strip:
            return None
        prev_cells = prev.get("cells")
        if not isinstance(prev_cells, list) or len(prev_cells) != len(self._wheel_strip):
            return None
        for i, raw in enumerate(self._wheel_strip):
            pc = prev_cells[i]
            if not isinstance(pc, dict):
                return None
            if (pc.get("sku") or "") != self._wheel_label_text(raw):
                return None
        c = self.wheel_canvas
        try:
            h = int(c.cget("height"))
            w = max(int(c.winfo_width()), 2)
        except tk.TclError:
            return None
        if w != int(prev.get("viewportW") or -1) or h != int(prev.get("canvasH") or -1):
            return None
        cx = w / 2.0
        scroll = float(self._wheel_scroll)
        cw = float(self.WHEEL_CELL)
        pl, pr = self._wheel_cell_xmargins()
        slot_mid = self._wheel_slot_center_offset()
        landed = abs(scroll - float(self._wheel_target_scroll)) < 1.5
        hs = float(HTML_WHEEL_DISPLAY_SCALE)
        hw = hs * float(HTML_WHEEL_STRIP_WIDTH_MUL)
        winning_sku = ""
        new_cells: list[dict] = []
        for i, raw in enumerate(self._wheel_strip):
            pc = prev_cells[i]
            cell_center_x = i * cw + slot_mid - scroll
            win_cell = landed and i == self._wheel_win_idx and abs(cell_center_x - cx) < cw * 0.38
            if win_cell:
                winning_sku = self._wheel_label_text(raw)
            nd = dict(pc)
            nd["win"] = bool(win_cell)
            new_cells.append(nd)
        spot = "Spot 1"
        if hasattr(self, "spin_counter_label"):
            try:
                spot = str(self.spin_counter_label.cget("text") or spot)
            except tk.TclError:
                pass
        spot_num = self._parse_spot_number_from_label(spot)
        oh, os_ = self._wheel_obs_overlay_lines(
            spot_num, landed, winning_sku, len(self._wheel_strip)
        )
        out = dict(prev)
        out["scroll"] = scroll * hw
        out["cells"] = new_cells
        out["pulse"] = bool(self._wheel_pulse_highlight)
        out["obsHeadline"] = oh or spot
        out["obsSubhead"] = os_
        out["spotLabel"] = spot
        out["spotNumber"] = spot_num
        out["spotAssignment"] = ""
        out["prizesLeftLabel"] = self._wheel_prizes_left_label_text()
        return out

    def _wheel_publish_html_snapshot(self) -> None:
        try:
            prev = getattr(self, "_wheel_html_snapshot_cache", None)
            if (
                getattr(self, "_anim_after", None) is not None
                and isinstance(prev, dict)
                and prev.get("mode") == "strip"
            ):
                fast = self._wheel_html_snapshot_spin_frame(prev)
                if fast is not None:
                    self._wheel_html_snapshot_cache = fast
                    return
            self._wheel_html_snapshot_cache = self._wheel_build_html_snapshot()
        except Exception:
            try:
                pl_err = self._wheel_prizes_left_label_text()
            except Exception:
                pl_err = "Total prizes left: —"
            self._wheel_html_snapshot_cache = {
                "mode": "banner",
                "chroma": OBS_CHROMA_KEY,
                "colors": {},
                "spotLabel": "",
                "spotNumber": None,
                "spotAssignment": "",
                "obsHeadline": "",
                "obsSubhead": "",
                "statusText": "",
                "bannerTitle": "Wheel snapshot error",
                "bannerSubtitle": "",
                "bannerAccent": "error",
                "prizesLeftLabel": pl_err,
            }

    def _on_ui_click_defocus_wheel_name_entry(self, event: tk.Event) -> None:
        """If the wheel name field has focus, move focus away when the user clicks another widget."""
        ent = getattr(self, "_new_wheel_name_entry", None)
        if ent is None or event.widget == ent:
            return
        try:
            if self.focus_get() != ent:
                return
        except tk.TclError:
            return
        try:
            event.widget.focus_set()
        except tk.TclError:
            pass
        try:
            if self.focus_get() == ent:
                cw = getattr(self, "_controls_win", None)
                if cw is not None:
                    try:
                        if cw.winfo_exists():
                            c = getattr(cw, "_spin_controls_scroll_canvas", None)
                            if c is not None:
                                c.focus_set()
                    except tk.TclError:
                        pass
        except tk.TclError:
            pass

    def _build_spin_controls_window(self) -> None:
        if self._controls_win is not None:
            try:
                if self._controls_win.winfo_exists():
                    return
            except tk.TclError:
                pass
            self._controls_win = None
        self._controls_win = SpinControlsWindow(self)

    def _show_spin_controls(self) -> None:
        cw = self._controls_win
        if cw is not None:
            try:
                if cw.winfo_exists():
                    cw.lift()
                    return
            except tk.TclError:
                self._controls_win = None
        self._controls_win = SpinControlsWindow(self)

    def _build_prize_list_window(self) -> None:
        if self._prize_win is not None:
            try:
                if self._prize_win.winfo_exists():
                    return
            except tk.TclError:
                pass
            self._prize_win = None
        self._prize_win = PrizeListOverlay(self)

    def _show_prize_board(self) -> None:
        pw = self._prize_win
        if pw is not None:
            try:
                if pw.winfo_exists():
                    pw.lift()
                    return
            except tk.TclError:
                self._prize_win = None
        self._prize_win = PrizeListOverlay(self)
        self._sync_obs_overlay()

    def destroy(self) -> None:
        self._cancel_prize_board_resize_refresh()
        pw = self._prize_win
        if pw is not None:
            try:
                pw.destroy()
            except tk.TclError:
                pass
            self._prize_win = None
        cw = self._controls_win
        if cw is not None:
            try:
                cw.destroy()
            except tk.TclError:
                pass
            self._controls_win = None
        self._cancel_spin_controls_wheel_sync()
        self._wheel_control_link_ref = None
        self.remaining_scroll = None  # type: ignore[assignment]
        self.remaining_canvas = None  # type: ignore[assignment]
        self.remaining_inner = None  # type: ignore[assignment]
        self._remaining_canvas_win = None  # type: ignore[assignment]
        self._prize_board_totals_label = None
        self._wheel_shutdown_html_server()
        super().destroy()

    def _build_drag_bar(self) -> None:
        self._drag_bar = tk.Frame(self, bg=DRAG_BAR, height=28)
        self._drag_bar.pack(fill=tk.X)
        tk.Button(
            self._drag_bar,
            text=" ✕ ",
            command=self.destroy,
            bg="#c0392b",
            fg="white",
            activebackground="#e74c3c",
            font=("Segoe UI", 10, "bold"),
            relief=tk.FLAT,
            padx=10,
        ).pack(side=tk.RIGHT, padx=4, pady=2)
        drag_lbl = tk.Label(
            self._drag_bar,
            text="  ◎ Energy Break — drag to move  ",
            bg=DRAG_BAR,
            fg="#1a1a1a",
            font=("Segoe UI", 10, "bold"),
        )
        drag_lbl.pack(side=tk.LEFT)
        for w in (self._drag_bar, drag_lbl):
            w.bind("<ButtonPress-1>", self._drag_start)
            w.bind("<B1-Motion>", self._drag_motion)

    def _drag_start(self, event: tk.Event) -> None:
        self._drag_offset = (event.x_root - self.winfo_rootx(), event.y_root - self.winfo_rooty())

    def _drag_motion(self, event: tk.Event) -> None:
        x = event.x_root - self._drag_offset[0]
        y = event.y_root - self._drag_offset[1]
        self.geometry(f"{self.winfo_width()}x{self.winfo_height()}+{x}+{y}")

    def _toggle_setup(self) -> None:
        if self._setup_expanded.get():
            self.setup_frame.pack(fill=tk.X, padx=8, pady=(0, 4), before=self._wheel_host)
        else:
            self.setup_frame.pack_forget()

    def _toggle_log(self) -> None:
        if self._log_visible.get():
            self.log_frame.pack(fill=tk.BOTH, expand=False, pady=(4, 0))
        else:
            self.log_frame.pack_forget()

    def _build_setup_form(self, parent: tk.Frame) -> None:
        parent.configure(bg=self._void_bg())
        ttk.Label(
            parent,
            text="List file (SKU, Qty, img): third column = picture path (images/… or full path) or image URL — shown on wheel cells; Pillow required for most formats.",
        ).pack(anchor=tk.W, padx=4)

        self._setup_path_row = tk.Frame(parent, bg=self._void_bg())
        self._setup_path_row.pack(fill=tk.X, pady=(4, 6), padx=4)
        self._register_void_bg(self._setup_path_row)
        self.path_var = tk.StringVar(value=str(draw_prize.default_list_path()))
        self.path_var.trace_add("write", self._on_path_write)
        self.path_entry = ttk.Entry(self._setup_path_row, textvariable=self.path_var, width=70)
        self.path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        ttk.Button(self._setup_path_row, text="Browse…", command=self._browse).pack(side=tk.RIGHT)

        self._setup_row2 = tk.Frame(parent, bg=self._void_bg())
        self._setup_row2.pack(fill=tk.X, padx=4, pady=(0, 6))
        self._register_void_bg(self._setup_row2)
        self.dry_run_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(self._setup_row2, text="Dry run (never write file)", variable=self.dry_run_var).pack(
            side=tk.LEFT, padx=(0, 16)
        )
        ttk.Checkbutton(self._setup_row2, text="Show event log", variable=self._log_visible, command=self._toggle_log).pack(
            side=tk.LEFT
        )

        self.log_frame = ttk.Frame(parent)
        ttk.Label(self.log_frame, text="Log:").pack(anchor=tk.W)
        self.out = scrolledtext.ScrolledText(self.log_frame, height=6, wrap=tk.WORD, font=("Consolas", 9))
        self.out.pack(fill=tk.BOTH, expand=True, pady=(2, 0))

    def _build_spin_controls_inner(self, ctrl_inner: tk.Frame) -> None:
        """Populate the Spin & controls window (also used when recreating that Toplevel)."""
        panel = PRIZE_BOARD_CONTENT_BG
        card_bg = SPIN_CONTROLS_CARD_BG
        edge = SPIN_CONTROLS_CARD_BORDER
        ctrl_inner.configure(bg=panel)

        def _section_title(text: str) -> None:
            tk.Label(
                ctrl_inner,
                text=text,
                font=("Segoe UI", 10, "bold"),
                bg=panel,
                fg=WHEEL_POINTER,
                anchor=tk.W,
            ).pack(fill=tk.X, pady=(4, 8))

        _section_title("Draw")
        card_draw = tk.Frame(ctrl_inner, bg=card_bg, highlightbackground=edge, highlightthickness=1)
        card_draw.pack(fill=tk.X, pady=(0, 14))
        draw_pad = tk.Frame(card_draw, bg=card_bg)
        draw_pad.pack(fill=tk.X, padx=14, pady=14)

        tk.Label(
            draw_pad,
            text=(
                "SPIN — Run the Wheel; the Prize is committed.\n"
                "SUPER SPIN — Run the Wheel; Allowed to Re-Roll Once. "
                "REROLL and KEEP appear in this row only after a Super spin stops."
            ),
            bg=card_bg,
            fg=WHEEL_FG,
            font=("Segoe UI", 9),
            wraplength=640,
            justify=tk.LEFT,
            anchor=tk.W,
        ).pack(fill=tk.X, pady=(0, 12))

        self.control_bar = tk.Frame(draw_pad, bg=card_bg)
        self.control_bar.pack(fill=tk.X)

        btn_kw: dict[str, object] = {
            "font": self._btn_font,
            "relief": tk.FLAT,
            "bd": 0,
            "padx": 16,
            "pady": 12,
            "cursor": "hand2",
        }
        self.spin_btn = tk.Button(
            self.control_bar,
            text="  SPIN  ",
            command=self._on_spin,
            bg=BTN_SPIN,
            fg="white",
            activebackground="#3498db",
            activeforeground="white",
            **btn_kw,
        )
        self.spin_btn.pack(side=tk.LEFT, padx=(0, 10))

        self.super_btn = tk.Button(
            self.control_bar,
            text="  SUPER SPIN  ",
            command=self._on_super_spin,
            bg=BTN_SUPER,
            fg="white",
            activebackground="#9b59b6",
            activeforeground="white",
            **btn_kw,
        )
        self.super_btn.pack(side=tk.LEFT, padx=(0, 10))

        self.reroll_btn = tk.Button(
            self.control_bar,
            text="  REROLL  ",
            command=self._on_reroll,
            bg=BTN_REROLL,
            fg="white",
            activebackground="#1abc9c",
            activeforeground="white",
            **btn_kw,
        )
        self.keep_btn = tk.Button(
            self.control_bar,
            text="  KEEP  ",
            command=self._on_keep,
            bg=BTN_KEEP,
            fg="white",
            activebackground="#e74c3c",
            activeforeground="white",
            **btn_kw,
        )

        _section_title("Edit spin")
        card_edit = tk.Frame(ctrl_inner, bg=card_bg, highlightbackground=edge, highlightthickness=1)
        card_edit.pack(fill=tk.X, pady=(0, 14))
        edit_pad = tk.Frame(card_edit, bg=card_bg)
        edit_pad.pack(fill=tk.X, padx=14, pady=14)
        tk.Label(
            edit_pad,
            text=(
                "Skip spot — Skip a # in the Saved Sequence (use when payment fails or pending).\n"
                "Undo spin — If you accidentally spun, press this button."
            ),
            bg=card_bg,
            fg=WHEEL_FG,
            font=("Segoe UI", 9),
            wraplength=640,
            justify=tk.LEFT,
            anchor=tk.W,
        ).pack(fill=tk.X, pady=(0, 12))
        edit_btn_row = tk.Frame(edit_pad, bg=card_bg)
        edit_btn_row.pack(fill=tk.X)
        self.skip_spot_btn = tk.Button(
            edit_btn_row,
            text="  Skip spot  ",
            command=self._on_skip_spot,
            font=("Segoe UI", 10, "bold"),
            bg=BTN_SKIP,
            fg="white",
            activebackground="#eb984e",
            activeforeground="white",
            relief=tk.FLAT,
            bd=0,
            padx=14,
            pady=10,
            cursor="hand2",
        )
        self.skip_spot_btn.pack(side=tk.LEFT, padx=(0, 10))

        self.undo_spin_btn = tk.Button(
            edit_btn_row,
            text="  Undo spin  ",
            command=self._on_undo_spin,
            font=("Segoe UI", 10, "bold"),
            bg=BTN_UNDO,
            fg="white",
            activebackground="#2874a6",
            activeforeground="white",
            state=tk.DISABLED,
            relief=tk.FLAT,
            bd=0,
            padx=14,
            pady=10,
            cursor="hand2",
        )
        self.undo_spin_btn.pack(side=tk.LEFT, padx=(0, 10))

        _section_title("Fill skipped spot")
        card_fill = tk.Frame(ctrl_inner, bg=card_bg, highlightbackground=edge, highlightthickness=1)
        card_fill.pack(fill=tk.X, pady=(0, 14))
        fill_pad = tk.Frame(card_fill, bg=card_bg)
        fill_pad.pack(fill=tk.X, padx=14, pady=14)
        tk.Label(
            fill_pad,
            text=(
                "Skipped spot still empty? Choose the spot, then run one normal SPIN — the prize is written into "
                "that existing row in the winner session file. Cancel fill mode to go back to the latest spot #."
            ),
            bg=card_bg,
            fg=WHEEL_MUTED,
            font=("Segoe UI", 9),
            wraplength=640,
            justify=tk.LEFT,
            anchor=tk.W,
        ).pack(fill=tk.X, pady=(0, 10))
        fill_btns = tk.Frame(fill_pad, bg=card_bg)
        fill_btns.pack(fill=tk.X)
        tk.Button(
            fill_btns,
            text="  Fill a skipped spot…  ",
            command=self._open_fill_skipped_spot_dialog,
            font=("Segoe UI", 10, "bold"),
            bg=WHEEL_CELL_BG,
            fg=WHEEL_FG,
            activebackground=WHEEL_POINTER,
            activeforeground="#1a1a2e",
            relief=tk.FLAT,
            bd=0,
            padx=12,
            pady=8,
            cursor="hand2",
        ).pack(side=tk.LEFT, padx=(0, 8))
        tk.Button(
            fill_btns,
            text="  Cancel fill mode  ",
            command=self.cancel_fill_skipped_spot_mode,
            font=("Segoe UI", 10),
            bg=WHEEL_CELL_BG,
            fg=WHEEL_FG,
            activebackground=WHEEL_POINTER,
            activeforeground="#1a1a2e",
            relief=tk.FLAT,
            bd=0,
            padx=12,
            pady=8,
            cursor="hand2",
        ).pack(side=tk.LEFT)

        _section_title("HTML wheel (OBS)")
        card_html = tk.Frame(ctrl_inner, bg=card_bg, highlightbackground=edge, highlightthickness=1)
        card_html.pack(fill=tk.X, pady=(0, 14))
        html_pad = tk.Frame(card_html, bg=card_bg)
        html_pad.pack(fill=tk.X, padx=14, pady=12)
        tk.Label(
            html_pad,
            text="Browser wheel for OBS (localhost). Use the same URL in a Browser Source after the app prints the port.",
            bg=card_bg,
            fg=WHEEL_MUTED,
            font=("Segoe UI", 9),
            wraplength=520,
            justify=tk.LEFT,
            anchor=tk.W,
        ).pack(fill=tk.X, pady=(0, 10))
        html_row = tk.Frame(html_pad, bg=card_bg)
        html_row.pack(fill=tk.X)
        tk.Button(
            html_row,
            text="  Open HTML wheel…  ",
            command=self._open_html_wheel_in_browser,
            font=("Segoe UI", 10, "bold"),
            bg=WHEEL_CELL_BG,
            fg=WHEEL_FG,
            activebackground=WHEEL_POINTER,
            activeforeground="#1a1a2e",
            relief=tk.FLAT,
            bd=0,
            padx=14,
            pady=10,
            cursor="hand2",
        ).pack(side=tk.LEFT, padx=(0, 10))

        _section_title("New wheel")
        card_misc = tk.Frame(ctrl_inner, bg=card_bg, highlightbackground=edge, highlightthickness=1)
        card_misc.pack(fill=tk.X)
        misc_pad = tk.Frame(card_misc, bg=card_bg)
        misc_pad.pack(fill=tk.X, padx=14, pady=14)
        tk.Label(
            misc_pad,
            text="When you want to run a new wheel, press Start new wheel. That creates a new Excel file for the next saves.",
            bg=card_bg,
            fg=WHEEL_FG,
            font=("Segoe UI", 9),
            wraplength=560,
            justify=tk.LEFT,
            anchor=tk.W,
        ).pack(fill=tk.X, pady=(0, 10))
        name_row = tk.Frame(misc_pad, bg=card_bg)
        name_row.pack(fill=tk.X, pady=(0, 8))
        tk.Label(
            name_row,
            text="Wheel name",
            bg=card_bg,
            fg=WHEEL_FG,
            font=("Segoe UI", 9, "bold"),
        ).pack(side=tk.LEFT, padx=(0, 8))
        self._new_wheel_name_entry = tk.Entry(
            name_row,
            textvariable=self._new_wheel_name_var,
            font=("Segoe UI", 10),
            width=28,
            bg=INV_SLOT_FACE,
            fg=WHEEL_FG,
            insertbackground=WHEEL_FG,
            highlightthickness=1,
            highlightbackground=edge,
            highlightcolor=WHEEL_POINTER,
        )
        self._new_wheel_name_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Label(
            misc_pad,
            text="Optional — letters and numbers are turned into part of the new file name (for example, Friday becomes winners_Friday_….xlsx).",
            bg=card_bg,
            fg=WHEEL_MUTED,
            font=("Segoe UI", 8),
            wraplength=560,
            justify=tk.LEFT,
            anchor=tk.W,
        ).pack(fill=tk.X, pady=(0, 10))
        tk.Button(
            misc_pad,
            text="  Start new wheel…  ",
            command=self._start_new_wheel,
            font=("Segoe UI", 10, "bold"),
            bg=WHEEL_CELL_BG,
            fg=WHEEL_FG,
            activebackground=WHEEL_POINTER,
            activeforeground="#1a1a2e",
            relief=tk.FLAT,
            bd=0,
            padx=14,
            pady=10,
            cursor="hand2",
        ).pack(anchor=tk.W)

        _section_title("Prize board")
        card_pb = tk.Frame(ctrl_inner, bg=card_bg, highlightbackground=edge, highlightthickness=1)
        card_pb.pack(fill=tk.X, pady=(0, 14))
        pb_pad = tk.Frame(card_pb, bg=card_bg)
        pb_pad.pack(fill=tk.X, padx=14, pady=12)
        pb_row = tk.Frame(pb_pad, bg=card_bg)
        pb_row.pack(fill=tk.X)
        tk.Label(
            pb_row,
            text="Board:",
            bg=card_bg,
            fg=WHEEL_FG,
            font=("Segoe UI", 9, "bold"),
        ).pack(side=tk.LEFT, padx=(0, 6))
        tk.Label(pb_row, text="columns", bg=card_bg, fg=WHEEL_MUTED, font=("Segoe UI", 9)).pack(side=tk.LEFT)
        tk.Spinbox(
            pb_row,
            from_=2,
            to=24,
            width=4,
            textvariable=self._prize_board_grid_cols,
            font=("Segoe UI", 10),
            bg=INV_SLOT_FACE,
            fg=WHEEL_FG,
            buttonbackground=WHEEL_CELL_BG,
            highlightthickness=0,
        ).pack(side=tk.LEFT, padx=(4, 10))
        tk.Label(pb_row, text="×", bg=card_bg, fg=WHEEL_MUTED, font=("Segoe UI", 10)).pack(side=tk.LEFT)
        tk.Spinbox(
            pb_row,
            from_=2,
            to=20,
            width=4,
            textvariable=self._prize_board_tile_rows_fit,
            font=("Segoe UI", 10),
            bg=INV_SLOT_FACE,
            fg=WHEEL_FG,
            buttonbackground=WHEEL_CELL_BG,
            highlightthickness=0,
        ).pack(side=tk.LEFT, padx=(4, 6))
        tk.Label(
            pb_row,
            text="rows (tile size)",
            bg=card_bg,
            fg=WHEEL_MUTED,
            font=("Segoe UI", 9),
        ).pack(side=tk.LEFT)
        tk.Label(
            pb_pad,
            text="Applies to the separate Prize board window (columns per row; rows set how tile height fits the panel).",
            bg=card_bg,
            fg=WHEEL_MUTED,
            font=("Segoe UI", 9),
            wraplength=520,
            justify=tk.LEFT,
            anchor=tk.W,
        ).pack(fill=tk.X, pady=(10, 0))

        self.btn_row = self.control_bar

    def _clear_undo_spin(self) -> None:
        self._undo_spin_snap = None
        self._sync_undo_spin_button()

    def _sync_undo_spin_button(self) -> None:
        if not hasattr(self, "undo_spin_btn"):
            return
        can = self._undo_spin_snap is not None and not self._busy
        if can:
            self.undo_spin_btn.config(state=tk.NORMAL, bg=BTN_UNDO)
        else:
            self.undo_spin_btn.config(state=tk.DISABLED, bg=BTN_DISABLED)

    def _push_undo_spin(
        self,
        list_path: Path,
        result: draw_prize.SpinResult,
        spot_before_append: int,
        winner_logged: bool,
    ) -> None:
        self._undo_spin_snap = _UndoSpinSnapshot(
            "spin",
            list_path.resolve(),
            result,
            spot_before_append,
            winner_logged,
        )
        self._sync_undo_spin_button()

    def _push_undo_skip_log(self, spot_written: int) -> None:
        try:
            lp = self._list_path().resolve()
        except OSError:
            lp = Path()
        self._undo_spin_snap = _UndoSpinSnapshot(
            "skip_log",
            lp,
            None,
            spot_written,
            True,
        )
        self._sync_undo_spin_button()

    def _push_undo_skip_counter_only(self, spot_before: int) -> None:
        try:
            lp = self._list_path().resolve()
        except OSError:
            lp = Path()
        self._undo_spin_snap = _UndoSpinSnapshot(
            "skip_counter_only",
            lp,
            None,
            spot_before,
            False,
        )
        self._sync_undo_spin_button()

    def _push_undo_backfill(
        self,
        list_path: Path,
        result: draw_prize.SpinResult,
        spot_filled: int,
        sheet_updated: bool,
    ) -> None:
        self._undo_spin_snap = _UndoSpinSnapshot(
            "backfill",
            list_path.resolve(),
            result,
            spot_filled,
            sheet_updated,
        )
        self._sync_undo_spin_button()

    def _winner_delete_row_for_spot(self, path: Path, spot_number: int) -> bool:
        """Remove the data row whose column A equals ``spot_number`` (last match if duplicated)."""
        if load_workbook is None or Workbook is None:
            return False
        try:
            wb = load_workbook(path)
            ws = wb.active
            if ws is None:
                wb.close()
                return False
            delete_at: int | None = None
            max_r = ws.max_row or 1
            for r in range(2, max_r + 1):
                a = ws.cell(row=r, column=1).value
                try:
                    n = int(a) if a is not None and str(a).strip() != "" else None
                except (TypeError, ValueError):
                    n = None
                if n == spot_number:
                    delete_at = r
            if delete_at is None:
                wb.close()
                return False
            ws.delete_rows(delete_at, 1)
            wb.save(path)
            wb.close()
            return True
        except Exception:
            return False

    def _apply_undo_spin_snapshot(self, snap: _UndoSpinSnapshot) -> str | None:
        """Return an error message, or None on success."""
        try:
            if snap.kind == "skip_counter_only":
                self._winner_next_spin = snap.spot_written
                return None
            if snap.kind == "skip_log":
                wp = getattr(self, "_winner_log_path", None)
                if wp and snap.winner_row_added:
                    if not self._winner_delete_row_for_spot(Path(wp), snap.spot_written):
                        return "Could not remove the skip row from the winner workbook (close Excel if it is open)."
                self._winner_next_spin = snap.spot_written
                return None
            if snap.kind == "backfill":
                if snap.result is None:
                    return "Internal error: bad undo snapshot."
                if snap.list_path.resolve() != Path(self._list_path()).resolve():
                    return "The prize list path changed — undo aborted."
                sess = draw_prize.open_draw_session(snap.list_path)
                sess.revert_decrement(snap.result)
                if snap.winner_row_added:
                    wp = getattr(self, "_winner_log_path", None)
                    if wp and not self._winner_clear_prize_cell_for_spot(Path(wp), snap.spot_written):
                        return (
                            "Prize quantity was restored, but the winner sheet cell could not be cleared. "
                            "Clear column B for that spot manually if needed."
                        )
                if hasattr(self, "_winner_spot_lookup_cache"):
                    self._winner_spot_lookup_cache = None
                return None
            if snap.kind != "spin" or snap.result is None:
                return "Internal error: bad undo snapshot."
            if snap.list_path.resolve() != Path(self._list_path()).resolve():
                return "The prize list path changed — undo aborted."
            sess = draw_prize.open_draw_session(snap.list_path)
            sess.revert_decrement(snap.result)
            if snap.winner_row_added:
                wp = getattr(self, "_winner_log_path", None)
                if wp:
                    if not self._winner_delete_row_for_spot(Path(wp), snap.spot_written):
                        return (
                            "Prize quantity was restored, but the winner sheet row could not be removed. "
                            "Delete that row manually if needed."
                        )
            self._winner_next_spin = snap.spot_written
            return None
        except draw_prize.PrizeDrawError as e:
            return str(e)
        except Exception as e:
            return f"{type(e).__name__}: {e}"

    def _set_overlay_winner_list_height(self) -> None:
        """Cap listbox viewport to OVERLAY_WINNER_LIST_VISIBLE_ROWS; scroll when there are more rows."""
        if not hasattr(self, "_overlay_listbox"):
            return
        lb = self._overlay_listbox
        n = max(1, lb.size())
        cap = int(type(self).OVERLAY_WINNER_LIST_VISIBLE_ROWS)
        lb.configure(height=min(cap, n))

    def _on_undo_spin(self) -> None:
        if self._busy:
            return
        snap = self._undo_spin_snap
        if snap is None:
            return
        if not messagebox.askyesno(
            "Undo spin",
            "Reverse the last completed spin or skip?\n\n"
            "This restores the prize list quantity if it changed, removes the matching winner row if one was logged, "
            "and steps the spot counter back. Super / Reroll in progress will be cancelled.",
        ):
            return
        err = self._apply_undo_spin_snapshot(snap)
        if err:
            messagebox.showerror("Undo spin", err)
            self._log(f"Undo failed: {err}\n")
            return
        self._hide_super_panel()
        self._clear_undo_spin()
        if hasattr(self, "_winner_spot_lookup_cache"):
            self._winner_spot_lookup_cache = None
        self._cancel_pulse()
        self._invalidate_session()
        self._update_spin_counter_label()
        self._reset_wheel_idle()
        self._refresh_prizes_label()
        self._update_draw_buttons_for_supply()
        self._wheel_publish_html_snapshot()
        self._log("Undo: last draw / skip reversed.\n")

    def _build_wheel_area(self, parent: tk.Frame) -> None:
        """HTML wheel is primary for OBS; Tk canvas lives off-screen for spin state and /api snapshots."""
        parent.configure(bg=self._void_bg())
        vb = self._void_bg()

        compact = tk.Frame(parent, bg=vb)
        compact.pack(fill=tk.X, expand=False, padx=12, pady=(12, 10))
        self._register_void_bg(compact)
        tk.Label(
            compact,
            text="Strip wheel + status run in Spin & controls and in the HTML page (OBS).",
            font=("Segoe UI", 10),
            bg=vb,
            fg=WHEEL_FG,
            wraplength=460,
            justify=tk.CENTER,
        ).pack(fill=tk.X, pady=(0, 12))
        row = tk.Frame(compact, bg=vb)
        row.pack(fill=tk.X)
        self._register_void_bg(row)
        self._prize_board_btn = tk.Button(
            row,
            text="  Prize board  ",
            command=self._show_prize_board,
            font=("Segoe UI", 10, "bold"),
            bg=WHEEL_CELL_BG,
            fg=WHEEL_FG,
            activebackground=WHEEL_POINTER,
            activeforeground="#1a1a2e",
            relief=tk.FLAT,
            padx=12,
            pady=6,
        )
        self._prize_board_btn.pack(side=tk.LEFT, padx=(0, 10))
        self._spin_controls_title_btn = tk.Button(
            row,
            text="  Spin & controls…  ",
            command=self._show_spin_controls,
            font=("Segoe UI", 10, "bold"),
            bg=WHEEL_CELL_BG,
            fg=WHEEL_FG,
            activebackground=WHEEL_POINTER,
            activeforeground="#1a1a2e",
            relief=tk.FLAT,
            padx=12,
            pady=6,
        )
        self._spin_controls_title_btn.pack(side=tk.LEFT)

        self._overlay_list_outer = tk.Frame(
            compact,
            bg=WHEEL_FRAME_BG,
            highlightbackground=WHEEL_CELL_BORDER,
            highlightthickness=1,
        )
        self._overlay_list_outer.pack(fill=tk.X, expand=False, pady=(10, 0))
        self._overlay_list_title = tk.Label(
            self._overlay_list_outer,
            text="Prize Wheel",
            bg=WHEEL_FRAME_BG,
            fg=WHEEL_TITLE,
            font=("Segoe UI", 10, "bold"),
            anchor=tk.W,
        )
        self._overlay_list_title.pack(fill=tk.X, padx=8, pady=(8, 4))
        list_row = tk.Frame(self._overlay_list_outer, bg=WHEEL_FRAME_BG)
        list_row.pack(fill=tk.X, padx=6, pady=(0, 8))
        self._overlay_list_scroll = tk.Scrollbar(
            list_row,
            orient=tk.VERTICAL,
            bg=WHEEL_FRAME_BG,
            troughcolor=WHEEL_CELL_BG,
        )
        self._overlay_listbox = tk.Listbox(
            list_row,
            height=1,
            font=("Segoe UI", 10),
            bg=WHEEL_CELL_BG,
            fg=WHEEL_FG,
            selectbackground=WHEEL_POINTER,
            selectforeground="#1a1a2e",
            highlightthickness=0,
            activestyle="dotbox",
            exportselection=False,
            yscrollcommand=self._overlay_list_scroll.set,
        )
        self._overlay_list_scroll.config(command=self._overlay_listbox.yview)
        self._overlay_list_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self._overlay_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        shw, shh = 1000, int(round(280 * WHEEL_VERTICAL_SCALE))
        self._wheel_sensor_host = tk.Frame(self, bg=vb, width=shw, height=shh)
        self._wheel_sensor_host.place(x=-3200, y=0)
        self._wheel_sensor_host.pack_propagate(False)
        self._register_void_bg(self._wheel_sensor_host)

        self._wheel_heading_block = tk.Frame(self._wheel_sensor_host, bg=vb)
        self._wheel_heading_block.pack(fill=tk.X, padx=10, pady=(8, 2))
        self._wheel_heading_block.grid_columnconfigure(0, weight=1)
        self._wheel_heading_block.grid_columnconfigure(1, weight=0)
        self._wheel_heading_block.grid_columnconfigure(2, weight=1)
        self._register_void_bg(self._wheel_heading_block)

        self._wheel_heading_label = tk.Label(
            self._wheel_heading_block,
            text="Prize Wheel",
            font=("Segoe UI", 20, "bold"),
            bg=vb,
            fg=WHEEL_TITLE,
            anchor=tk.CENTER,
        )
        self._wheel_heading_label.grid(row=0, column=1, sticky=tk.N)
        self._register_void_bg(self._wheel_heading_label)

        self.prizes_label = tk.Label(
            self._wheel_heading_block,
            text="Total prizes: —",
            font=("Segoe UI", 17, "bold"),
            bg=vb,
            fg=WHEEL_POINTER,
            anchor=tk.E,
            justify=tk.RIGHT,
        )
        self.prizes_label.grid(row=0, column=2, sticky=tk.NE)
        self._register_void_bg(self.prizes_label)

        self.spin_counter_label = tk.Label(
            self._wheel_sensor_host,
            text="Spot 1",
            font=("Segoe UI", 20, "bold"),
            bg=vb,
            fg=WHEEL_POINTER,
        )
        self.spin_counter_label.pack(fill=tk.X, padx=10, pady=(0, 2))
        self._register_void_bg(self.spin_counter_label)

        self.wheel_canvas = tk.Canvas(
            self._wheel_sensor_host,
            height=self.WHEEL_CELL,
            bg=vb,
            highlightthickness=0,
        )
        self.wheel_canvas.pack(fill=tk.BOTH, expand=True, padx=10, pady=(2, 4))
        self._register_void_bg(self.wheel_canvas)
        self.wheel_canvas.bind("<Configure>", self._on_wheel_canvas_configure)

        self.wheel_status = tk.Label(
            self._wheel_sensor_host,
            text="",
            font=("Segoe UI", 10, "italic"),
            bg=vb,
            fg=WHEEL_FG,
            wraplength=shw - 32,
            justify=tk.CENTER,
        )
        self.wheel_status.pack(pady=(0, 6))
        self._register_void_bg(self.wheel_status)

    def _on_wheel_canvas_configure(self, _event: tk.Event | None = None) -> None:
        """Keep the idle preview strip centered under the pointer when the canvas size changes."""
        if (
            not self._busy
            and self._anim_after is None
            and self._pulse_after is None
            and self._pending_super is None
            and len(self._wheel_strip)
            == int(type(self).WHEEL_IDLE_STRIP_LEN) * int(type(self).WHEEL_IDLE_CAROUSEL_COPIES)
        ):
            mid = self._wheel_win_idx
            cw_px = max(int(self.wheel_canvas.winfo_width()), 2)
            self._wheel_target_scroll = self._wheel_scroll_to_center_index(float(mid), float(cw_px))
            self._wheel_scroll = self._wheel_target_scroll
            self._wheel_idle_offset = 0.0
        self._wheel_redraw()

    def _cancel_wheel_idle_drift(self) -> None:
        aid = getattr(self, "_wheel_idle_drift_after", None)
        if aid is not None:
            try:
                self.after_cancel(aid)
            except tk.TclError:
                pass
            self._wheel_idle_drift_after = None
        self._wheel_idle_offset = 0.0

    def _wheel_idle_drift_active(self) -> bool:
        if self._busy or self._anim_after is not None or self._pulse_after is not None:
            return False
        if self._pending_super is not None:
            return False
        if getattr(self, "_customer_empty_wheel", False):
            return False
        if (
            len(self._wheel_strip)
            != int(type(self).WHEEL_IDLE_STRIP_LEN) * int(type(self).WHEEL_IDLE_CAROUSEL_COPIES)
        ):
            return False
        if self._wheel_banner_title and not self._wheel_strip:
            return False
        return True

    def _wheel_scroll_for_strip_render(self) -> float:
        s = float(self._wheel_scroll)
        if self._wheel_idle_drift_active():
            return s + float(getattr(self, "_wheel_idle_offset", 0.0) or 0.0)
        return s

    def _wheel_idle_drift_tick(self) -> None:
        self._wheel_idle_drift_after = None
        if not self._wheel_idle_drift_active():
            self._wheel_idle_offset = 0.0
            return
        period = float(type(self).WHEEL_IDLE_STRIP_LEN) * float(self.WHEEL_CELL)
        if period < 1.0:
            return
        ms = float(int(type(self).WHEEL_IDLE_DRIFT_MS))
        delta = float(type(self).WHEEL_IDLE_DRIFT_PX_PER_SEC) * (ms / 1000.0)
        self._wheel_idle_offset = (self._wheel_idle_offset + delta) % period
        self._wheel_redraw(_idle_drift_frame=True)
        if self._wheel_idle_drift_active():
            self._wheel_idle_drift_after = self.after(
                int(type(self).WHEEL_IDLE_DRIFT_MS),
                self._wheel_idle_drift_tick,
            )

    def _spin_controls_wheel_link_reset(self) -> None:
        self._cancel_spin_controls_wheel_sync()
        self._wheel_control_link_ref = None

    def _cancel_spin_controls_wheel_sync(self) -> None:
        aid = self._spin_controls_wheel_sync_after
        if aid is not None:
            try:
                self.after_cancel(aid)
            except tk.TclError:
                pass
            self._spin_controls_wheel_sync_after = None

    def _schedule_spin_controls_wheel_sync(self) -> None:
        self._cancel_spin_controls_wheel_sync()
        self._spin_controls_wheel_sync_after = self.after(75, self._flush_spin_controls_wheel_sync)

    def _flush_spin_controls_wheel_sync(self) -> None:
        self._spin_controls_wheel_sync_after = None
        cw = self._controls_win
        if cw is None or not hasattr(self, "wheel_canvas"):
            return
        try:
            if not cw.winfo_exists():
                return
        except tk.TclError:
            return
        try:
            cww = int(cw.winfo_width())
            ch = int(cw.winfo_height())
        except tk.TclError:
            return
        if cww < 400 or ch < 340:
            return
        h_min = int(type(self).WHEEL_CANVAS_H_LINK_MIN)
        h_max = int(type(self).WHEEL_CANVAS_H_LINK_MAX)
        ref = self._wheel_control_link_ref
        if ref is None:
            try:
                self.update_idletasks()
            except tk.TclError:
                pass
            try:
                cur_canvas_h = int(self.wheel_canvas.cget("height"))
            except tk.TclError:
                return
            self._wheel_control_link_ref = (ch, cur_canvas_h)
            return
        ref_ch, ref_canvas_h = ref
        new_canvas_h = ref_canvas_h + (ch - ref_ch)
        new_canvas_h = max(h_min, min(h_max, new_canvas_h))
        try:
            cur_h = int(self.wheel_canvas.cget("height"))
        except tk.TclError:
            return
        resized = False
        if new_canvas_h != cur_h:
            self._wheel_image_cache.clear()
            try:
                self.wheel_canvas.configure(height=new_canvas_h)
            except tk.TclError:
                return
            resized = True
            try:
                self._ensure_wheel_photos_for_strip(self._wheel_strip)
            except tk.TclError:
                pass
        if resized:
            self._on_wheel_canvas_configure()

    def _btn_set_busy(self, btn: tk.Button, busy: bool) -> None:
        if busy:
            btn.config(state=tk.DISABLED, bg=BTN_DISABLED)
        else:
            btn.config(state=tk.NORMAL)

    def _total_qty_remaining(self) -> int | None:
        try:
            p = self._list_path()
        except OSError:
            return None
        if not p.is_file():
            return None
        try:
            _, rows = draw_prize.load_rows(p)
            return sum(q for _, q, _ in rows)
        except (draw_prize.PrizeDrawError, OSError, ValueError, TypeError):
            return None

    def _winnable_units_total(self) -> int | None:
        """Units still in the draw (sum of positive quantities). None if the list file cannot be read."""
        try:
            p = self._list_path()
        except OSError:
            return None
        if not p.is_file():
            return None
        try:
            _, rows = draw_prize.load_rows(p)
            return sum(q for _, q, _ in rows if q > 0)
        except (draw_prize.PrizeDrawError, OSError, ValueError, TypeError):
            return None

    def _style_buttons_depleted(self) -> None:
        if not hasattr(self, "spin_btn"):
            return
        self._hide_super_reroll_keep_buttons()
        self.spin_btn.config(state=tk.DISABLED, bg=BTN_DISABLED)
        self.super_btn.config(state=tk.DISABLED, bg=BTN_DISABLED)
        self.skip_spot_btn.config(state=tk.NORMAL, bg=BTN_SKIP)
        self._sync_undo_spin_button()

    def _wheel_paint_depleted_customer(self) -> None:
        """Clear the strip wheel and show a friendly message when nothing is left to draw."""
        self._cancel_wheel_idle_drift()
        self._wheel_sku_to_img = {}
        self._wheel_strip = []
        self._wheel_banner_title = "All prizes have been given out"
        self._wheel_banner_subtitle = (
            "Add stock to your prize list file, then save — new draws will unlock automatically."
        )
        self._wheel_banner_is_error = False
        self.wheel_canvas.delete("all")
        w = max(int(self.wheel_canvas.winfo_width()), 400)
        h = int(self.wheel_canvas.cget("height"))
        self.wheel_canvas.create_text(
            w / 2,
            h / 2 - 10,
            text="All prizes have been given out",
            fill=WHEEL_STRIP_WIN_RING,
            font=self._title_font,
            width=w - 40,
        )
        self.wheel_canvas.create_text(
            w / 2,
            h / 2 + 22,
            text="Add stock to your prize list file, then save — new draws will unlock automatically.",
            fill=WHEEL_STRIP_FG,
            font=("Segoe UI", 10),
            width=w - 48,
        )
        self._wheel_draw_strip_pointer(self.wheel_canvas, w / 2.0, float(WHEEL_STRIP_PAD_TOP))
        self.wheel_status.configure(
            text="There are no prizes left on the wheel. Reload your list with more quantities to keep going.",
            fg=WHEEL_TITLE,
        )
        self._wheel_publish_html_snapshot()

    def _paint_normal_idle_wheel_only(self) -> None:
        """Idle home strip: show real winnable SKUs (with images when available), not placeholder dashes."""
        self._wheel_banner_title = None
        self._wheel_banner_subtitle = None
        self._wheel_banner_is_error = False
        n = self.WHEEL_IDLE_STRIP_LEN
        mid = n // 2
        cw_px = max(int(self.wheel_canvas.winfo_width()), 520)
        self._wheel_win_idx = mid
        self._wheel_target_scroll = self._wheel_scroll_to_center_index(float(mid), float(cw_px))
        self._wheel_scroll = self._wheel_target_scroll

        def _fallback_strip() -> None:
            self._wheel_sku_to_img = {}
            ln = n * int(type(self).WHEEL_IDLE_CAROUSEL_COPIES)
            self._wheel_strip = ["···"] * ln

        try:
            p = self._list_path()
        except OSError:
            _fallback_strip()
            self.wheel_status.configure(
                text="SPIN = save now · SUPER SPIN = Reroll/Keep buttons appear after the wheel stops (no save until Keep)",
                fg=WHEEL_FG,
            )
            self.after_idle(self._wheel_redraw)
            return
        if not p.is_file():
            _fallback_strip()
            self.wheel_status.configure(
                text="SPIN = save now · SUPER SPIN = Reroll/Keep buttons appear after the wheel stops (no save until Keep)",
                fg=WHEEL_FG,
            )
            self.after_idle(self._wheel_redraw)
            return
        try:
            _, rows = draw_prize.load_rows(p)
        except draw_prize.PrizeDrawError:
            _fallback_strip()
            self.wheel_status.configure(
                text="SPIN = save now · SUPER SPIN = Reroll/Keep buttons appear after the wheel stops (no save until Keep)",
                fg=WHEEL_FG,
            )
            self.after_idle(self._wheel_redraw)
            return

        self._wheel_sku_to_img = self._sku_to_img_map_from_rows(rows)
        first = draw_prize.sample_wheel_strip_labels(rows, n)
        if not first:
            self._wheel_strip = ["···"] * (n * int(type(self).WHEEL_IDLE_CAROUSEL_COPIES))
            self._wheel_sku_to_img = {}
        else:
            self._wheel_strip = first + first
            self._ensure_wheel_photos_for_strip(self._wheel_strip)

        self.wheel_status.configure(
            text="SPIN = save now · SUPER SPIN = Reroll/Keep buttons appear after the wheel stops (no save until Keep)",
            fg=WHEEL_FG,
        )
        self.after_idle(self._wheel_redraw)

    def _update_draw_buttons_for_supply(self) -> None:
        if not hasattr(self, "spin_btn") or self._busy:
            return
        units = self._winnable_units_total()
        depleted = units is not None and units == 0
        if depleted:
            self._style_buttons_depleted()
        elif self._pending_super is not None:
            self._style_super_controls_active()
        else:
            self._style_main_controls_idle()

    def _reconcile_depleted_wheel_if_applicable(self) -> None:
        if not hasattr(self, "wheel_canvas"):
            return
        if self._busy or self._pending_super is not None or self._pulse_after is not None:
            return
        units = self._winnable_units_total()
        depleted = units is not None and units == 0
        if depleted:
            self._wheel_paint_depleted_customer()
            self._customer_empty_wheel = True
        elif self._customer_empty_wheel:
            self._customer_empty_wheel = False
            self._paint_normal_idle_wheel_only()

    def _update_prize_board_totals_label(
        self,
        rows_data: list[tuple[str, int, str]] | None,
        msg: str | None,
    ) -> None:
        lab = getattr(self, "_prize_board_totals_label", None)
        if lab is None:
            return
        try:
            if not lab.winfo_exists():
                return
        except tk.TclError:
            return
        if msg is not None:
            lab.configure(text="—")
            return
        if not rows_data:
            lab.configure(text="0 items available")
            return
        total_units = sum(q for _, q, _ in rows_data)
        n_types = len(rows_data)
        p = "prize types" if n_types != 1 else "prize type"
        lab.configure(text=f"{total_units} items available · {n_types} {p}")

    def _refresh_remaining_skus_panel(self) -> None:
        ri = getattr(self, "remaining_inner", None)
        if ri is None:
            self._update_prize_board_totals_label(None, "—")
            return
        try:
            if not ri.winfo_exists():
                return
        except tk.TclError:
            return
        for w in ri.winfo_children():
            w.destroy()
        msg: str | None = None
        rows_data: list[tuple[str, int, str]] = []
        try:
            p = self._list_path()
        except OSError:
            msg = "—"
        else:
            if not p.is_file():
                msg = "List file not found."
            else:
                try:
                    _, rows = draw_prize.load_rows(p)
                except draw_prize.PrizeDrawError as e:
                    msg = str(e)
                else:
                    rows_data = [(sku.strip(), q, img) for sku, q, img in rows if q > 0]
                    rows_data.sort(key=lambda t: t[0].lower())
        vb = self._prize_board_surface_bg()
        if msg is not None:
            self._update_prize_board_totals_label(None, msg)
            lb = tk.Label(
                ri,
                text=msg,
                font=("Segoe UI", 9),
                bg=vb,
                fg=WHEEL_MUTED,
            )
            lb.pack(anchor=tk.W, padx=4, pady=6)
            return
        if not rows_data:
            self._update_prize_board_totals_label([], None)
            lb = tk.Label(
                ri,
                text="No SKUs with qty left (all are 0).",
                font=("Segoe UI", 9),
                bg=vb,
                fg=WHEEL_MUTED,
            )
            lb.pack(anchor=tk.W, padx=4, pady=6)
            return

        slot, thumb, cols = self._prize_board_slot_thumb_for_inner(ri)
        layout_key = (slot, thumb, cols)
        if self._prize_board_last_layout != layout_key:
            self._inventory_image_cache.clear()
            self._prize_board_last_layout = layout_key

        inset = INV_SLOT_INSET
        inner_side = slot - 2 * inset
        g = INV_GRID_GAP
        sku_font = ("Segoe UI", max(7, min(11, slot // 8)), "bold")
        qty_font = ("Segoe UI", max(9, min(14, slot // 6)), "bold")
        for idx, (sku, q, img) in enumerate(rows_data):
            r, c = divmod(idx, cols)
            outer = tk.Frame(
                ri,
                width=slot,
                height=slot,
                bg=INV_SLOT_EDGE,
                highlightthickness=0,
            )
            outer.grid(row=r, column=c, padx=g, pady=g, sticky=tk.N)
            outer.grid_propagate(False)
            face = tk.Frame(outer, bg=INV_SLOT_FACE)
            face.place(x=inset, y=inset, width=inner_side, height=inner_side)
            refn = _normalize_image_ref(img)
            ph = self._inventory_photo_for_normalized_ref(refn, thumb) if refn else None
            if ph is not None:
                img_lbl = tk.Label(face, image=ph, bg=INV_SLOT_FACE)
                img_lbl.pack(expand=True, pady=(max(2, slot // 14), 2))
            else:
                short = sku if len(sku) <= 12 else sku[:11] + "…"
                tk.Label(
                    face,
                    text=short,
                    font=sku_font,
                    bg=INV_SLOT_FACE,
                    fg=WHEEL_FG,
                    wraplength=max(24, inner_side - 12),
                    justify=tk.CENTER,
                ).pack(expand=True, padx=4, pady=4)
            qty = tk.Label(
                outer,
                text=str(q),
                font=qty_font,
                bg=INV_QTY_BADGE_BG,
                fg=INV_QTY_BADGE_FG,
                padx=max(4, slot // 14),
                pady=1,
            )
            qty.place(relx=1.0, rely=1.0, anchor=tk.SE, x=-3, y=-3)
        self._update_prize_board_totals_label(rows_data, None)

    def _refresh_prizes_label(self) -> None:
        if not hasattr(self, "prizes_label"):
            return
        n = self._total_qty_remaining()
        if n is None:
            self.prizes_label.configure(text="Total prizes: —")
        else:
            self.prizes_label.configure(text=f"Total prizes: {n}")
        self._update_draw_buttons_for_supply()
        self._reconcile_depleted_wheel_if_applicable()
        self._refresh_remaining_skus_panel()
        self._refresh_overlay_winner_session_list()
        if not self._winner_overlay_poll_started:
            self._winner_overlay_poll_started = True
            self._schedule_winner_overlay_poll()
        self._wheel_publish_html_snapshot()

    def _winner_session_workbook_for_overlay(self) -> Path | None:
        """Workbook shown in the overlay: active session file, else newest under winner_sessions/."""
        p = getattr(self, "_winner_log_path", None)
        if p is not None:
            pp = Path(p)
            if pp.is_file():
                return pp
        base = script_dir() / "winner_sessions"
        return self._find_latest_winner_session_path(base)

    def _schedule_winner_overlay_poll(self) -> None:
        """Poll winner workbook mtime so external edits still refresh the overlay list."""
        aid = self._winner_overlay_poll_after
        if aid is not None:
            try:
                self.after_cancel(aid)
            except tk.TclError:
                pass
        self._winner_overlay_poll_after = self.after(2500, self._winner_overlay_poll_tick)

    def _winner_overlay_poll_tick(self) -> None:
        self._winner_overlay_poll_after = None
        if hasattr(self, "_overlay_listbox"):
            try:
                self._maybe_refresh_winner_session_overlay_from_mtime()
            except Exception:
                pass
        self._schedule_winner_overlay_poll()

    def _maybe_refresh_winner_session_overlay_from_mtime(self) -> None:
        path = self._winner_session_workbook_for_overlay()
        if path is None or not path.is_file():
            sig: tuple[str, float] = ("", 0.0)
        else:
            try:
                sig = (str(path.resolve()), path.stat().st_mtime)
            except OSError:
                sig = ("", 0.0)
        if sig == self._winner_overlay_last_sig:
            return
        self._winner_overlay_last_sig = sig
        self._refresh_overlay_winner_session_list()

    def _refresh_overlay_winner_session_list(self) -> None:
        """Operator-facing rows from the latest / active winner session Excel (main overlay tab).

        Rows are shown in reverse sheet order (typically highest spot / most recent at the top).
        """
        if not hasattr(self, "_overlay_listbox"):
            return
        lb = self._overlay_listbox
        title = getattr(self, "_overlay_list_title", None)
        try:
            lb.delete(0, tk.END)
            path = self._winner_session_workbook_for_overlay()
            if title is not None:
                if path is not None and path.is_file():
                    title.configure(text=f"Prize Wheel — {path.name}")
                else:
                    title.configure(text="Prize Wheel (no file yet)")
            if load_workbook is None:
                lb.insert(tk.END, "Install openpyxl to view the wheel log.")
                self._set_overlay_winner_list_height()
                self._resize_overlay_to_fit_list()
                return
            if path is None or not path.is_file():
                lb.insert(tk.END, "(No winners_*.xlsx in winner_sessions yet)")
                self._set_overlay_winner_list_height()
                self._resize_overlay_to_fit_list()
                return
            try:
                wb = load_workbook(path, read_only=True, data_only=True)
                try:
                    ws = wb.active
                    if ws is None:
                        lb.insert(tk.END, "(Empty workbook)")
                    else:
                        lines: list[str] = []
                        for row in ws.iter_rows(min_row=2, max_col=2, values_only=True):
                            a = row[0] if row else None
                            b = row[1] if row and len(row) > 1 else None
                            if a is None and (b is None or str(b).strip() == ""):
                                continue
                            try:
                                if a is not None and str(a).strip() != "":
                                    spot_s = f"Spot {int(a)}"
                                else:
                                    spot_s = str(a).strip() if a is not None else "—"
                            except (TypeError, ValueError):
                                spot_s = str(a).strip() if a is not None else "—"
                            if b is None or (isinstance(b, str) and not b.strip()) or b == "":
                                prize = "—"
                            else:
                                prize = str(b).strip()
                            lines.append(f"{spot_s}  —  {prize}")
                        if not lines:
                            lb.insert(tk.END, "(No winners logged yet in this session)")
                        else:
                            for line in reversed(lines):
                                lb.insert(tk.END, line)
                finally:
                    wb.close()
            except Exception as e:
                lb.insert(tk.END, f"Could not read workbook: {e}")
            self._set_overlay_winner_list_height()
            self._resize_overlay_to_fit_list()
        finally:
            try:
                p2 = self._winner_session_workbook_for_overlay()
                if p2 is not None and p2.is_file():
                    self._winner_overlay_last_sig = (str(p2.resolve()), p2.stat().st_mtime)
                else:
                    self._winner_overlay_last_sig = ("", 0.0)
            except OSError:
                self._winner_overlay_last_sig = ("", 0.0)

    def _resize_overlay_to_fit_list(self) -> None:
        """Resize the overlay so title + capped winner list (see OVERLAY_WINNER_LIST_VISIBLE_ROWS) fit."""
        if not hasattr(self, "_overlay_listbox"):
            return
        try:
            self.update_idletasks()
        except tk.TclError:
            return
        try:
            x, y = self.winfo_rootx(), self.winfo_rooty()
        except tk.TclError:
            return
        w = max(380, self.winfo_reqwidth())
        h = max(200, self.winfo_reqheight())
        try:
            self.geometry(f"{w}x{h}+{x}+{y}")
        except tk.TclError:
            pass

    @staticmethod
    def _wheel_label_text(s: str, max_chars: int = 16) -> str:
        s = (s or "").strip()
        if len(s) <= max_chars:
            return s
        return s[: max_chars - 1] + "…"

    def _wheel_gen_strip(
        self,
        rows: list[tuple[str, int, str]],
        winner: str,
    ) -> tuple[list[str], int, float]:
        n = self.WHEEL_STRIP_LEN
        win_idx = random.randint(18, min(28, n - 6))
        strip = draw_prize.build_wheel_spin_strip(
            rows,
            n,
            winner=winner,
            win_idx=win_idx,
        )
        w = max(int(self.wheel_canvas.winfo_width()), 520)
        target = self._wheel_scroll_to_center_index(float(win_idx), float(w))
        return strip, win_idx, max(0.0, target)

    def _wheel_loading_preview_setup(self, cells: int | None = None) -> None:
        """While waiting on the worker, show random winnable SKUs (potential hits) instead of placeholders."""
        self._cancel_wheel_idle_drift()
        if cells is None:
            cells = int(type(self).WHEEL_LOADING_STRIP_LEN)
        self._clear_obs_wheel_result()
        self._wheel_pulse_highlight = False
        fallback = ["···"] * cells
        win_idx = min(cells // 2, max(0, cells - 1))
        cw_px = max(int(self.wheel_canvas.winfo_width()), 520)

        def _center_strip(strip: list[str]) -> None:
            self._wheel_strip = strip
            self._wheel_win_idx = win_idx
            self._wheel_target_scroll = self._wheel_scroll_to_center_index(float(win_idx), float(cw_px))
            self._wheel_scroll = self._wheel_target_scroll

        try:
            p = self._list_path()
        except OSError:
            self._wheel_sku_to_img = {}
            _center_strip(fallback)
            self._wheel_redraw()
            return
        if not p.is_file():
            self._wheel_sku_to_img = {}
            _center_strip(fallback)
            self._wheel_redraw()
            return
        try:
            _, rows = draw_prize.load_rows(p)
        except draw_prize.PrizeDrawError:
            self._wheel_sku_to_img = {}
            _center_strip(fallback)
            self._wheel_redraw()
            return
        self._wheel_sku_to_img = self._sku_to_img_map_from_rows(rows)
        strip = draw_prize.sample_wheel_strip_labels(rows, cells)
        if not strip:
            _center_strip(fallback)
        else:
            _center_strip(strip)
            self._ensure_wheel_photos_for_strip(self._wheel_strip)
        self._wheel_redraw()

    def _wheel_draw_strip_pointer(self, c: tk.Canvas, cx: float, pad_top: float) -> None:
        """Down-pointing indicator above the strip (tip aims into the row)."""
        tip_y = pad_top + 3
        c.create_polygon(
            cx,
            tip_y,
            cx - 11,
            tip_y - 17,
            cx + 11,
            tip_y - 17,
            fill=WHEEL_STRIP_POINTER,
            outline=WHEEL_STRIP_POINTER_EDGE,
            width=1,
        )

    def _wheel_redraw(self, _idle_drift_frame: bool = False) -> None:
        if len(self._wheel_strip) > 0:
            self._wheel_banner_title = None
            self._wheel_banner_subtitle = None
            self._wheel_banner_is_error = False
        c = self.wheel_canvas
        c.delete("all")
        h = int(c.cget("height"))
        w = max(int(c.winfo_width()), 2)
        cx = w / 2.0
        scroll_vis = self._wheel_scroll_for_strip_render()
        scroll_model = float(self._wheel_scroll)
        cw = float(self.WHEEL_CELL)
        pl, pr = self._wheel_cell_xmargins()
        pad_top = float(WHEEL_STRIP_PAD_TOP)
        pad_bot = float(WHEEL_STRIP_PAD_BOTTOM)
        cell_h = max(32.0, float(h) - pad_top - pad_bot)
        landed = abs(scroll_model - self._wheel_target_scroll) < 1.5
        gap = WHEEL_STRIP_CELL_GAP

        y_track = min(float(h) - 3.0, pad_top + cell_h + 3.0)
        c.create_line(0, y_track, w, y_track, fill=WHEEL_STRIP_TRACK, width=1)

        n_idle_phys = int(type(self).WHEEL_IDLE_STRIP_LEN) * int(type(self).WHEEL_IDLE_CAROUSEL_COPIES)
        idle_carousel = len(self._wheel_strip) == n_idle_phys
        for i, raw in enumerate(self._wheel_strip):
            x0 = i * cw - scroll_vis + pl
            x1 = i * cw - scroll_vis + cw - pr
            if x1 < 0 or x0 > w:
                continue
            rx0, rx1 = x0 + gap, x1 - gap
            if rx1 <= rx0 + 1.0:
                continue
            cell_center_x = i * cw + (pl + cw - pr) / 2.0 - scroll_vis
            if idle_carousel:
                win_cell = False
            elif len(self._wheel_strip) == self.WHEEL_IDLE_STRIP_LEN:
                win_cell = landed and abs(cell_center_x - cx) < cw * 0.38
            else:
                win_cell = landed and i == self._wheel_win_idx and abs(cell_center_x - cx) < cw * 0.38
            outline_w = 1
            if win_cell:
                outline_w = 3 if self._wheel_pulse_highlight else 2
            fill_c = WHEEL_STRIP_WIN_BG if win_cell else WHEEL_STRIP_SLOT_BG
            outline_c = WHEEL_STRIP_WIN_RING if win_cell else WHEEL_STRIP_SLOT_BD
            fg_c = WHEEL_STRIP_WIN_FG if win_cell else WHEEL_STRIP_FG
            c.create_rectangle(
                rx0,
                pad_top + 2,
                rx1,
                pad_top + cell_h - 2,
                fill=fill_c,
                outline=outline_c,
                width=outline_w,
            )
            img_ref = _normalize_image_ref(self._wheel_sku_to_img.get(raw, ""))
            ph = self._wheel_image_cache.get(img_ref) if img_ref else None
            cx_cell = (rx0 + rx1) / 2.0
            cy_cell = pad_top + 2.0 + (cell_h - 4.0) / 2.0
            slot_font = self._wheel_font_bold if win_cell else self._wheel_font
            if ph is not None:
                c.create_image(cx_cell, cy_cell, image=ph)
            else:
                c.create_text(
                    cx_cell,
                    cy_cell,
                    text=self._wheel_label_text(raw),
                    fill=fg_c,
                    font=slot_font,
                    width=int(max(12.0, rx1 - rx0 - 12.0)),
                )

        self._wheel_draw_strip_pointer(c, cx, pad_top)
        self._wheel_publish_html_snapshot()
        if (
            not _idle_drift_frame
            and self._wheel_idle_drift_active()
            and self._wheel_idle_drift_after is None
        ):
            self._wheel_idle_drift_after = self.after(
                int(type(self).WHEEL_IDLE_DRIFT_MS),
                self._wheel_idle_drift_tick,
            )

    def _reset_wheel_idle(self) -> None:
        self._cancel_pulse()
        self._cancel_wheel_idle_drift()
        self._clear_obs_wheel_result()
        self._wheel_pulse_highlight = False
        units = self._winnable_units_total()
        if units is not None and units == 0:
            self._update_draw_buttons_for_supply()
            self._reconcile_depleted_wheel_if_applicable()
            return
        self._customer_empty_wheel = False
        self._paint_normal_idle_wheel_only()
        self._update_draw_buttons_for_supply()

    def _hide_super_reroll_keep_buttons(self) -> None:
        """Super-only actions stay out of the draw bar until a Super spin is waiting for Reroll/Keep."""
        if not hasattr(self, "reroll_btn"):
            return
        for w in (self.reroll_btn, self.keep_btn):
            try:
                w.pack_forget()
            except tk.TclError:
                pass

    def _show_super_reroll_keep_buttons(self) -> None:
        if not hasattr(self, "reroll_btn"):
            return
        self.reroll_btn.pack(side=tk.LEFT, padx=(0, 10), after=self.super_btn)
        self.keep_btn.pack(side=tk.LEFT, padx=(0, 10), after=self.reroll_btn)

    def _style_main_controls_idle(self) -> None:
        if not hasattr(self, "spin_btn"):
            return
        self._hide_super_reroll_keep_buttons()
        self.spin_btn.config(state=tk.NORMAL, bg=BTN_SPIN)
        self.super_btn.config(state=tk.NORMAL, bg=BTN_SUPER)
        self.skip_spot_btn.config(state=tk.NORMAL, bg=BTN_SKIP)
        self._sync_undo_spin_button()

    def _style_super_controls_active(self) -> None:
        if not hasattr(self, "spin_btn"):
            return
        self.spin_btn.config(state=tk.DISABLED, bg=BTN_DISABLED)
        self.super_btn.config(state=tk.DISABLED, bg=BTN_DISABLED)
        self._show_super_reroll_keep_buttons()
        if self._super_reroll_used:
            self.reroll_btn.config(state=tk.DISABLED, bg=BTN_DISABLED)
        else:
            self.reroll_btn.config(state=tk.NORMAL, bg=BTN_REROLL)
        self.keep_btn.config(state=tk.NORMAL, bg=BTN_KEEP)
        self.skip_spot_btn.config(state=tk.NORMAL, bg=BTN_SKIP)
        self._sync_undo_spin_button()

    def _wheel_show_error(self, message: str) -> None:
        self._cancel_wheel_idle_drift()
        self._clear_obs_wheel_result()
        self._wheel_sku_to_img = {}
        self._wheel_strip = []
        self._wheel_banner_title = str(message)
        self._wheel_banner_subtitle = ""
        self._wheel_banner_is_error = True
        self.wheel_canvas.delete("all")
        w = max(int(self.wheel_canvas.winfo_width()), 400)
        h = int(self.wheel_canvas.cget("height"))
        self.wheel_canvas.create_text(
            w / 2,
            h / 2 + 6,
            text=message,
            fill=WHEEL_ACCENT,
            font=self._title_font,
            width=w - 40,
        )
        self._wheel_draw_strip_pointer(self.wheel_canvas, w / 2.0, float(WHEEL_STRIP_PAD_TOP))
        self.wheel_status.configure(text=message, fg=WHEEL_ACCENT)
        self._style_main_controls_idle()
        self._wheel_publish_html_snapshot()

    def _cancel_pulse(self) -> None:
        if self._pulse_after is not None:
            try:
                self.after_cancel(self._pulse_after)
            except tk.TclError:
                pass
            self._pulse_after = None

    def _pulse_wheel_win(self, flashes: int = 10) -> None:
        self._cancel_pulse()
        state = [0]

        def step() -> None:
            state[0] += 1
            if state[0] > flashes:
                self._wheel_pulse_highlight = False
                self._wheel_redraw()
                self._pulse_after = None
                self._update_draw_buttons_for_supply()
                self._reconcile_depleted_wheel_if_applicable()
                return
            self._wheel_pulse_highlight = state[0] % 2 == 1
            self._wheel_redraw()
            self._pulse_after = self.after(110, step)

        step()

    def _run_wheel_spin(
        self,
        rows: list[tuple[str, int, str]],
        winner: str,
        subtitle: str,
        done: object,
        sku_to_img: dict[str, str] | None = None,
    ) -> None:
        self._clear_obs_wheel_result()
        self._wheel_banner_title = None
        self._wheel_banner_subtitle = None
        self._wheel_banner_is_error = False
        self._clear_animation()
        self._cancel_pulse()
        self._cancel_wheel_idle_drift()
        self._wheel_strip, self._wheel_win_idx, self._wheel_target_scroll = self._wheel_gen_strip(
            rows, winner
        )
        self._wheel_sku_to_img = dict(sku_to_img) if sku_to_img else {}
        self._ensure_wheel_photos_for_strip(self._wheel_strip)
        self._wheel_scroll = 0.0
        self._wheel_pulse_highlight = False
        self.wheel_status.configure(text="Wheel spinning…", fg=WHEEL_TITLE)
        self._wheel_redraw()

        start = time.monotonic()
        dur = float(type(self).WHEEL_SPIN_DURATION_SEC)
        ease_pow = float(type(self).WHEEL_SPIN_EASE_OUT_POWER)
        tick_ms = int(type(self).WHEEL_SPIN_TICK_MS)
        target = self._wheel_target_scroll

        def tick() -> None:
            elapsed = time.monotonic() - start
            t = min(1.0, elapsed / dur)
            # ease-out polynomial: lower power = less time crawling at near-zero speed before landing.
            omt = 1.0 - t
            ease = 1.0 - omt**ease_pow
            self._wheel_scroll = target * ease
            self._wheel_redraw()
            if t >= 1.0:
                self._anim_after = None
                self.wheel_status.configure(text=subtitle, fg=WHEEL_FG)
                done()
            else:
                self._anim_after = self.after(tick_ms, tick)

        self._anim_after = self.after(tick_ms, tick)

    def _list_path(self) -> Path:
        t = self.path_var.get().strip()
        if not t:
            return draw_prize.default_list_path()
        return Path(t).expanduser().resolve()

    def _maybe_refresh_idle_wheel(self) -> None:
        """Re-paint the idle preview strip when the list path or stock changes (not during spin/super)."""
        if (
            hasattr(self, "wheel_canvas")
            and not self._busy
            and self._pending_super is None
            and self._pulse_after is None
            and self._anim_after is None
        ):
            self._reset_wheel_idle()

    def _on_path_write(self, *_args: object) -> None:
        try:
            key = str(self._list_path())
        except OSError:
            self._clear_undo_spin()
            self._refresh_prizes_label()
            self._maybe_refresh_idle_wheel()
            return
        if self._session_target and key != self._session_target:
            self._hide_super_panel()
            self._invalidate_session()
        self._clear_undo_spin()
        self._refresh_prizes_label()
        self._maybe_refresh_idle_wheel()

    def _log(self, text: str) -> None:
        self.out.insert(tk.END, text)
        self.out.see(tk.END)

    @staticmethod
    def _find_latest_winner_session_path(base: Path) -> Path | None:
        if not base.is_dir():
            return None
        paths = list(base.glob("winners_*.xlsx"))
        if not paths:
            return None
        return max(paths, key=lambda p: p.stat().st_mtime)

    @staticmethod
    def _wheel_name_slug_for_filename(raw: str, max_len: int = 40) -> str:
        """ASCII slug for winners_<slug>_<timestamp>.xlsx; empty string if nothing usable remains."""
        s = (raw or "").strip()
        if not s:
            return ""
        s = re.sub(r"\s+", "_", s)
        s = re.sub(r"[^0-9A-Za-z_-]+", "", s)
        s = s.strip("_-")
        if not s:
            return ""
        if len(s) > max_len:
            s = s[:max_len].rstrip("_-.")
        return s

    @staticmethod
    def _parse_spot_number_from_label(label: str) -> int | None:
        m = re.search(r"Spot\s+(\d+)", (label or "").strip(), re.I)
        if not m:
            return None
        try:
            return int(m.group(1))
        except ValueError:
            return None

    def _winner_assignment_display_for_spot(self, spot_num: int) -> str:
        """Prize (SKU) text for this spot # from the session winner workbook, for OBS overlay."""
        path = getattr(self, "_winner_log_path", None)
        if load_workbook is None:
            return "Install openpyxl for winner log"
        if path is None:
            return "No winner session file"
        p = Path(path)
        if not p.is_file():
            return "No winner session file"
        try:
            mtime = p.stat().st_mtime
        except OSError:
            return "Winner log unreadable"
        cache = getattr(self, "_winner_spot_lookup_cache", None)
        if cache is not None:
            c_path, c_mtime, c_spot, c_text = cache
            if c_path == str(p.resolve()) and c_mtime == mtime and c_spot == spot_num:
                return c_text
        found = False
        sku_cell: str | None = None
        try:
            wb = load_workbook(p, read_only=True, data_only=True)
            try:
                ws = wb.active
                if ws is None:
                    txt = "Winner sheet empty"
                    self._winner_spot_lookup_cache = (str(p.resolve()), mtime, spot_num, txt)
                    return txt
                for row in ws.iter_rows(min_row=2, max_col=2, values_only=True):
                    a = row[0] if row else None
                    if a is None or a == "":
                        continue
                    try:
                        n = int(a)
                    except (TypeError, ValueError):
                        continue
                    if n != spot_num:
                        continue
                    found = True
                    b = row[1] if row and len(row) > 1 else None
                    if b is None or (isinstance(b, str) and not b.strip()) or b == "":
                        sku_cell = ""
                    else:
                        sku_cell = str(b).strip()
                    break
            finally:
                wb.close()
        except Exception:
            return "Winner log unreadable"
        if not found:
            txt = ""
        elif not sku_cell:
            txt = "No prize (skipped / empty)"
        else:
            txt = sku_cell
        self._winner_spot_lookup_cache = (str(p.resolve()), mtime, spot_num, txt)
        return txt

    @staticmethod
    def _next_winner_spot_from_workbook(path: Path) -> int:
        """Next spot # after the highest integer in column A (rows below header)."""
        if load_workbook is None:
            return 1
        wb = load_workbook(path, read_only=True, data_only=True)
        try:
            ws = wb.active
            if ws is None:
                return 1
            m = 0
            for row in ws.iter_rows(min_row=2, max_col=1, values_only=True):
                val = row[0] if row else None
                if val is None or val == "":
                    continue
                try:
                    n = int(val)
                except (TypeError, ValueError):
                    continue
                if n > m:
                    m = n
            return m + 1 if m else 1
        finally:
            wb.close()

    def list_skipped_spots_eligible_for_fill(self) -> list[int]:
        """Spot numbers in the active winner session that have a row with an empty prize (skipped / pending).

        Use with ``prepare_spin_for_skipped_spot`` so the next normal SPIN writes the prize into that row
        instead of appending a new spot. After a successful fill, the app returns to the usual next spot.
        """
        path = getattr(self, "_winner_log_path", None)
        if path is None or load_workbook is None:
            return []
        p = Path(path)
        if not p.is_file():
            return []
        out: list[int] = []
        seen: set[int] = set()
        try:
            wb = load_workbook(p, read_only=True, data_only=True)
            try:
                ws = wb.active
                if ws is None:
                    return []
                for row in ws.iter_rows(min_row=2, max_col=2, values_only=True):
                    a = row[0] if row else None
                    b = row[1] if row and len(row) > 1 else None
                    if a is None or str(a).strip() == "":
                        continue
                    try:
                        n = int(a)
                    except (TypeError, ValueError):
                        continue
                    if n in seen:
                        continue
                    b_s = "" if b is None else str(b).strip()
                    if b_s != "":
                        continue
                    seen.add(n)
                    out.append(n)
            finally:
                wb.close()
        except Exception:
            return []
        out.sort()
        return out

    def _winner_row_has_empty_prize(self, workbook_path: Path, spot_number: int) -> bool:
        if load_workbook is None:
            return False
        try:
            wb = load_workbook(workbook_path, read_only=True, data_only=True)
            try:
                ws = wb.active
                if ws is None:
                    return False
                for row in ws.iter_rows(min_row=2, max_col=2, values_only=True):
                    a = row[0] if row else None
                    b = row[1] if row and len(row) > 1 else None
                    if a is None or str(a).strip() == "":
                        continue
                    try:
                        n = int(a)
                    except (TypeError, ValueError):
                        continue
                    if n != spot_number:
                        continue
                    b_s = "" if b is None else str(b).strip()
                    return b_s == ""
            finally:
                wb.close()
        except Exception:
            return False
        return False

    def prepare_spin_for_skipped_spot(self, spot_number: int) -> str | None:
        """Arm the next **normal SPIN** so its prize is written into an existing skipped row for ``spot_number``.

        Validates against the active Energy Break winner session workbook (empty prize cell for that spot).
        On success returns ``None``; on failure returns a short error message. Does not run the wheel — press SPIN.

        Call ``cancel_fill_skipped_spot_mode`` to return to the usual next spot without spinning.
        """
        if self._busy or self._pending_super is not None:
            return "Finish the current spin (including Super / Reroll / Keep) first."
        if load_workbook is None:
            return "Install openpyxl to use the winner session file."
        path = getattr(self, "_winner_log_path", None)
        if path is None:
            return "No winner session file is configured."
        p = Path(path)
        if not p.is_file():
            return "Winner session file is missing on disk."
        if not self._winner_row_has_empty_prize(p, int(spot_number)):
            return (
                f"There is no row for Spot #{spot_number} with an empty prize in this session. "
                "Check the winner list or pick another spot."
            )
        self._backfill_target_spot = int(spot_number)
        self._update_spin_counter_label()
        self._wheel_publish_html_snapshot()
        nxt = self._winner_next_spin
        self._log(
            f"Fill-skip mode: next SPIN writes the prize into Spot #{spot_number}'s empty row. "
            f"After that, the wheel goes back to the normal queue (next new row would be #{nxt}).\n"
        )
        return None

    def cancel_fill_skipped_spot_mode(self, *, silent: bool = False) -> None:
        """Leave fill-skip mode without spinning; the counter returns to the usual next spot."""
        if getattr(self, "_backfill_target_spot", None) is None:
            return
        self._backfill_target_spot = None
        self._update_spin_counter_label()
        self._wheel_publish_html_snapshot()
        if not silent:
            self._log("Fill-skip mode cancelled — using normal next spot again.\n")

    def _open_fill_skipped_spot_dialog(self) -> None:
        """Pick a skipped spot from the winner session, then use normal SPIN to write the prize into that row."""
        if self._busy or self._pending_super is not None:
            messagebox.showinfo("Fill skipped spot", "Finish the current spin first.")
            return
        spots = self.list_skipped_spots_eligible_for_fill()
        if not spots:
            messagebox.showinfo(
                "Fill skipped spot",
                "No rows with an empty prize were found in the active winner session file.\n\n"
                "Skipped spots must appear in the sheet with column B empty.",
            )
            return
        top = tk.Toplevel(self)
        top.title("Fill a skipped spot")
        top.configure(bg=PRIZE_BOARD_CONTENT_BG)
        top.transient(self)
        top.grab_set()
        tk.Label(
            top,
            text="Choose a spot that was skipped (empty prize). Then press SPIN on the main controls.",
            bg=PRIZE_BOARD_CONTENT_BG,
            fg=WHEEL_FG,
            font=("Segoe UI", 10),
            wraplength=400,
            justify=tk.LEFT,
        ).pack(fill=tk.X, padx=14, pady=(14, 8))
        lb = tk.Listbox(
            top,
            height=min(12, max(4, len(spots))),
            font=("Segoe UI", 11),
            bg=INV_SLOT_FACE,
            fg=WHEEL_FG,
            selectbackground=WHEEL_POINTER,
            selectforeground="#1a1a2e",
            highlightthickness=0,
        )
        for s in spots:
            lb.insert(tk.END, f"Spot {s}  —  (empty)")
        lb.pack(fill=tk.BOTH, expand=True, padx=14, pady=4)
        btn_row = tk.Frame(top, bg=PRIZE_BOARD_CONTENT_BG)
        btn_row.pack(fill=tk.X, padx=14, pady=(4, 14))

        def arm() -> None:
            sel = lb.curselection()
            if not sel:
                messagebox.showinfo("Fill skipped spot", "Select a spot in the list first.")
                return
            idx = int(sel[0])
            spot_n = spots[idx]
            err = self.prepare_spin_for_skipped_spot(spot_n)
            if err:
                messagebox.showerror("Fill skipped spot", err)
                return
            try:
                top.grab_release()
            except tk.TclError:
                pass
            top.destroy()
            messagebox.showinfo(
                "Fill skipped spot",
                f"Spot #{spot_n} is armed. Press SPIN once; the prize will be saved to that row.\n\n"
                f"Afterwards the wheel returns to the normal next spot (#{self._winner_next_spin}).",
            )

        tk.Button(
            btn_row,
            text="  Use this spot for next SPIN  ",
            command=arm,
            font=("Segoe UI", 10, "bold"),
            bg=BTN_SPIN,
            fg="white",
            relief=tk.FLAT,
            padx=12,
            pady=8,
            cursor="hand2",
        ).pack(side=tk.LEFT, padx=(0, 8))
        tk.Button(
            btn_row,
            text="  Close  ",
            command=lambda: _close_fill_skipped_dialog(top),
            font=("Segoe UI", 10),
            bg=WHEEL_CELL_BG,
            fg=WHEEL_FG,
            relief=tk.FLAT,
            padx=12,
            pady=8,
            cursor="hand2",
        ).pack(side=tk.LEFT)

    def _update_winner_row_prize_for_spot(self, spot: int, sku: str) -> bool:
        """Set column B for the first row whose column A equals ``spot`` and B is empty."""
        path = self._winner_log_path
        if path is None or load_workbook is None:
            return False
        p = Path(path)
        wb = None
        try:
            wb = load_workbook(p)
            ws = wb.active
            if ws is None:
                return False
            max_r = ws.max_row or 1
            sku_clean = (sku or "").strip()
            if not sku_clean:
                return False
            for r in range(2, max_r + 1):
                a = ws.cell(row=r, column=1).value
                try:
                    n = int(a) if a is not None and str(a).strip() != "" else None
                except (TypeError, ValueError):
                    n = None
                if n != spot:
                    continue
                b = ws.cell(row=r, column=2).value
                b_s = "" if b is None else str(b).strip()
                if b_s != "":
                    continue
                ws.cell(row=r, column=2, value=sku_clean)
                wb.save(p)
                return True
            return False
        except OSError:
            return False
        except Exception:
            return False
        finally:
            if wb is not None:
                try:
                    wb.close()
                except Exception:
                    pass

    def _winner_clear_prize_cell_for_spot(self, path: Path, spot_number: int) -> bool:
        """Clear column B (empty prize) for the last data row whose column A equals ``spot_number``."""
        if load_workbook is None:
            return False
        wb = None
        try:
            wb = load_workbook(path)
            ws = wb.active
            if ws is None:
                return False
            target: int | None = None
            max_r = ws.max_row or 1
            for r in range(2, max_r + 1):
                a = ws.cell(row=r, column=1).value
                try:
                    n = int(a) if a is not None and str(a).strip() != "" else None
                except (TypeError, ValueError):
                    n = None
                if n == spot_number:
                    target = r
            if target is None:
                return False
            ws.cell(row=target, column=2, value="")
            wb.save(path)
            return True
        except Exception:
            return False
        finally:
            if wb is not None:
                try:
                    wb.close()
                except Exception:
                    pass

    def _create_new_winner_workbook(self, base: Path, wheel_name_slug: str = "") -> Path | None:
        """Create winners_*.xlsx with header row. If wheel_name_slug is set, file is winners_<slug>_<stamp>.xlsx."""
        if Workbook is None:
            return None
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        slug = (wheel_name_slug or "").strip()
        if slug:
            path = base / f"winners_{slug}_{stamp}.xlsx"
        else:
            path = base / f"winners_{stamp}.xlsx"
        try:
            wb = Workbook()
            ws = wb.active
            if ws is None:
                raise RuntimeError("No active worksheet")
            ws.title = "Winners"
            ws.append(["Spot #", "Prize (SKU)"])
            ws.column_dimensions["A"].width = 10
            ws.column_dimensions["B"].width = 44
            wb.save(path)
            return path
        except OSError:
            return None

    def _init_winner_session_log(self) -> None:
        """Append to the most recent winners_*.xlsx in winner_sessions/, or create one if none exist."""
        self._winner_next_spin = 1
        self._winner_log_path = None
        if Workbook is None or load_workbook is None:
            self._log(
                "Warning: openpyxl is not installed — session winner list disabled. "
                "Run: py -3 -m pip install openpyxl\n"
            )
            return
        base = script_dir() / "winner_sessions"
        try:
            base.mkdir(parents=True, exist_ok=True)
        except OSError as e:
            self._log(f"Warning: could not create winner_sessions folder: {e}\n")
            return

        latest = self._find_latest_winner_session_path(base)
        if latest is not None:
            try:
                nxt = self._next_winner_spot_from_workbook(latest)
                self._winner_log_path = latest.resolve()
                self._winner_next_spin = nxt
                self._log(
                    f"Wheel log: appending to {latest.name} (next spot #{nxt}). "
                    f"Use Start new wheel in Spin & controls when you want a new file.\n"
                )
                return
            except Exception as e:
                self._log(
                    f"Warning: could not read latest winner file ({latest.name}); creating new sheet. {e}\n"
                )

        path = self._create_new_winner_workbook(
            base, self._wheel_name_slug_for_filename(self._new_wheel_name_var.get())
        )
        if path is not None:
            self._winner_log_path = path.resolve()
            self._winner_next_spin = 1
            self._log(f"Wheel log (new file): {path.name}\n")
        else:
            self._log("Warning: could not create session winner list.\n")

    def _start_new_wheel(self) -> None:
        if Workbook is None or load_workbook is None:
            messagebox.showinfo("New wheel", "Install openpyxl to save winners to Excel.")
            return
        if getattr(self, "_backfill_target_spot", None) is not None:
            messagebox.showinfo(
                "New wheel",
                "Cancel \"Fill a skipped spot\" first (Cancel fill mode in Spin & controls).",
            )
            return
        if self._busy or self._pending_super is not None:
            messagebox.showinfo(
                "New wheel",
                "Finish or cancel the current spin (including Super / Reroll / Keep) before starting a new wheel.",
            )
            return
        if not messagebox.askyesno(
            "New wheel",
            "Create a new Excel file for this wheel?\n\n"
            "New spins will save there. Your current file is left on disk unchanged.",
        ):
            return
        base = script_dir() / "winner_sessions"
        try:
            base.mkdir(parents=True, exist_ok=True)
        except OSError as e:
            messagebox.showerror("New wheel", str(e))
            return
        slug = self._wheel_name_slug_for_filename(self._new_wheel_name_var.get())
        path = self._create_new_winner_workbook(base, slug)
        if path is None:
            messagebox.showerror("New wheel", "Could not create a new workbook.")
            return
        self._winner_log_path = path.resolve()
        self._winner_next_spin = 1
        self._update_spin_counter_label()
        self._clear_undo_spin()
        self._log(f"New wheel file: {path.name}\n")
        self._refresh_overlay_winner_session_list()

    def _append_winner_sheet_row(self, sku: str | None) -> bool:
        """Append a new row, or (fill-skip mode) update column B on an existing skipped row."""
        path = self._winner_log_path
        if path is None or load_workbook is None:
            try:
                self._refresh_overlay_winner_session_list()
            except Exception:
                pass
            return False
        bf = getattr(self, "_backfill_target_spot", None)
        if bf is not None and sku is not None and str(sku).strip() != "":
            ok = self._update_winner_row_prize_for_spot(int(bf), str(sku).strip())
            if ok and hasattr(self, "_winner_spot_lookup_cache"):
                self._winner_spot_lookup_cache = None
            try:
                self._refresh_overlay_winner_session_list()
            except Exception:
                pass
            return ok
        sku_cell = "" if sku is None else sku
        ok = False
        wb = None
        try:
            wb = load_workbook(path)
            ws = wb.active
            if ws is None:
                raise RuntimeError("No active worksheet")
            ws.append([self._winner_next_spin, sku_cell])
            wb.save(path)
            self._winner_next_spin += 1
            self._update_spin_counter_label()
            ok = True
        except OSError as e:
            self._log(f"Warning: could not append to winner list: {e}\n")
        except Exception as e:
            self._log(f"Warning: could not append to winner list ({type(e).__name__}): {e}\n")
        finally:
            if wb is not None:
                try:
                    wb.close()
                except Exception:
                    pass
            try:
                self._refresh_overlay_winner_session_list()
            except Exception:
                pass
        return ok

    def _record_spin_winner(self, sku: str) -> bool:
        return self._append_winner_sheet_row(sku)

    def _update_spin_counter_label(self) -> None:
        if not hasattr(self, "spin_counter_label"):
            return
        bf = getattr(self, "_backfill_target_spot", None)
        if bf is not None:
            self.spin_counter_label.configure(text=f"Spot {bf} — fill skipped row")
        else:
            nxt = self._winner_next_spin
            self.spin_counter_label.configure(text=f"Spot {nxt}")

    def _on_skip_spot(self) -> None:
        if self._busy:
            return
        if getattr(self, "_backfill_target_spot", None) is not None:
            messagebox.showinfo(
                "Skip spot",
                "Cancel \"Fill a skipped spot\" first (Cancel fill mode in Spin & controls).",
            )
            return
        n = self._winner_next_spin
        if not messagebox.askyesno(
            "Skip spot",
            f"A row will be added for spot #{n} with an empty prize (no SKU). "
            f"The next row will use #{n + 1}. Use this when a buyer's payment did not clear.",
        ):
            return
        if self._append_winner_sheet_row(None):
            self._log(f"Skipped spot #{n}: row added with empty SKU. Next winner list #: {self._winner_next_spin}.\n")
            self._push_undo_skip_log(n)
        else:
            self._winner_next_spin += 1
            self._update_spin_counter_label()
            self._log(
                f"Skipped spot #{n} (winner file unavailable — counter advanced only). "
                f"Next #: {self._winner_next_spin}.\n"
            )
            self._push_undo_skip_counter_only(n)
            self._refresh_overlay_winner_session_list()

    def _browse(self) -> None:
        path = filedialog.askopenfilename(
            title="Input list",
            initialdir=str(script_dir()),
            filetypes=[("Tab-separated / text", "*.txt *.tsv *.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.path_var.set(path)
            self._invalidate_session()

    def _invalidate_session(self) -> None:
        self._session = None
        self._session_target = ""
        self._wheel_image_cache.clear()
        self._inventory_image_cache.clear()

    @staticmethod
    def _sku_to_img_map_from_rows(rows: list[tuple[str, int, str]]) -> dict[str, str]:
        out: dict[str, str] = {}
        for sku, _q, img in rows:
            if sku:
                out[sku] = img or ""
        return out

    def _resolve_prize_image_path(self, ref: str) -> Path | None:
        refn = _normalize_image_ref(ref)
        if not refn or refn.lower().startswith(("http://", "https://")):
            return None
        try:
            list_base = self._list_path().parent
            app_base = script_dir()
        except OSError:
            return None
        ref_path = Path(refn)
        rel_join = _path_for_project_join(refn)
        candidates: list[Path] = []
        seen: set[str] = set()

        def add_candidate(p: Path) -> None:
            try:
                key = str(p.resolve())
            except OSError:
                key = str(p)
            if key not in seen:
                seen.add(key)
                candidates.append(p)

        if ref_path.is_absolute():
            add_candidate(ref_path.expanduser().resolve())
        add_candidate((list_base / rel_join).resolve())
        add_candidate((app_base / rel_join).resolve())
        add_candidate((Path.cwd() / rel_join).resolve())
        if not ref_path.is_absolute():
            add_candidate(ref_path.expanduser().resolve())
        for cand in candidates:
            if cand.is_file():
                return cand
        return None

    def _wheel_thumb_max_px(self) -> tuple[int, int]:
        h = int(self.wheel_canvas.cget("height"))
        cell_h = max(32, h - WHEEL_STRIP_PAD_TOP - WHEEL_STRIP_PAD_BOTTOM - 4)
        cw = float(self.WHEEL_CELL)
        pl, pr = self._wheel_cell_xmargins()
        pad = float(pl + pr) + 8.0
        return max(28, int(cw - pad)), max(28, int(cell_h - pad))

    def _pil_image_to_tk_photo(self, im, mw: int, mh: int) -> object | None:
        try:
            from PIL import Image as PIM
            from PIL import ImageTk as PIMTk
        except ImportError:
            return None
        w, h = im.size
        scale = min(mw / max(w, 1), mh / max(h, 1), 1.0)
        nw, nh = max(1, int(w * scale)), max(1, int(h * scale))
        try:
            resample = PIM.Resampling.LANCZOS
        except AttributeError:
            resample = PIM.LANCZOS  # type: ignore[attr-defined]
        rgba = im.convert("RGBA")
        resized = rgba.resize((nw, nh), resample)
        try:
            return PIMTk.PhotoImage(resized, master=self)
        except tk.TclError:
            rgb = PIM.new("RGB", resized.size, (255, 255, 255))
            rgb.paste(resized, mask=resized.split()[3])
            return PIMTk.PhotoImage(rgb, master=self)

    def _pil_image_to_wheel_photo(self, im) -> object | None:
        mw, mh = self._wheel_thumb_max_px()
        return self._pil_image_to_tk_photo(im, mw, mh)

    def _open_pil_image_from_ref(self, refn: str):
        """Return a PIL Image from URL or disk path, or None."""
        if not refn or not _pillow_runtime_available():
            return None
        try:
            from PIL import Image as PILImage
        except ImportError:
            return None
        if refn.lower().startswith(("http://", "https://")):
            try:
                req = urllib.request.Request(
                    refn,
                    headers={
                        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
                        "Accept": "image/avif,image/webp,image/apng,image/*,*/*;q=0.8",
                    },
                )
                ctx = ssl.create_default_context()
                with urllib.request.urlopen(req, timeout=20, context=ctx) as resp:
                    data = resp.read()
            except Exception:
                return None
            if len(data) < 32 or _looks_like_html_head(data):
                return None
            try:
                im = PILImage.open(BytesIO(data))
                im.load()
                return im
            except Exception:
                return None
        path = self._resolve_prize_image_path(refn)
        if path is None:
            return None
        try:
            im = PILImage.open(path)
            im.load()
            return im
        except Exception:
            return None

    def _wheel_photo_for_normalized_ref(self, refn: str) -> object | None:
        """Load (or return cached) ImageTk.PhotoImage for wheel cell; caches by normalized ref."""
        if not refn or not _pillow_runtime_available():
            return None
        if refn in self._wheel_image_cache:
            return self._wheel_image_cache[refn]
        im = self._open_pil_image_from_ref(refn)
        if im is None:
            return None
        photo = self._pil_image_to_wheel_photo(im)
        if photo is not None:
            self._wheel_image_cache[refn] = photo
        return photo

    def _inventory_photo_for_normalized_ref(self, refn: str, thumb_px: int) -> object | None:
        """Square thumb for prize-board grid; cache key includes thumb size."""
        if not refn or not _pillow_runtime_available():
            return None
        key = (refn, thumb_px)
        if key in self._inventory_image_cache:
            return self._inventory_image_cache[key]
        im = self._open_pil_image_from_ref(refn)
        if im is None:
            return None
        photo = self._pil_image_to_tk_photo(im, thumb_px, thumb_px)
        if photo is not None:
            self._inventory_image_cache[key] = photo
        return photo

    def _ensure_wheel_photos_for_strip(self, strip: list[str]) -> None:
        smap = self._wheel_sku_to_img
        done: set[str] = set()
        for sku in strip:
            if sku in ("—", "?", "···", ""):
                continue
            ref = (smap or {}).get(sku, "")
            refn = _normalize_image_ref(ref)
            if not refn or refn in done or refn in self._wheel_image_cache:
                if refn:
                    done.add(refn)
                continue
            done.add(refn)
            self._wheel_photo_for_normalized_ref(refn)

    def _ensure_session(self) -> draw_prize.FileDrawSession:
        key = str(self._list_path())
        if self._session is not None and self._session_target == key:
            return self._session
        self._session = draw_prize.open_draw_session(Path(key))
        self._session_target = key
        return self._session

    def _clear_animation(self) -> None:
        if self._anim_after is not None:
            try:
                self.after_cancel(self._anim_after)
            except tk.TclError:
                pass
            self._anim_after = None

    def _set_busy(self, busy: bool) -> None:
        self._busy = busy
        if busy:
            self._cancel_wheel_idle_drift()
        self.path_entry.state(["disabled"] if busy else ["!disabled"])
        if busy:
            if hasattr(self, "spin_btn"):
                self._btn_set_busy(self.spin_btn, True)
                self._btn_set_busy(self.super_btn, True)
                if hasattr(self, "reroll_btn") and self.reroll_btn.winfo_ismapped():
                    self._btn_set_busy(self.reroll_btn, True)
                    self._btn_set_busy(self.keep_btn, True)
                self._btn_set_busy(self.skip_spot_btn, True)
            if hasattr(self, "undo_spin_btn"):
                self._btn_set_busy(self.undo_spin_btn, True)
        elif self._pending_super is not None:
            self._style_super_controls_active()
        else:
            self._style_main_controls_idle()
        self._sync_undo_spin_button()

    def _hide_super_panel(self) -> None:
        self._pending_super = None
        self._super_reroll_used = False
        if not self._busy:
            self._style_main_controls_idle()

    def _show_super_panel(self, session: draw_prize.FileDrawSession, result: draw_prize.SpinResult) -> None:
        self._pending_super = (session, result)
        self._style_super_controls_active()

    def _worker_pick(
        self,
        on_ok: object,
        on_err: object,
    ) -> None:
        def run() -> None:
            try:
                session = self._ensure_session()
                session.refresh()
                if not any(q > 0 for _, q, _ in session.rows):
                    raise draw_prize.PrizeDrawError("No prizes left (all quantities are 0).")
                i = draw_prize.pick_sku(session.rows)
                sku, qty, img = session.rows[i]
                r = draw_prize.SpinResult(i, sku, qty, img)
                self.after(0, lambda: on_ok(session, r, session.rows))
            except draw_prize.PrizeDrawError as e:
                self.after(0, lambda: on_err(str(e)))
            except Exception as e:
                self.after(0, lambda: on_err(f"{type(e).__name__}: {e}"))

        threading.Thread(target=run, daemon=True).start()

    def _worker_commit(
        self,
        session: draw_prize.FileDrawSession,
        result: draw_prize.SpinResult,
        on_done: object,
    ) -> None:
        def run() -> None:
            try:
                if not self.dry_run_var.get():
                    session.commit(result)
                self.after(0, lambda: on_done(None))
            except draw_prize.PrizeDrawError as e:
                self.after(0, lambda: on_done(str(e)))
            except Exception as e:
                self.after(0, lambda: on_done(f"{type(e).__name__}: {e}"))

        threading.Thread(target=run, daemon=True).start()

    def _on_spin(self) -> None:
        if self._busy:
            return
        self._hide_super_panel()
        self._set_busy(True)
        self._cancel_pulse()
        self.wheel_status.configure(text="Loading picks from your list…", fg=WHEEL_TITLE)
        self._wheel_loading_preview_setup()

        def ok(
            session: draw_prize.FileDrawSession,
            r: draw_prize.SpinResult,
            rows: list[tuple[str, int, str]],
        ) -> None:
            def after_anim() -> None:
                dry = self.dry_run_var.get()
                self._log(f"Prize: {r.sku}  (was Qty {r.qty}){'  [dry run]' if dry else ''}\n")

                def commit_done(err: str | None) -> None:
                    self._set_busy(False)
                    if err:
                        messagebox.showerror("Could not save", err)
                        self._log(f"Error: {err}\n")
                        self._invalidate_session()
                        self._wheel_show_error(str(err))
                        self._refresh_prizes_label()
                    else:
                        if not dry:
                            self._log(f"Updated: {r.sku} Qty -> {max(0, r.qty - 1)}\n")
                            bf = getattr(self, "_backfill_target_spot", None)
                            if bf is not None:
                                winner_ok = self._record_spin_winner(r.sku)
                                self._push_undo_backfill(session.path.resolve(), r, int(bf), winner_ok)
                                if winner_ok:
                                    self._set_obs_wheel_result(int(bf), r.sku)
                                    self.cancel_fill_skipped_spot_mode(silent=True)
                                    self._log(
                                        f"Filled skipped Spot #{bf} in the winner session with {r.sku}.\n"
                                    )
                                else:
                                    self._log(
                                        "Prize list was saved, but the winner sheet row could not be filled "
                                        "(try closing Excel). Fill-skip mode is still armed for another try.\n"
                                    )
                            else:
                                spot_before = self._winner_next_spin
                                self._set_obs_wheel_result(spot_before, r.sku)
                                winner_ok = self._record_spin_winner(r.sku)
                                self._push_undo_spin(session.path.resolve(), r, spot_before, winner_ok)
                            self._refresh_prizes_label()
                            last = ""
                            uq = self._winnable_units_total()
                            if uq is not None and uq == 0:
                                last = "  ·  That was the last prize — list is empty."
                            self.wheel_status.configure(
                                text=f"Winner: {r.sku}",
                                fg=WHEEL_POINTER,
                            )
                            self._pulse_wheel_win()
                        else:
                            self.wheel_status.configure(
                                text=f"{r.sku}  ·  dry run (file unchanged), was Qty {r.qty}",
                                fg=WHEEL_TITLE,
                            )

                if dry:
                    self._log("Dry run: file not updated.\n")
                    commit_done(None)
                else:
                    self._worker_commit(session, r, commit_done)

            self._run_wheel_spin(
                rows,
                r.sku,
                f"Landed on: {r.sku}  ·  Qty {r.qty}" + ("  ·  DRY RUN" if self.dry_run_var.get() else ""),
                after_anim,
                self._sku_to_img_map_from_rows(session.rows),
            )

        def err(msg: str) -> None:
            self._set_busy(False)
            self._invalidate_session()
            self._refresh_prizes_label()
            units = self._winnable_units_total()
            if units is not None and units == 0:
                messagebox.showinfo(
                    "Prizes finished",
                    "Every prize has been given out. Add quantities to your list file to draw again.",
                )
            else:
                self._wheel_show_error(str(msg))
                messagebox.showerror("Spin", msg)
            self._log(f"Error: {msg}\n")

        self._worker_pick(ok, err)

    def _on_super_spin(self) -> None:
        if self._busy:
            return
        if getattr(self, "_backfill_target_spot", None) is not None:
            messagebox.showinfo(
                "Super spin",
                "Fill-skip mode uses a normal SPIN only. Cancel fill mode first, or complete that spin.",
            )
            return
        self._hide_super_panel()
        self._set_busy(True)
        self._cancel_pulse()
        self.wheel_status.configure(text="Super spin — loading picks…", fg=WHEEL_TITLE)
        self._wheel_loading_preview_setup()

        def ok(
            session: draw_prize.FileDrawSession,
            r: draw_prize.SpinResult,
            rows: list[tuple[str, int, str]],
        ) -> None:
            def after_anim() -> None:
                self._show_super_panel(session, r)
                self._set_busy(False)
                self._log(f"Super spin result: {r.sku}  (was Qty {r.qty})\n")
                self.wheel_status.configure(
                    text=f"Keep or Reroll!",
                    fg=WHEEL_FG,
                )

            self._run_wheel_spin(
                rows,
                r.sku,
                f"Showing: {r.sku}  ·  Qty {r.qty}",
                after_anim,
                self._sku_to_img_map_from_rows(session.rows),
            )

        def err(msg: str) -> None:
            self._set_busy(False)
            self._invalidate_session()
            self._refresh_prizes_label()
            units = self._winnable_units_total()
            if units is not None and units == 0:
                messagebox.showinfo(
                    "Prizes finished",
                    "Every prize has been given out. Add quantities to your list file to draw again.",
                )
            else:
                self._wheel_show_error(str(msg))
                messagebox.showerror("Super spin", msg)
            self._log(f"Error: {msg}\n")

        self._worker_pick(ok, err)

    def _on_reroll(self) -> None:
        if self._busy or self._pending_super is None or self._super_reroll_used:
            return
        self._set_busy(True)
        self._cancel_pulse()
        self.wheel_status.configure(text="Rerolling — loading picks…", fg=WHEEL_TITLE)
        self._wheel_loading_preview_setup()

        def ok(
            s: draw_prize.FileDrawSession,
            r: draw_prize.SpinResult,
            rows: list[tuple[str, int, str]],
        ) -> None:
            def after_anim() -> None:
                self._super_reroll_used = True
                self._show_super_panel(s, r)
                self._set_busy(False)
                self._log(f"Reroll: {r.sku}  (was Qty {r.qty})\n")
                self.wheel_status.configure(
                    text=f"No more rerolls — KEEP to save −1. Pointer: {r.sku}  ·  qty {r.qty}",
                    fg=WHEEL_FG,
                )

            self._run_wheel_spin(
                rows,
                r.sku,
                f"Reroll: {r.sku}  ·  Qty {r.qty}",
                after_anim,
                self._sku_to_img_map_from_rows(s.rows),
            )

        def err(msg: str) -> None:
            self._set_busy(False)
            self._invalidate_session()
            self._reset_wheel_idle()
            self._refresh_prizes_label()
            units = self._winnable_units_total()
            if units is not None and units == 0:
                messagebox.showinfo(
                    "Prizes finished",
                    "Every prize has been given out. Add quantities to your list file to draw again.",
                )
            else:
                self._wheel_show_error(str(msg))
                messagebox.showerror("Reroll", msg)
            self._log(f"Error: {msg}\n")

        self._worker_pick(ok, err)

    def _on_keep(self) -> None:
        if self._busy or self._pending_super is None:
            return
        session, result = self._pending_super
        self._set_busy(True)
        dry = self.dry_run_var.get()
        self._log(f"Keep: {result.sku}  (was Qty {result.qty}){'  [dry run]' if dry else ''}\n")

        def commit_done(err: str | None) -> None:
            self._hide_super_panel()
            self._set_busy(False)
            if err:
                messagebox.showerror("Could not save", err)
                self._log(f"Error: {err}\n")
                self._invalidate_session()
                self._wheel_show_error(str(err))
                self._refresh_prizes_label()
            else:
                if not dry:
                    self._log(f"Updated: {result.sku} Qty -> {max(0, result.qty - 1)}\n")
                    spot_before = self._winner_next_spin
                    self._set_obs_wheel_result(spot_before, result.sku)
                    winner_ok = self._record_spin_winner(result.sku)
                    self._push_undo_spin(session.path.resolve(), result, spot_before, winner_ok)
                    self._refresh_prizes_label()
                    last = ""
                    uq = self._winnable_units_total()
                    if uq is not None and uq == 0:
                        last = "  ·  That was the last prize — list is empty."
                    self.wheel_status.configure(
                        text=f"Kept: {result.sku}  ·  was {result.qty}  →  now {max(0, result.qty - 1)}{last}",
                        fg=WHEEL_POINTER,
                    )
                    self._pulse_wheel_win()
                else:
                    self._log("Dry run: file not updated.\n")
                    self.wheel_status.configure(
                        text=f"{result.sku}  ·  dry run — qty still {result.qty} on file",
                        fg=WHEEL_TITLE,
                    )

        if dry:
            commit_done(None)
        else:
            self._worker_commit(session, result, commit_done)


def _close_fill_skipped_dialog(top: tk.Toplevel) -> None:
    try:
        top.grab_release()
    except tk.TclError:
        pass
    try:
        top.destroy()
    except tk.TclError:
        pass


def main() -> None:
    app = DrawPrizeApp()
    app.mainloop()


if __name__ == "__main__":
    main()
