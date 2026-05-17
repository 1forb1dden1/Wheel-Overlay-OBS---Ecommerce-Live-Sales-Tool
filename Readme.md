# Energy Break — Prize wheel system

Energy Break runs live prize-wheel draws for shows: weighted spins, winner logging to Excel, a compact operator overlay, a **Spin & controls** panel, an optional **Prize board** window, and an **HTML wheel** for OBS (browser source on `http://127.0.0.1:8765/`).

Only **one app instance** should run at a time so the HTML/OBS feed and port stay in sync.

## Requirements

- Windows (primary; batch launcher included)
- Python 3.10+ with dependencies: `pip install -r requirements.txt`
- **openpyxl** — required for winner session `.xlsx` files

**Run from source:** double-click `Windows Energy Break Execute.bat` or:

```bat
py -3 draw_prize_ui.py
```

---

## What’s on screen

| Area | Purpose |
|------|---------|
| **Main overlay** | Drag bar, **Show setup**, winner list preview, **Choose winner spreadsheet…**, **Prize board**, **Spin & controls…** |
| **Show setup** (expand on overlay) | Prize list path, images folder, **wheel preset**, dry run, event log, **Reset saved session…** |
| **Spin & controls** | Live strip wheel, **SPIN** / **SUPER SPIN**, skip/undo/fill-skip, HTML wheel link, wheel preset (duplicate of setup) |
| **Prize board** | Grid of remaining inventory (separate window) |
| **HTML wheel (OBS)** | Same strip and labels as Spin & controls, for Browser Source |

The moving strip you show on stream is driven from **Spin & controls** and mirrored at **`http://127.0.0.1:8765/`**.

---

## Files and folders (next to the app)

| Item | Role |
|------|------|
| `energy_break_state.json` | Saved session: last preset, list path, images path, active winner spreadsheet (auto-written while you work) |
| `winner_sessions/` | Winner log workbooks (one active file per run; see below) |
| `<wheelId>.xlsx` | Prize list for a **wheel preset** (e.g. `wheel58912.xlsx`) |
| `<wheelId>/` | Images for that preset (e.g. `wheel58912/GrassEnergy.jpg`) |
| `Input List.xlsx` | Optional generic list if you set paths manually (not auto-filled on reset) |

Prize lists are tab-separated text inside `.xlsx` / `.txt`: columns **SKU**, **Qty**, optional **img** (file name or path under the images folder).

---

## Before you can spin

All of the following must be ready (SPIN stays disabled until they are):

1. **Prize list** — set in Show setup or by **Apply preset**
2. **Images folder** — set in Show setup or by **Apply preset**
3. **Winner spreadsheet** — created/reused when you change preset, or picked with **Choose winner spreadsheet…** on the main overlay

Install **openpyxl** if the app says winner Excel files are unavailable.

---

## Show setup — paths and options

1. On the overlay, turn on **Show setup (choose files & options)**.
2. **Step 1 — Prize list:** browse or type path to your list file.
3. **Step 2 — Images folder:** folder containing prize artwork referenced in the list.
4. Scroll for **Wheel preset** and session tools (see below).
5. Optional: **Dry run (never write file)** — practice spins; does not write the prize list or winner sheet.
6. Optional: **Show event log** — technical messages at the bottom of setup.

There is **no “Use defaults”** button. After **Reset saved session…**, list and images paths are **cleared**; set them again with a preset or the Choose… buttons.

### Reset saved session…

- Deletes `energy_break_state.json`
- Clears wheel preset, prize list path, images path, and active winner spreadsheet in memory
- Does **not** delete preset files, `winner_sessions/` workbooks, or prize lists on disk
- Next launch behaves like a fresh session until you configure paths again

---

## Wheel presets

A preset ties a wheel id to a prize list and images folder next to the app:

- `wheel58912.xlsx` + `wheel58912/` (images)

In **Show setup** (or **Spin & controls**):

| Control | Action |
|---------|--------|
| Preset id field | Type an id (e.g. `wheel58912`) |
| **Apply preset** | Load that list + images folder; handle winner spreadsheet (below) |
| **Match from file…** | Pick any prize list file; id is taken from the file name |

Paths and preset id are saved in `energy_break_state.json`.

### Winner spreadsheet when switching presets

When you **change to a different preset**, the app:

1. Looks in `winner_sessions/` for an existing workbook for that wheel (name starts with `{wheelId}_winners` or legacy `winners_{wheelId}_`)
2. **Reuses** the newest match if found (e.g. switching `wheel58191` → `wheel123` → `wheel58191` opens the same `wheel123` file again)
3. **Creates** a new file only if none exists for that wheel

**New file name format** (Texas Central time):

`{wheelId}_winners_{Month}_{day}_{hour}-{minute}{AM|PM}.xlsx`

Example: `wheel58912_winners_May_16_3-45PM.xlsx`

### Choose winner spreadsheet… (main overlay)

Above the **Prize Wheel** winner list on the overlay:

- Pick any existing `.xlsx` in `winner_sessions/` (or elsewhere) to receive spins
- Does not change the wheel preset; useful for continuing an older log manually

---

## Quick reference — wheel buttons (Spin & controls)

| Button | What it does |
|--------|----------------|
| **SPIN** | Spin once; prize **committed** when done (unless dry run). Use for fill-skip flows. |
| **SUPER SPIN** | Spin; then **KEEP** or one **REROLL** → then **KEEP**. |
| **REROLL** | After SUPER only — spin again once. |
| **KEEP** | After SUPER only — confirm and save landed prize. |
| **Skip spot** | No spin; logs an empty prize row and advances the spot sequence. |
| **Undo spin** | Undo last eligible spin/skip-like action. |
| **Fill a skipped spot…** | Arms a skipped row; **next SPIN** fills it. |
| **Cancel fill mode** | Leave fill mode without spinning. |
| **Open HTML wheel…** | Browser copy of the OBS strip + labels. |

---

## Step-by-step workflows

### Normal draw — SPIN

**Use case:** Regular spin after payment.

1. Confirm setup (list, images, winner spreadsheet) and spot counter on the overlay list.
2. Click **SPIN** in Spin & controls.
3. Landed prize is written to the active winner spreadsheet and the prize list quantity decreases (unless **Dry run**).
4. **Undo spin** may be available afterward if you need to reverse a mistake or failed payment.

### Super draw — SUPER SPIN

**Use case:** Hype spin with one optional reroll.

1. Click **SUPER SPIN**.
2. When the strip stops, choose **KEEP** or **REROLL**.
3. **KEEP** — save current result (respects dry run).
4. **REROLL** — one more spin; then **KEEP** on the final result.

### Skip spot

**Use case:** Payment failed for this spot; leave it empty for now.

1. Click **Skip spot** and confirm if asked.
2. Row is added to the winner sheet with an empty prize; spot advances.

### Undo spin

**Use case:** Mistaken spin or payment failed after the fact.

1. Available when the wheel is idle and an undo exists.
2. Click **Undo spin** and confirm if asked.
3. Restores prize list quantity and winner row where possible.

### Fill a skipped spot

**Use case:** Customer pays later for a spot you skipped earlier.

1. Click **Fill a skipped spot…** and pick the spot.
2. Next **SPIN** writes the prize into that row (not a new spot).
3. **Cancel fill mode** to abort without spinning.

**Note:** Super spin reroll flow is not used for fill-skip; use normal **SPIN**. To redo a filled skip, undo then spin again.

### OBS — HTML wheel

1. Start Energy Break (single instance).
2. In Spin & controls → **Open HTML wheel…** to preview.
3. In OBS, add a **Browser** source: URL **`http://127.0.0.1:8765/`**
4. Key out the dark green chroma behind the wheel if needed (see event log chroma hint).
5. If the wheel looks frozen or wrong, ensure only one app is running and refresh the browser source after restart.

---

## Main overlay — Prize Wheel list

Below **Choose winner spreadsheet…**, the **Prize Wheel** panel shows rows from the active (or latest) winner workbook so operators can see spots and prizes without opening Excel.

---

## Troubleshooting

| Issue | What to try |
|-------|-------------|
| SPIN disabled | Set prize list + images (preset or setup), attach a winner spreadsheet |
| “Another instance” / wrong OBS wheel | Close extra copies; only one process should own port **8765** |
| Could not save list / winner file | Close the file in Excel |
| HTML wheel out of date | Restart app, refresh OBS browser source |
| After reset, nothing configured | Apply a wheel preset or set paths manually in Show setup |

---

## Legacy note on images

Older readme screenshots may show controls in different places. As of the current build:

- **Wheel preset** lives under **Show setup** (and is duplicated at the top of **Spin & controls**).
- **Choose winner spreadsheet…** is on the **main overlay**, above the Prize Wheel list.
- Session reset clears paths; it does not restore Input List / Images automatically.
