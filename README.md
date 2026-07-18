# Traverse Form Generator (Vercel-ready)

Recreates the "Form 10" traverse-computation workbook (TOC's, Job History,
INDEX, DTM, TRVS, Field notes, ABSTRACT, AREA) as a real `.xlsx` with live
formulas, for **any number of corner points**.

This layout is structured to deploy directly on Vercel: one Python
serverless function (`api/index.py`) plus a static frontend (`public/index.html`),
served from the same domain so there's no CORS to configure.

```
vercel.json          <- routes /api/* to the function, everything else to /public
api/
  index.py            <- Flask app + the whole generation engine, in one file
  requirements.txt
public/
  index.html           <- the data-entry form + live parcel preview
sample_request.json    <- worked example (real numbers, verified against the original file)
Traverse_Form_Extension_Guide.docx
```

## Deploy to Vercel

**Option A -- Vercel CLI (fastest):**
```bash
npm i -g vercel      # if you don't have it already
cd this-project-folder
vercel                # first deploy, follow the prompts
vercel --prod          # promote to production
```

**Option B -- Git-connected project:**
1. Push this folder to a GitHub/GitLab/Bitbucket repo.
2. In the Vercel dashboard: **New Project → Import** that repo.
3. Vercel will read `vercel.json` and build both the function and the
   static site automatically -- no framework preset or build command needed.
4. Deploy. Your form is at `https://<your-project>.vercel.app/`, and the
   API is at `https://<your-project>.vercel.app/api/generate`.

That's it -- `public/index.html` already calls `/api/generate` as a
relative path, so it works immediately at whatever domain Vercel gives you
(or a custom domain you attach later). No environment variables needed.

## Why one file for the whole backend

Vercel's Python runtime is most reliable when a function doesn't depend on
importing sibling modules -- so `api/index.py` contains the entire
generation engine *and* the Flask routes in a single file, rather than
importing a separate `traverse_generator.py`. It's the same code as before,
just consolidated for deployment safety.

## Local development

**With the Vercel CLI (closest to production):**
```bash
vercel dev
```
This runs the Python function and serves `public/` together, exactly like
production, at `http://localhost:3000`.

**Or run the Flask app directly (no Vercel CLI needed):**
```bash
cd api
pip install -r requirements.txt
python index.py                     # serves http://localhost:5000
```
Then open `public/index.html` in a browser and change `API_URL` near the
top of its `<script>` to `"http://localhost:5000"` (it defaults to a
relative path, which only works when frontend and API share an origin).

## The data model

One JSON object in, one `.xlsx` out. See `sample_request.json` for a full
real-world example. Quick shape:

```jsonc
{
  "opening": {
    "backsight1": {"name": "Cm1", "n": 140792.76, "e": 364922.15},
    "backsight2": {"name": "Cm2", "n": 140778.85, "e": 363970.53},
    "start":      {"name": "Cm3", "n": 141456.32, "e": 363862.81, "is_plot_corner": true}
  },
  "closing": {
    "end":        {"name": "cm75", "n": 143571.19, "e": 362131.82, "is_plot_corner": true},
    "foresight1": {"name": "cm74", "n": 143569.35, "e": 361634.81},
    "foresight2": {"name": "cm79", "n": 143206.69, "e": 361650.79}
  },
  "measured_total_distance": 18782.419,
  "corners": [
    {"name": "cm51", "n": 143587.19, "e": 362847.18, "is_plot_corner": true},
    ...
  ],
  "job_meta": {"job_number": "", "client": "", "surveyor": "", "date": "", "description": ""},
  "index_notes": {"instrument": "Total Station", "district": "", "elevation": 1220, "msl_correction": -0.1911, "scale_factor": -0.3737}
}
```

- `opening` / `closing` — the two triangles of known control stations that
  give the traverse its starting and closing bearings (fixed at 3 + 3
  stations regardless of corner-point count — see the extension guide for
  why).
- `corners` — one entry per traverse station, in walking order, any length
  ≥ 1. Just name + coordinates — leg distances are computed from the
  coordinates, not entered.
- `is_plot_corner` — tag any station (including `start` and `end`) that's
  an actual parcel boundary corner. Only tagged stations feed ABSTRACT and
  AREA.
- **Angle corrections aren't an input at all.** The opening pair is fixed
  at 0; the closing pair auto-computes so their average lands exactly on
  `(2 × number of stations)` — the target misclosure observed consistently
  across real jobs — split by a minimal ±2 offset. The resulting
  per-station angular misclosure is always exactly `2`.
- `measured_total_distance` — the one figure that genuinely can't be
  derived from this traverse's own coordinates (verified against the
  original file: it's roughly 2.4× this polygon's own perimeter, i.e. the
  total distance for the wider survey job).
- `job_meta` / `index_notes` — optional; feed the Job History and INDEX
  sheets.

## What each sheet does

| Sheet | Built from | Notes |
|---|---|---|
| **TOC's** | static | Contents list, matching the original's item list. |
| **Job History** | `job_meta` | Job number / client / surveyor / date / description. |
| **INDEX** | `index_notes` | Elevation, mean-sea-level correction, scale factor → combined correction → multiplication factor. |
| **DTM** | `opening`, `closing` | Fixed size (3+3 control stations). |
| **TRVS** | `corners` | Resizes to `len(corners)`. Bearings, distances, coordinate corrections. |
| **Field notes** | `TRVS`, `INDEX` | Reverse-derives plausible backsight/foresight field readings from the adjusted TRVS bearings, generalized for any N. See the note in the code's module docstring for the couple of places this deliberately cleans up small inconsistencies in the original example. |
| **ABSTRACT** | tagged plot corners | Station / Northing / Easting, pulled live from TRVS via formula. |
| **AREA** | tagged plot corners | Shoelace-formula area & perimeter, polygon closed automatically. |

## District auto-fill (INDEX sheet corrections)

The "Distance correction" section has a **District / Town** dropdown that
auto-fills elevation, MSL correction, and scale factor from a standard
reference table (UTM values only — TM values in the source were ignored
per instruction). The table (`DISTRICT_DATA` in `public/index.html`) was
transcribed from a scanned standard-corrections table (124 towns); a
handful of entries too degraded to read confidently were left out rather
than guessed. To add or fix an entry, find `const DISTRICT_DATA = {...}`
near the top of the `<script>` block. Worth spot-checking a few entries
against the original document before relying on this for a real job.

## Verified against two independent real files

Checked against two real jobs' actual data (`Form_10.xls`, an 11-leg
traverse, and `Form_6.xls`, a 7-leg traverse from a different job entirely
— different district, different station names). Every TRVS column
(bearings, distances, corrections, angular/linear misclosure) and every
Field notes formula matches both originals exactly, row for row, across
N = 1, 3, 6, 10, and 25 corner points with zero formula errors.

Cross-checking against the second file caught two real bugs the first pass
missed (it had only diffed the coordinate-correction columns, not the
bearing-display ones): the per-corner angular correction was referencing
the wrong row (wrong displayed bearing, though corrected coordinates were
unaffected), and Field notes' index-correction/angle-check chaining had a
few edge cases wrong. Both are fixed and re-verified against both files.

Formatting was checked too: the intermediate decimal-degree bearing
columns (TRVS F/J/K, DTM F, Field notes E) are hidden in the original,
matching now. One thing deliberately *not* replicated: both original
files' AREA sheet has a visible block of hardcoded coordinates for an
unrelated plot — a stale leftover, not the intended design — so the
generator keeps its own correct, job-specific, live-formula AREA sheet
instead of copying that.

## Third correction pass: DMS rounding-carry bug

A real bug: every degrees/minutes/seconds conversion (TRVS, DTM, Field
notes) split degrees/minutes with `TRUNC` but rounded seconds separately
with `ROUND`, so 59.6″ could round up to an invalid `60″` instead of
carrying into the next minute (`283°37'60"` instead of `283°38'00"`).
Every site now converts through total rounded seconds and derives D/M/S
via integer division and modulo, so a carry always propagates correctly.
Re-verified against the original file — zero value differences outside
this edge case. The same fix applies to Field notes' randomized
double-reading residual, which previously could show an invalid
`180°00'-3"` instead of properly borrowing a minute (`179°59'57"`).

TRVS's final misclosure row alignment was also revised: `E`/`L`/`Q`/`S`
right-aligned, `G`/`R`/`T` left-aligned.

## Second formatting correction pass

TOC's rows 6–15 now merge `B:H` per row with left-aligned dot-leader text.
INDEX hides rows 1–26 and 29–77 (only the blank spacer at 27–28 and the
content block at 78–100 stay visible) with left-aligned paragraph text.
DTM hides column `G` too (same category as `F`), fits column widths, and
each control station now has an optional **Remarks** field. Field notes'
double-reading residual is randomized within a realistic ±4″/-3″ band
instead of a flat 0; the "AT" marker is right-aligned with a lighter
per-cell border instead of a full-row one; and the table grid now covers
every row, including blank spacers. TRVS's `P`/`R` corrections weight off
column `O` again (safe — `O` mirrors `M`, not `L`, so no circular
reference); the misclosure sentence is now fully dynamic instead of a
static label; and alignment/column widths are fitted throughout.

## Formatting fidelity (latest correction pass)

A follow-up review against both sample files caught a further round of
things: sheet tab order now matches the originals exactly (TOC's, Job
History, INDEX, ABSTRACT, DTM, Field notes, TRVS, AREA); merged cells now
match on DTM/TRVS/Field notes/ABSTRACT/AREA; `O` (leg distance) now equals
`L` (final corrected distance) rather than `M` (raw distance) — done
without creating a circular reference by moving the `P`/`R` weighting to
key off `M` directly; angular misclosure per station is now `(number of
stations × 2)`; additional hidden columns (TRVS `M`/`O`, AREA `E`/`F`,
Field notes `F`); ABSTRACT's blank spacer columns are gone (Stn/Northings/
Eastings are now contiguous and centered); and colors/number
formats/alignment now match the samples' scheme throughout (DTM-linked
cells in blue, the angular correction in a different blue, leg distance in
brown, everything else computed in black; 3dp for coordinates/distances,
whole numbers for degree components, centered throughout).

## Extending this project

- **Auth / persistence** — stateless right now: one request in, one file
  out, nothing saved.
- **Validation** — add schema validation (e.g. `pydantic`) before exposing
  this beyond local/personal use.
- **Recalculated cache values** — openpyxl writes formulas without cached
  results. Excel/Google Sheets/LibreOffice recalculate on open, so this is
  invisible in normal use; a viewer that doesn't recalculate on open would
  show blanks until it does.
