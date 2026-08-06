# UTKBudgetExtractor

Tools for turning UTK proposal-budget spreadsheets into LaTeX-ready numbers and
a Department of Energy (DOE) budget justification.

Point the tool at a folder of UTK "Proposal Budget" spreadsheets
(`.xlsx`, the central `TCE 001 (Rev 03.09.26)` layout — see
[`examples/Proposal_Budget_Basic.xlsx`](examples/Proposal_Budget_Basic.xlsx))
and it produces, for **every** input file:

1. **A LaTeX definitions file** named after the input (`PI_Smith.xlsx` →
   `PI_Smith.tex`) containing every labelled number as a `\newcommand`, ready to
   `\input` into a proposal or justification.
2. **A DOE budget justification for that input** (`PI_Smith_justification.tex`)
   — only that budget, with travel called out in detail. It `\input`s its own
   defs file (item 1), so it both compiles on its own and drops into a larger
   proposal (see [Including justifications](#including-justifications)).
3. **A merged Excel workbook** (`merged.xlsx`) **in the same format as the
   inputs**, combining all of the budgets (see [Merging](#how-the-merge-works)).
4. **Merged LaTeX definitions** (`merged.tex`) for the combined budget.
5. **A separate combined-total justification** (`justification.tex`) for the
   summed budget.
6. **A driver** (`all_justifications.tex`) that pulls the whole request into one
   PDF, in reading order: both summary tables, then the combined/all-programs
   justification, then the individual ones.
7. **Two faculty summary tables**, ready to `\input` into a larger document:
   `faculty_summary.tex` (one row per faculty with their direct, indirect, and
   total DOE ask) and `faculty_summary_by_year.tex` (one row per faculty, **one
   column per year**, so a program manager can see the ask per year per PI at a
   glance). With program sub-folders (a multi-thrust proposal, e.g. `Energy
   Frontier/`, `Intensity Frontier/`, `Theory Frontier/`), both group the PIs
   under each sub-folder's name with a per-thrust subtotal.

All generated files land in **one flat output directory** — no sub-folders — so
they are easy to `\input` from a single place.

## Install

```bash
pip install -r requirements.txt   # just openpyxl
```

## Usage

```bash
python -m utkbudget /path/to/budgets            # writes to ./output
python -m utkbudget /path/to/budgets -o out_dir # custom output dir
# equivalently:
python UTKBudgetExtractor.py /path/to/budgets -o out_dir
```

### Flat folder

```
budgets/
├── PI_Smith.xlsx
└── PI_Jones.xlsx
```

produces

```
output/
├── PI_Smith.tex                 # \newcommand defs for Smith
├── PI_Smith_justification.tex   # Smith's justification (\input's PI_Smith.tex)
├── PI_Jones.tex                 # \newcommand defs for Jones
├── PI_Jones_justification.tex   # Jones's justification (\input's PI_Jones.tex)
├── merged.xlsx                  # Smith + Jones, summed
├── merged.tex                   # \newcommand defs for the sum
├── justification.tex            # combined-total justification (the sum)
├── all_justifications.tex       # summaries + combined + each justification
├── faculty_summary.tex          # one-row-per-faculty request table
└── faculty_summary_by_year.tex  # one row per faculty, one column per year
```

### Folder with program sub-folders

Hand it a folder whose sub-folders separate the DOE programs and each program is
processed independently **and** a fully merged version is produced across all of
them:

```
budgets/
├── Program 1/
│   ├── PI_Smith.xlsx
│   └── PI_Jones.xlsx
├── Program 2/
│   └── PI_Lee.xlsx
└── Program 3/
    └── PI_Richers.xlsx
```

produces — everything in **one flat directory**, no sub-folders:

```
output/
├── PI_Smith.tex                    # defs, one per input file
├── PI_Smith_justification.tex      # Smith only (\input's PI_Smith.tex)
├── PI_Jones.tex
├── PI_Jones_justification.tex
├── PI_Lee.tex                      # (from Program 2)
├── PI_Lee_justification.tex
├── PI_Richers.tex                  # (from Program 3)
├── PI_Richers_justification.tex
├── ProgramOne_merged.xlsx          # per-program merged workbook + defs
├── ProgramOne_merged.tex
├── ProgramOne_justification.tex    # Program 1 combined total
├── ProgramTwo_merged.xlsx  ...     # likewise for the other programs
├── merged.xlsx                     # fully merged across every program
├── merged.tex                      # defs for the grand total
├── justification.tex               # per-program + fully merged grand total
├── all_justifications.tex          # summaries + combined + each justification
├── faculty_summary.tex             # PIs grouped by program, per-program subtotals
└── faculty_summary_by_year.tex     # PIs by program x year, per-program subtotals
```

If two programs hold a like-named spreadsheet, the second one's outputs get a
`_2` suffix (`PI_Smith_2.tex`) rather than overwriting the first.

## Including the generated LaTeX

**Every** generated `.tex` works both ways — compile it on its own, or `\input`
it into a larger proposal. That covers the per-PI justifications, the two summary
tables, and `all_justifications.tex` (which pulls in the summary tables and
every justification at once). Each wraps its preamble in an `\ifdefined` guard:

```latex
% in your proposal's preamble:
\def\budgetjustificationincluded{}    % skip the generated files' own preambles
\def\budgetjustificationpath{output/} % where the generated files live
...
% in the body — any of these:
\input{output/faculty_summary}
\input{output/faculty_summary_by_year}
\input{output/PI_Smith_justification}
\input{output/all_justifications}     % summaries + every justification
```

`\budgetjustificationpath` is needed because TeX resolves relative `\input`
paths against the *main* document's directory, not the included file's — so
without it the nested `\input`s (a justification pulling in its defs file) would
not be found. Set it once and every nested `\input` follows. If you compile from
inside `output/`, leave it unset.

To build any of them standalone:

```bash
latexmk -pdf output/all_justifications.tex      # or: pdflatex it twice
```

## Using the output in a proposal

```latex
\input{output/merged.tex}          % defines \CombinedGrandTotal, etc.
...
We request a total of \$\CombinedGrandTotal{} from the Department of Energy,
including \$\CombinedTravelTotal{} for travel.
```

Macro names are `\<Prefix><Field>`, where `<Prefix>` is derived from the input
filename (per-file defs), the program name (per-program merge), or `Combined`
(the grand merge). Useful fields include `GrandTotal`, `DirectTotal`,
`IndirectTotal`, `WagesTotal`, `FringeTotal`, `TravelTotal`,
`DomesticTravelTotal`, `ForeignTravelTotal`, `Travel`/`DomesticTravel` per-year
(`...YearOne` … `...YearFive`), `EquipmentTotal`, `ParticipantSupportTotal`,
`TuitionSubtotalTotal`, `OverheadRateYearOne`, and `FandARateType`.

## How the merge works

The merged workbook is a real copy of the budget spreadsheet (the first input is
used as the structural template, so every sheet, style, and formula is
preserved). The merge is **formula-safe: it never overwrites a cell that holds a
formula — only genuine user-input cells are copied** — and every subtotal,
total, salary, fringe, tuition, F&A, and roll-up formula recomputes from those
inputs when the workbook is opened in Excel.

Sections are combined in the way that best fits each one, and for every person
the *inputs* are copied (name, raise flag, UT/JFO, base salary, appointment,
person-months for all five periods, tenure flags) — the salary, fringe, and
totals are formulas that recompute. Fringe rates are institutional constants
that ship with the template and are never touched.

- **Named people are concatenated** (one row each): senior personnel, other
  professionals, admin/clerical, other personnel, and equipment.
- **Post-docs and GRAs are consolidated by base salary.** Lines that share a
  base salary become one line whose per-period months are the sum of
  `headcount × months`. Because the salary formula is linear in
  `headcount × months`, this reproduces the cost exactly — and it lets you merge
  **more than four GRAs** (or more than three post-docs) as long as they share a
  base. Distinct base salaries stay on separate lines.
- **Undergraduate researchers collapse to a single line** (months summed; if the
  bases differ the line is normalised to a $1 base with `Σ(base × months)` so
  the cost is still exact).
- **Travel is consolidated to one domestic and one foreign summary row per
  period**, reproducing each period's subtotal exactly (the row carries the
  per-category totals with days = travelers = 1).
- **Supplies are grouped by description and summed**; subcontracts and
  participant-support blocks are concatenated. Manually-entered other-direct
  costs (publication, shipping, …) are summed; the per-GRA tuition/fee costs are
  carried (tuition recomputes for the merged GRA count).
- **If a section still overflows** — e.g. more than four *distinct* GRA base
  salaries, or more than twelve senior personnel — the overflow entries are
  dropped and the tool prints a large, impossible-to-miss warning naming the
  section and the dropped entries. Add rows to that template section and re-run,
  or split the proposal.

> Consolidating lines means the spreadsheet applies its per-line `ROUND()` fewer
> times, so a consolidated total can differ from the naive sum of the separately
> rounded inputs by a few dollars (the consolidated figure is the more accurate
> one). Non-rounded categories match to the cent.

- **Final numbers are rounded to the nearest dollar.** Every figure the tool
  *produces* — the TeX macro values, the summary tables, and the derived dollar
  values written into the merged workbook (consolidated travel, supplies, and
  summed manual rows) — is rounded to the nearest whole dollar. Values *copied*
  from the inputs (base salaries, equipment amounts, ...) are never altered.
- **Workbook-wide inputs are checked for conflicts.** Some inputs cannot be
  summed — the salary/tuition inflation rates, the F&A base and rate type, and
  the per-GRA tuition/fee costs apply to the whole workbook. The merge carries
  the first input's values and prints the same large warning if the inputs
  disagree, since the merged totals then cannot equal the sum of the inputs.

Because the merged file stores formulas without cached values, open it once in
Excel (or another engine that evaluates formulas) to populate the computed
totals.

## Adapting to a new spreadsheet revision

All cell locations live in named constants at the top of
[`utkbudget/extractor.py`](utkbudget/extractor.py). If UTK ships a new revision
that moves rows around, update the row numbers there — the rest of the pipeline
follows automatically.

## Development

```bash
python -m unittest discover -s tests
```

The package is small and modular:

| Module | Responsibility |
| --- | --- |
| `utkbudget/extractor.py` | Read one workbook → a labelled `Budget` |
| `utkbudget/merge.py` | Sum budgets; write the merged `.xlsx` |
| `utkbudget/texdefs.py` | Write `\newcommand` definition files |
| `utkbudget/justification.py` | Generate the DOE justification document |
| `utkbudget/cli.py` | Walk the folder(s) and tie it all together |
