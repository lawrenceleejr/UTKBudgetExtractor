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
2. **A merged Excel workbook** (`merged.xlsx`) **in the same format as the
   inputs**, combining all of the budgets (see [Merging](#how-the-merge-works)).
3. **Merged LaTeX definitions** (`merged.tex`) for the combined budget.
4. **A DOE budget-justification document** (`justification.tex`) with a written
   justification for each input file followed by a justification of the combined
   sum, with travel called out in detail.

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
├── PI_Smith.tex          # \newcommand defs for Smith
├── PI_Jones.tex          # \newcommand defs for Jones
├── merged.xlsx           # Smith + Jones, summed
├── merged.tex            # \newcommand defs for the sum
└── justification.tex     # justification per file + the combined sum
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

produces

```
output/
├── ProgramOne/
│   ├── PI_Smith.tex
│   ├── PI_Jones.tex
│   ├── ProgramOne_merged.xlsx
│   ├── ProgramOne_merged.tex
│   └── ProgramOne_justification.tex   # per file + program sum
├── ProgramTwo/ ...
├── ProgramThree/ ...
├── merged.xlsx          # fully merged across every program
├── merged.tex           # defs for the grand total
└── justification.tex    # per-program + fully merged grand total
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

- **Line items are concatenated.** Each input's senior personnel, postdocs,
  other professionals, GRAs, undergraduates, admin/clerical, other personnel,
  and equipment items are stacked, in file order, into the template's sections.
  For each person the *inputs* are copied (name, raise flag, UT/JFO, base
  salary, appointment, person-months for all five periods, and tenure flags);
  the salary and fringe amounts are formulas and recompute. The fringe-rate
  cells are institutional constants that ship with the template and are never
  touched.
- **Detail sheets are merged too.** Travel trips (per period, domestic/foreign),
  supply lines, subcontract lines, and participant-support blocks are
  concatenated on the `TRAVEL`, `SUPPLIES`, `SUBCONTRACTS`, and `PARTICIPANT
  SUPPORT COSTS` sheets, so the main-sheet totals that reference them recompute.
- **Manually-entered other-direct costs** (publication, shipping, etc.) are
  summed; the per-GRA tuition/fee costs are carried (tuition recomputes for the
  merged GRA count).
- **If a section runs out of rows** — e.g. more than 12 senior personnel, more
  than 4 GRAs, or more than 10 domestic trips in a period across the merged
  proposals — the overflow entries are dropped and the tool prints a large,
  impossible-to-miss warning naming the section and the dropped entries. Add
  rows to that section in the template and re-run, or split the proposal.
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
