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
2. **A merged Excel workbook** (`merged.xlsx`) combining all of the inputs — a
   "Merged" summary sheet plus one sheet per input file.
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
