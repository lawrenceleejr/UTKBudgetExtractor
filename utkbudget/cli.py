"""Command-line entry point for UTKBudgetExtractor.

Hand the tool a folder of UTK proposal-budget spreadsheets and it will, for
every ``.xlsx`` input, write:

* ``<input>.tex`` -- a LaTeX ``\\newcommand`` definition file (one per input);
* ``merged.xlsx`` -- a single workbook merging all of the inputs;
* ``merged.tex`` -- definitions for the merged budget;
* ``justification.tex`` -- a DOE-style budget justification with a section per
  input file and a section for the combined sum.

If the folder instead contains *sub-folders* (e.g. ``Intensity Frontier``,
``Energy Frontier``, ``Theory Frontier``), each sub-folder is treated as a
program and gets its own merged workbook, definitions, and justification, and a
fully merged version across every program is produced at the top level.
"""

from __future__ import annotations

import argparse
import os
import sys
from typing import List, Tuple

from .extractor import Budget, extract_budget
from .justification import write_document
from .merge import merge_budgets, write_merged_workbook
from .texdefs import tex_prefix, write_defs

GRAND_PREFIX = "Combined"


def find_xlsx(folder: str) -> List[str]:
    """Return spreadsheet files directly inside ``folder`` (sorted)."""
    out = []
    for name in sorted(os.listdir(folder)):
        if name.startswith(("~$", ".")):
            continue  # Excel lock files / hidden files
        if name.lower().endswith((".xlsx", ".xlsm")):
            out.append(os.path.join(folder, name))
    return out


def find_xlsx_recursive(folder: str) -> List[str]:
    out = []
    for root, _dirs, files in os.walk(folder):
        for name in sorted(files):
            if name.startswith(("~$", ".")):
                continue
            if name.lower().endswith((".xlsx", ".xlsm")):
                out.append(os.path.join(root, name))
    return sorted(out)


def _unique(prefix: str, used: set) -> str:
    """Return a prefix not already in ``used`` (append A, B, ... on clash)."""
    if prefix not in used:
        used.add(prefix)
        return prefix
    i = 0
    while True:
        cand = f"{prefix}{chr(ord('A') + i)}"
        if cand not in used:
            used.add(cand)
            return cand
        i += 1


def process_group(name: str, files: List[str], out_dir: str,
                  group_prefix: str, file_stub: str = None
                  ) -> Tuple[Budget, str, List[Budget]]:
    """Process one group of spreadsheets into ``out_dir``.

    ``group_prefix`` is the LaTeX macro prefix for the merged total.
    ``file_stub`` is the filename stem for the merged/justification outputs;
    when ``None`` the merged products are simply ``merged.xlsx`` /
    ``merged.tex`` / ``justification.tex`` (used for a single flat folder).

    Returns ``(merged_budget, group_merged_tex_path, budgets)``.
    """
    if file_stub:
        merged_xlsx_name = f"{file_stub}_merged.xlsx"
        merged_tex_name = f"{file_stub}_merged.tex"
        just_name = f"{file_stub}_justification.tex"
    else:
        merged_xlsx_name = "merged.xlsx"
        merged_tex_name = "merged.tex"
        just_name = "justification.tex"
    os.makedirs(out_dir, exist_ok=True)
    print(f"\n=== {name}: {len(files)} file(s) -> {out_dir} ===")

    used_prefixes = {group_prefix}
    budgets: List[Budget] = []
    per_input: List[Tuple[str, Budget]] = []
    defs_inputs: List[str] = []          # relative \input paths for justification
    sections: List[Tuple[str, str, bool]] = []

    for path in files:
        base = os.path.splitext(os.path.basename(path))[0]
        print(f"  reading {os.path.basename(path)}")
        budget = extract_budget(path)
        budgets.append(budget)
        prefix = _unique(tex_prefix(base), used_prefixes)

        tex_name = f"{base}.tex"
        write_defs(budget, prefix, os.path.join(out_dir, tex_name))
        defs_inputs.append(tex_name)
        sections.append((prefix, base, False))
        per_input.append((base, budget))

    # Merge the group
    merged = merge_budgets(budgets, source=f"{name} (merged)")
    merged_xlsx = os.path.join(out_dir, merged_xlsx_name)
    write_merged_workbook(merged, per_input, merged_xlsx)
    print(f"  wrote {os.path.relpath(merged_xlsx)}")

    write_defs(merged, group_prefix, os.path.join(out_dir, merged_tex_name),
               header_note=f"Merged total for {name}")
    defs_inputs.append(merged_tex_name)
    sections.append((group_prefix, f"{name} (Combined)", True))

    just_path = os.path.join(out_dir, just_name)
    write_document(
        just_path,
        title=f"Budget Justification --- {name}",
        defs_inputs=defs_inputs,
        sections=sections,
        intro=(
            "This document justifies the funds requested from the U.S. Department "
            f"of Energy for the {name} budget. A justification is provided for each "
            "contributing budget, followed by a justification of the combined total."
        ),
    )
    print(f"  wrote {os.path.relpath(just_path)}")

    return merged, os.path.join(out_dir, merged_tex_name), budgets


def run(input_dir: str, output_dir: str) -> None:
    if not os.path.isdir(input_dir):
        sys.exit(f"error: input path is not a directory: {input_dir}")
    os.makedirs(output_dir, exist_ok=True)

    # Identify program sub-folders (immediate subdirectories that hold xlsx).
    subgroups: List[Tuple[str, str, List[str]]] = []
    for name in sorted(os.listdir(input_dir)):
        sub = os.path.join(input_dir, name)
        if os.path.isdir(sub):
            files = find_xlsx_recursive(sub)
            if files:
                subgroups.append((name, sub, files))

    root_files = find_xlsx(input_dir)

    if not subgroups:
        # ---- Flat mode: a single folder of spreadsheets ----------------
        if not root_files:
            sys.exit(f"error: no .xlsx files found in {input_dir}")
        process_group("All Budgets", root_files, output_dir, GRAND_PREFIX)
        print(f"\nDone. Outputs written to {output_dir}/")
        return

    # ---- Grouped mode: one program per sub-folder ----------------------
    group_used = set()
    group_merged: List[Tuple[str, Budget]] = []
    master_defs_inputs: List[str] = []
    master_sections: List[Tuple[str, str, bool]] = []
    all_budgets: List[Budget] = []

    # Any loose files at the root are treated as their own group.
    pending = list(subgroups)
    if root_files:
        pending.insert(0, (os.path.basename(os.path.normpath(input_dir)) or "Root",
                           input_dir, root_files))

    for name, _src, files in pending:
        gprefix = _unique(tex_prefix(name) + "Sum", group_used)
        group_out = os.path.join(output_dir, tex_prefix(name) or "group")
        merged, merged_tex_path, budgets = process_group(
            name, files, group_out, gprefix, file_stub=tex_prefix(name))
        group_merged.append((name, merged))
        all_budgets.extend(budgets)
        # Relative path from output_dir for the master document's \input
        rel = os.path.relpath(merged_tex_path, output_dir)
        master_defs_inputs.append(rel)
        master_sections.append((gprefix, name, False))

    # ---- Fully merged version across every program ---------------------
    print(f"\n=== Fully merged ({len(all_budgets)} file(s)) -> {output_dir} ===")
    grand = merge_budgets(all_budgets, source="All programs (merged)")
    grand_xlsx = os.path.join(output_dir, "merged.xlsx")
    write_merged_workbook(grand, group_merged, grand_xlsx)
    print(f"  wrote {os.path.relpath(grand_xlsx)}")

    write_defs(grand, GRAND_PREFIX, os.path.join(output_dir, "merged.tex"),
               header_note="Fully merged total across all programs")
    master_defs_inputs.append("merged.tex")
    master_sections.append((GRAND_PREFIX, "All Programs (Combined)", True))

    master_just = os.path.join(output_dir, "justification.tex")
    write_document(
        master_just,
        title="Budget Justification --- All Programs",
        defs_inputs=master_defs_inputs,
        sections=master_sections,
        intro=(
            "This master justification covers the full request to the U.S. "
            "Department of Energy across all programs. A justification is provided "
            "for each program, followed by a justification of the fully merged "
            "total. Per-file justifications are available in each program's "
            "sub-folder."
        ),
    )
    print(f"  wrote {os.path.relpath(master_just)}")
    print(f"\nDone. Outputs written to {output_dir}/")


def main(argv=None) -> None:
    parser = argparse.ArgumentParser(
        prog="utkbudget",
        description="Extract UTK proposal-budget spreadsheets into LaTeX "
                    "definitions, a merged workbook, and a DOE budget "
                    "justification document.",
    )
    parser.add_argument("input_dir",
                        help="Folder of .xlsx budgets (optionally with program "
                             "sub-folders).")
    parser.add_argument("-o", "--output-dir", default="output",
                        help="Where to write outputs (default: ./output).")
    args = parser.parse_args(argv)
    run(args.input_dir, args.output_dir)


if __name__ == "__main__":
    main()
