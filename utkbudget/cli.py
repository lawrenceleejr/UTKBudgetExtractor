"""Command-line entry point for UTKBudgetExtractor.

Hand the tool a folder of UTK proposal-budget spreadsheets and it will, for
every ``.xlsx`` input, write:

* ``<input>.tex`` -- a LaTeX ``\\newcommand`` definition file (one per input);
* ``<input>_justification.tex`` -- a DOE-style justification for that budget
  that ``\\input``s its own defs file, so it compiles on its own and drops
  straight into a larger proposal;
* ``merged.xlsx`` / ``merged.tex`` -- a single workbook merging all inputs and
  its definitions;
* ``justification.tex`` -- a justification of the combined sum;
* ``all_justifications.tex`` -- every justification in one document
  (``latexmk -pdf all_justifications.tex``);
* ``faculty_summary.tex`` / ``faculty_summary_by_year.tex`` -- request tables,
  one row per faculty (the latter with one column per year).

If the folder instead contains *sub-folders* (e.g. ``Program 1``,
``Program 2``, ``Program 3``), each sub-folder is treated as a program and gets
its own merged workbook, definitions, and justifications, and a fully merged
version across every program is produced.

Every generated file is written into ONE flat output directory -- no
sub-directories -- so a larger document can ``\\input`` them all from one place.
"""

from __future__ import annotations

import argparse
import os
import sys
import textwrap
from typing import List, Tuple

from .extractor import PERIOD_WORDS, Budget, extract_budget
from .justification import (build_faculty_summary, build_faculty_summary_by_year,
                            build_pdf_driver, write_document)
from .merge import (MergeIssue, RowOverflow, ValueConflict, merge_budgets,
                    write_merged_workbook)
from .texdefs import escape_tex, tex_prefix, write_defs

GRAND_PREFIX = "Combined"


def warn_merge_issues(context: str, xlsx_path: str,
                      issues: List[MergeIssue]) -> None:
    """Print an impossible-to-miss warning for merge problems (a section ran
    out of rows, or the inputs disagree on a workbook-wide value).

    Long messages are WRAPPED, never truncated, so the full warning is always
    visible in the terminal."""
    width = 78
    inner = width - 8  # room inside "!!! ... !!!"

    def rows(text: str = "", center: bool = False, indent: str = "") -> list:
        # Wrap to the box width so nothing is cut off; blank text -> one blank row.
        segments = textwrap.wrap(text, inner, subsequent_indent=indent) if text else [""]
        out = []
        for seg in segments:
            body = seg.center(inner) if center else seg.ljust(inner)
            out.append(f"!!! {body} !!!")
        return out

    overflows = [i for i in issues if isinstance(i, RowOverflow)]
    conflicts = [i for i in issues if isinstance(i, ValueConflict)]

    bar = "!" * width
    lines = ["", bar, bar]
    lines += rows("PROBLEMS WHILE MERGING -- CHECK THE MERGED SPREADSHEET", center=True)
    lines += [bar]
    lines += rows(f"Context : {context}")
    lines += rows(f"Workbook: {os.path.relpath(xlsx_path)}")
    if overflows:
        lines += rows() + rows("RAN OUT OF ROWS (overflow entries were DROPPED):")
        for of in overflows:
            n = of.needed - of.capacity
            lines += rows(f"  Section '{of.section}' has {of.capacity} row(s) "
                          f"but the merge needs {of.needed}.", indent="    ")
            lines += rows(f"    -> {n} DROPPED: {', '.join(of.dropped)}", indent="       ")
        lines += rows("  Add rows to those template sections, or split the "
                      "proposal, then re-run.", indent="  ")
    if conflicts:
        lines += rows() + rows("CONFLICTING WORKBOOK-WIDE INPUTS:")
        for vc in conflicts:
            lines += rows(f"  {vc.section}: {vc.detail}", indent="    ")
        lines += rows("  The merged totals will NOT equal the sum of the inputs "
                      "until fixed.", indent="  ")
    lines += [bar, bar, ""]
    sys.stderr.write("\n".join(lines) + "\n")
    sys.stderr.flush()


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


def _pi_name(budget) -> str:
    """A display name for the budget: the PI Name(s) cell, else the first
    named senior person, else empty."""
    pi = budget.get("PINames")
    if pi and str(pi).strip():
        return str(pi).strip()
    for i in range(12):
        name = budget.get(f"Senior{chr(65 + i)}Name")
        if name and str(name).strip():
            return str(name).strip()
    return ""


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


def _unique_name(base: str, used: set) -> str:
    """A filename stem not already taken (``PI_Smith``, ``PI_Smith_2``, ...).

    Needed because every generated file goes in ONE flat output directory."""
    if base not in used:
        used.add(base)
        return base
    i = 2
    while f"{base}_{i}" in used:
        i += 1
    used.add(f"{base}_{i}")
    return f"{base}_{i}"


def process_group(name: str, files: List[str], out_dir: str,
                  group_prefix: str, file_stub: str = None,
                  used_names: set = None
                  ) -> Tuple[Budget, str, List[Budget]]:
    """Process one group of spreadsheets into ``out_dir``.

    ``group_prefix`` is the LaTeX macro prefix for the merged total.
    ``file_stub`` is the filename stem for the merged/combined outputs; when
    ``None`` they are simply ``merged.xlsx`` / ``merged.tex`` /
    ``justification.tex`` (used for a single flat folder).

    Each input file gets its OWN ``<file>_justification.tex`` that ``\\input``s
    that file's dedicated defs file, so a single PI can compile theirs (or drop
    it into a larger proposal).  The combined-total justification is separate.

    Every group writes into the SAME flat ``out_dir`` -- no sub-directories --
    so ``used_names`` is threaded through all the groups to keep like-named
    spreadsheets in different program folders from overwriting each other.

    Returns ``(merged_budget, group_merged_tex_path, budgets,
    individual_justification_paths, combined_justification_path)``.
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
    if used_names is None:
        used_names = set()
    print(f"\n=== {name}: {len(files)} file(s) -> {out_dir} ===")

    used_prefixes = {group_prefix}
    budgets: List[Budget] = []
    individual_paths: List[str] = []

    for path in files:
        base = os.path.splitext(os.path.basename(path))[0]
        print(f"  reading {os.path.basename(path)}")
        budget = extract_budget(path)
        budgets.append(budget)
        prefix = _unique(tex_prefix(base), used_prefixes)
        # Everything lands in one flat directory, so two programs holding a
        # like-named spreadsheet would otherwise overwrite each other.
        base = _unique_name(base, used_names)

        tex_name = f"{base}.tex"
        write_defs(budget, prefix, os.path.join(out_dir, tex_name))

        # Justification for this single budget: it \input{}s its own dedicated
        # defs file (unique prefix), so the numbers live in one place and the
        # file can be dropped into a larger proposal without macro clashes.  The
        # rendered document carries no source filename -- generic title, the PI
        # name as a subtitle, no filename-derived heading.
        ind_path = os.path.join(out_dir, f"{base}_justification.tex")
        write_document(
            ind_path,
            title="Budget Justification",
            subtitle=_pi_name(budget) or None,   # PI name as a subtitle
            defs_inputs=[tex_name],
            sections=[(prefix, budget, None, False, True)],
            intro="",                            # no boilerplate lead sentence
        )
        individual_paths.append(ind_path)

    # Merged budget definitions (summed field-by-field, drives the .tex).
    merged = merge_budgets(budgets, source=f"{name} (merged)")

    # Merged workbook, in the same format as the inputs.
    merged_xlsx = os.path.join(out_dir, merged_xlsx_name)
    issues = write_merged_workbook(files, merged_xlsx)
    print(f"  wrote {os.path.relpath(merged_xlsx)}")
    if issues:
        warn_merge_issues(name, merged_xlsx, issues)

    write_defs(merged, group_prefix, os.path.join(out_dir, merged_tex_name),
               header_note=f"Merged total for {name}")

    # Separate combined-total justification (the sum only).
    just_path = os.path.join(out_dir, just_name)
    write_document(
        just_path,
        title=f"Budget Justification --- {name} (Combined)",
        defs_inputs=[merged_tex_name],
        sections=[(group_prefix, merged, None, True, False)],
        intro=("This document justifies the combined budget -- the sum of the "
               f"{len(files)} contributing budget(s) in {escape_tex(name)} -- "
               "requested from the U.S. Department of Energy."),
    )
    print(f"  wrote {os.path.relpath(just_path)}")

    return (merged, os.path.join(out_dir, merged_tex_name), budgets,
            individual_paths, just_path)


def _val(b: Budget, name: str) -> float:
    v = b.get(name)
    try:
        return float(v) if v not in (None, "") else 0.0
    except (TypeError, ValueError):
        return 0.0


def _faculty_summary_entries(budgets: List[Budget]):
    """(faculty name, direct, indirect, total) for each budget, for the summary
    table -- one row per faculty with their final DOE ask."""
    return [(_pi_name(b) or "(unnamed)",
             _val(b, "DirectTotal"), _val(b, "IndirectTotal"),
             _val(b, "GrandTotal"))
            for b in budgets]


def _faculty_year_entries(budgets: List[Budget]):
    """(faculty name, [per-period total, ...]) for the by-year summary table --
    each PI's total request in each budget period."""
    return [(_pi_name(b) or "(unnamed)",
             [_val(b, f"GrandYear{w}") for w in PERIOD_WORDS])
            for b in budgets]


def _write_faculty_summary(output_dir: str, budgets: List[Budget] = None,
                           groups: List[Tuple[str, List[Budget]]] = None) -> None:
    """Write the two summary tables: ``faculty_summary.tex`` (one row per faculty
    with their direct/indirect/total ask) and ``faculty_summary_by_year.tex``
    (one row per faculty, one column per year).

    Pass ``budgets`` for flat tables, or ``groups`` (``(sub-folder name,
    budgets)`` pairs) to group the PIs by thrust with per-thrust subtotals."""
    for fname, build, rows in (
            ("faculty_summary.tex", build_faculty_summary, _faculty_summary_entries),
            ("faculty_summary_by_year.tex", build_faculty_summary_by_year,
             _faculty_year_entries)):
        if groups is not None:
            content = build(groups=[(name, rows(bs)) for name, bs in groups])
        else:
            content = build(rows(budgets))
        path = os.path.join(output_dir, fname)
        with open(path, "w") as fh:
            fh.write(content)
        print(f"  wrote {os.path.relpath(path)}")


def _write_pdf_driver(output_dir: str, just_paths: List[str]) -> None:
    """Write ``all_justifications.tex`` -- a driver that compiles every
    justification into a single PDF -- and print the compile command."""
    rel = [os.path.basename(p) for p in just_paths]   # all in one flat directory
    driver = os.path.join(output_dir, "all_justifications.tex")
    with open(driver, "w") as fh:
        fh.write(build_pdf_driver(rel))
    print(f"  wrote {os.path.relpath(driver)}")
    print("\n  Compile every justification into one PDF with:")
    print(f"    latexmk -pdf -cd {os.path.join(output_dir, 'all_justifications.tex')}")
    print("    (or run pdflatex on it twice)")


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
        _, _, budgets, individual_paths, combined_path = process_group(
            "All Budgets", root_files, output_dir, GRAND_PREFIX)
        _write_faculty_summary(output_dir, budgets)
        _write_pdf_driver(output_dir, individual_paths + [combined_path])
        print(f"\nDone. Outputs written to {output_dir}/")
        return

    # ---- Grouped mode: one program per sub-folder ----------------------
    group_used = set()
    group_merged: List[Tuple[str, Budget]] = []
    master_defs_inputs: List[str] = []
    master_sections: List[Tuple[str, str, bool]] = []
    all_budgets: List[Budget] = []
    all_files: List[str] = []
    all_just_paths: List[str] = []
    group_budgets: List[Tuple[str, List[Budget]]] = []

    # Any loose files at the root are treated as their own group.
    pending = list(subgroups)
    if root_files:
        pending.insert(0, (os.path.basename(os.path.normpath(input_dir)) or "Root",
                           input_dir, root_files))

    used_names: set = set()
    for name, _src, files in pending:
        gprefix = _unique(tex_prefix(name) + "Sum", group_used)
        # Flat output: every program writes into the one output directory.
        merged, merged_tex_path, budgets, individual_paths, _combined = process_group(
            name, files, output_dir, gprefix, file_stub=tex_prefix(name),
            used_names=used_names)
        group_merged.append((name, merged))
        group_budgets.append((name, budgets))
        all_budgets.extend(budgets)
        all_files.extend(files)
        all_just_paths.extend(individual_paths)   # combined-per-program omitted
                                                  # from the booklet (the master
                                                  # already sums each program)
        # Relative path from output_dir for the master document's \input
        rel = os.path.relpath(merged_tex_path, output_dir)
        master_defs_inputs.append(rel)
        master_sections.append((gprefix, merged, name, False, False))

    # ---- Fully merged version across every program ---------------------
    print(f"\n=== Fully merged ({len(all_files)} file(s)) -> {output_dir} ===")
    grand = merge_budgets(all_budgets, source="All programs (merged)")
    grand_xlsx = os.path.join(output_dir, "merged.xlsx")
    issues = write_merged_workbook(all_files, grand_xlsx)
    print(f"  wrote {os.path.relpath(grand_xlsx)}")
    if issues:
        warn_merge_issues("All Programs", grand_xlsx, issues)

    write_defs(grand, GRAND_PREFIX, os.path.join(output_dir, "merged.tex"),
               header_note="Fully merged total across all programs")
    master_defs_inputs.append("merged.tex")
    master_sections.append((GRAND_PREFIX, grand, "All Programs (Combined)", True, False))

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
    all_just_paths.append(master_just)

    _write_faculty_summary(output_dir, groups=group_budgets)
    _write_pdf_driver(output_dir, all_just_paths)
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
