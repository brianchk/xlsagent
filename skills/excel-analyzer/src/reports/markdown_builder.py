"""Agent-optimized Markdown report generator."""

from __future__ import annotations

from pathlib import Path

from xls_extract import (
    FormulaCategory,
    SheetVisibility,
    WorkbookAnalysis,
)


class MarkdownReportBuilder:
    """Generates agent-optimized Markdown documentation."""

    def __init__(self, analysis: WorkbookAnalysis, output_dir: Path):
        """Initialize the builder.

        Args:
            analysis: The workbook analysis results
            output_dir: Directory to write markdown files
        """
        self.analysis = analysis
        self.output_dir = output_dir

    def build(self) -> None:
        """Generate all markdown files."""
        self.output_dir.mkdir(parents=True, exist_ok=True)

        # Create subdirectories
        (self.output_dir / "sheets").mkdir(exist_ok=True)
        (self.output_dir / "formulas").mkdir(exist_ok=True)
        (self.output_dir / "features").mkdir(exist_ok=True)
        (self.output_dir / "issues").mkdir(exist_ok=True)

        if self.analysis.vba_modules:
            (self.output_dir / "vba").mkdir(exist_ok=True)

        if self.analysis.power_queries:
            (self.output_dir / "power_query").mkdir(exist_ok=True)

        if self.analysis.screenshots:
            (self.output_dir / "screenshots").mkdir(exist_ok=True)

        # Generate files
        self._write_readme()
        self._write_summary()
        self._write_sheets()
        self._write_formulas()
        self._write_features()
        self._write_issues()
        self._write_vba()
        self._write_power_query()
        self._write_screenshots_index()

    def _write_readme(self) -> None:
        """Write the main README.md entry point."""
        content = f"""# {self.analysis.file_name} Analysis

This directory contains a comprehensive analysis of the Excel workbook.

## Quick Navigation

- [Summary](summary.md) - Key facts and statistics
- [Sheets](sheets/_index.md) - All worksheets
- [Formulas](formulas/_index.md) - Formula analysis
- [Features](features/) - Conditional formatting, validations, etc.
- [Issues](issues/) - Errors and external references
"""

        if self.analysis.vba_modules:
            content += "- [VBA](vba/_index.md) - VBA macro code\n"

        if self.analysis.power_queries:
            content += "- [Power Query](power_query/_index.md) - M code\n"

        if self.analysis.screenshots:
            content += "- [Screenshots](screenshots/_index.md) - Visual captures\n"

        content += f"""
## File Info

| Property | Value |
|----------|-------|
| File Name | {self.analysis.file_name} |
| File Size | {self._format_size(self.analysis.file_size)} |
| Macro Enabled | {'Yes' if self.analysis.is_macro_enabled else 'No'} |
| Sheet Count | {len(self.analysis.sheets)} |
| Formula Count | {len(self.analysis.formulas)} |
"""

        self._write_file("README.md", content)

    def _write_summary(self) -> None:
        """Write the summary.md file for quick context."""
        a = self.analysis

        # Count formulas by category
        formula_cats = {}
        for f in a.formulas:
            cat = f.category.value
            formula_cats[cat] = formula_cats.get(cat, 0) + 1

        # Count visible/hidden sheets
        visible = sum(1 for s in a.sheets if s.visibility == SheetVisibility.VISIBLE)
        hidden = sum(1 for s in a.sheets if s.visibility == SheetVisibility.HIDDEN)
        very_hidden = sum(1 for s in a.sheets if s.visibility == SheetVisibility.VERY_HIDDEN)

        content = f"""# Summary: {a.file_name}

## At a Glance

- **{len(a.sheets)} sheets** ({visible} visible, {hidden} hidden, {very_hidden} very hidden)
- **{len(a.formulas)} formulas** across all sheets
- **{len(a.named_ranges)} named ranges** ({sum(1 for n in a.named_ranges if n.is_lambda)} LAMBDA functions)
- **{len(a.tables)} structured tables**
- **{len(a.pivot_tables)} pivot tables**
- **{len(a.charts)} charts**

## Features Present

"""
        features = []
        if a.conditional_formats:
            features.append(f"- Conditional Formatting ({len(a.conditional_formats)} rules)")
        if a.data_validations:
            features.append(f"- Data Validations ({len(a.data_validations)} rules)")
        if a.vba_modules:
            features.append(f"- VBA Macros ({len(a.vba_modules)} modules)")
        if a.power_queries:
            features.append(f"- Power Query ({len(a.power_queries)} queries)")
        if a.has_dax:
            features.append(f"- Power Pivot/DAX ({a.dax_detection_note or 'detected'})")
        if a.comments:
            features.append(f"- Comments ({len(a.comments)} total)")
        if a.hyperlinks:
            features.append(f"- Hyperlinks ({len(a.hyperlinks)} total)")
        if a.controls:
            features.append(f"- Form Controls ({len(a.controls)} total)")
        if a.connections:
            features.append(f"- Data Connections ({len(a.connections)} total)")

        content += "\n".join(features) if features else "- No advanced features detected"

        content += "\n\n## Formula Categories\n\n"
        if formula_cats:
            for cat, count in sorted(formula_cats.items(), key=lambda x: -x[1]):
                content += f"- {cat}: {count}\n"
        else:
            content += "- No formulas found\n"

        content += "\n## Issues\n\n"
        if a.error_cells:
            content += f"- **{len(a.error_cells)} error cells** (see issues/errors.md)\n"
        if a.external_refs:
            content += f"- **{len(a.external_refs)} external references** (see issues/external_refs.md)\n"
        if not a.error_cells and not a.external_refs:
            content += "- No issues detected\n"

        if a.warnings:
            content += "\n## Extraction Warnings\n\n"
            for w in a.warnings:
                content += f"- {w.extractor}: {w.message}\n"

        self._write_file("summary.md", content)

    def _write_sheets(self) -> None:
        """Write sheet documentation."""
        # Index file
        index_content = "# Sheets\n\n"
        index_content += "| # | Name | Visibility | Rows | Cols | Features |\n"
        index_content += "|---|------|------------|------|------|----------|\n"

        for sheet in self.analysis.sheets:
            features = []
            if sheet.has_formulas:
                features.append("formulas")
            if sheet.has_charts:
                features.append("charts")
            if sheet.has_pivots:
                features.append("pivots")
            if sheet.has_tables:
                features.append("tables")
            if sheet.has_conditional_formatting:
                features.append("CF")
            if sheet.has_data_validation:
                features.append("DV")
            if sheet.has_comments:
                features.append("comments")

            features_str = ", ".join(features) if features else "-"
            vis = sheet.visibility.value

            index_content += f"| {sheet.index + 1} | [{sheet.name}]({self._sanitize_filename(sheet.name)}.md) | {vis} | {sheet.row_count} | {sheet.col_count} | {features_str} |\n"

        self._write_file("sheets/_index.md", index_content)

        # Individual sheet files
        for sheet in self.analysis.sheets:
            self._write_sheet_detail(sheet)

    def _write_sheet_detail(self, sheet) -> None:
        """Write detailed documentation for a single sheet."""
        content = f"# Sheet: {sheet.name}\n\n"

        content += f"""## Overview

| Property | Value |
|----------|-------|
| Index | {sheet.index + 1} |
| Visibility | {sheet.visibility.value} |
| Used Range | {sheet.used_range or 'Empty'} |
| Rows | {sheet.row_count} |
| Columns | {sheet.col_count} |
"""

        if sheet.tab_color:
            content += f"| Tab Color | {sheet.tab_color} |\n"

        if sheet.merged_cell_ranges:
            content += f"\n## Merged Cells ({len(sheet.merged_cell_ranges)})\n\n"
            for r in sheet.merged_cell_ranges[:20]:  # Limit display
                content += f"- {r}\n"
            if len(sheet.merged_cell_ranges) > 20:
                content += f"\n*...and {len(sheet.merged_cell_ranges) - 20} more*\n"

        # Formulas in this sheet
        sheet_formulas = [f for f in self.analysis.formulas if f.location.sheet == sheet.name]
        if sheet_formulas:
            content += f"\n## Formulas ({len(sheet_formulas)})\n\n"
            content += "| Cell | Category | Formula |\n"
            content += "|------|----------|--------|\n"
            for f in sheet_formulas[:50]:
                formula_preview = f.formula_clean[:60] + "..." if len(f.formula_clean) > 60 else f.formula_clean
                content += f"| {f.location.cell} | {f.category.value} | `{formula_preview}` |\n"
            if len(sheet_formulas) > 50:
                content += f"\n*...and {len(sheet_formulas) - 50} more formulas*\n"

        # Tables in this sheet
        sheet_tables = [t for t in self.analysis.tables if t.sheet == sheet.name]
        if sheet_tables:
            content += f"\n## Tables ({len(sheet_tables)})\n\n"
            for t in sheet_tables:
                content += f"### {t.display_name}\n"
                content += f"- Range: {t.range}\n"
                content += f"- Columns: {', '.join(t.columns)}\n"
                if t.style_name:
                    content += f"- Style: {t.style_name}\n"
                content += "\n"

        # Pivot tables in this sheet
        sheet_pivots = [p for p in self.analysis.pivot_tables if p.sheet == sheet.name]
        if sheet_pivots:
            content += f"\n## Pivot Tables ({len(sheet_pivots)})\n\n"
            for p in sheet_pivots:
                content += f"### {p.name}\n"
                content += f"- Location: {p.location}\n"
                if p.row_fields:
                    content += f"- Row Fields: {', '.join(p.row_fields)}\n"
                if p.column_fields:
                    content += f"- Column Fields: {', '.join(p.column_fields)}\n"
                if p.data_fields:
                    content += f"- Data Fields: {', '.join(p.data_fields)}\n"
                content += "\n"

        # Charts in this sheet
        sheet_charts = [c for c in self.analysis.charts if c.sheet == sheet.name]
        if sheet_charts:
            content += f"\n## Charts ({len(sheet_charts)})\n\n"
            for c in sheet_charts:
                content += f"- **{c.name}**: {c.chart_type}"
                if c.title:
                    content += f' ("{c.title}")'
                content += "\n"

        filename = f"sheets/{self._sanitize_filename(sheet.name)}.md"
        self._write_file(filename, content)

    def _write_formulas(self) -> None:
        """Write formula documentation."""
        # Index file
        content = "# Formulas\n\n"

        # Summary by category
        cats = {}
        for f in self.analysis.formulas:
            cat = f.category
            cats[cat] = cats.get(cat, 0) + 1

        content += "## By Category\n\n"
        for cat in FormulaCategory:
            count = cats.get(cat, 0)
            if count > 0:
                content += f"- **{cat.value}**: {count}\n"

        # Named ranges with LAMBDA
        lambdas = [n for n in self.analysis.named_ranges if n.is_lambda]
        if lambdas:
            content += f"\n## LAMBDA Functions ({len(lambdas)})\n\n"
            for n in lambdas:
                content += f"### {n.name}\n\n```\n{n.value}\n```\n\n"

        # Regular named ranges
        regular_names = [n for n in self.analysis.named_ranges if not n.is_lambda]
        if regular_names:
            content += f"\n## Named Ranges ({len(regular_names)})\n\n"
            content += "| Name | Value | Scope |\n"
            content += "|------|-------|-------|\n"
            for n in regular_names:
                scope = n.scope or "Global"
                value_preview = n.value[:40] + "..." if len(n.value) > 40 else n.value
                content += f"| {n.name} | `{value_preview}` | {scope} |\n"

        # Complex formulas (dynamic array, LAMBDA usage)
        complex_formulas = [
            f for f in self.analysis.formulas
            if f.category in (FormulaCategory.DYNAMIC_ARRAY, FormulaCategory.LAMBDA)
        ]
        if complex_formulas:
            content += f"\n## Complex Formulas ({len(complex_formulas)})\n\n"
            for f in complex_formulas[:30]:
                content += f"### {f.location.address}\n\n"
                content += f"**Category**: {f.category.value}\n\n"
                content += f"```excel\n{f.formula_clean}\n```\n\n"

        self._write_file("formulas/_index.md", content)

    def _write_features(self) -> None:
        """Write feature documentation files."""
        # Conditional Formatting
        if self.analysis.conditional_formats:
            content = "# Conditional Formatting\n\n"
            content += f"Total rules: {len(self.analysis.conditional_formats)}\n\n"

            for cf in self.analysis.conditional_formats:
                content += f"## {cf.range}\n\n"
                content += f"- Type: {cf.rule_type.value}\n"
                content += f"- Description: {cf.description}\n"
                if cf.formula:
                    content += f"- Formula: `{cf.formula}`\n"
                content += "\n"

            self._write_file("features/conditional_formatting.md", content)

        # Data Validations
        if self.analysis.data_validations:
            content = "# Data Validations\n\n"
            content += f"Total rules: {len(self.analysis.data_validations)}\n\n"

            for dv in self.analysis.data_validations:
                content += f"## {dv.range}\n\n"
                content += f"- Type: {dv.type}\n"
                if dv.formula1:
                    content += f"- Formula/List: `{dv.formula1}`\n"
                if dv.input_message:
                    content += f"- Input Message: {dv.input_message}\n"
                if dv.error_message:
                    content += f"- Error Message: {dv.error_message}\n"
                content += "\n"

            self._write_file("features/data_validations.md", content)

        # Tables
        if self.analysis.tables:
            content = "# Structured Tables\n\n"
            content += f"Total tables: {len(self.analysis.tables)}\n\n"

            for t in self.analysis.tables:
                content += f"## {t.display_name}\n\n"
                content += f"- Sheet: {t.sheet}\n"
                content += f"- Range: {t.range}\n"
                content += f"- Columns: {len(t.columns)}\n"
                if t.columns:
                    content += f"  - {', '.join(t.columns)}\n"
                content += "\n"

            self._write_file("features/tables.md", content)

        # Charts
        if self.analysis.charts:
            content = "# Charts\n\n"
            content += f"Total charts: {len(self.analysis.charts)}\n\n"

            for c in self.analysis.charts:
                content += f"## {c.name}\n\n"
                content += f"- Sheet: {c.sheet}\n"
                content += f"- Type: {c.chart_type}\n"
                if c.title:
                    content += f"- Title: {c.title}\n"
                if c.data_range:
                    content += f"- Data: {c.data_range}\n"
                content += "\n"

            self._write_file("features/charts.md", content)

        # Pivot Tables
        if self.analysis.pivot_tables:
            content = "# Pivot Tables\n\n"
            content += f"Total pivot tables: {len(self.analysis.pivot_tables)}\n\n"

            for p in self.analysis.pivot_tables:
                content += f"## {p.name}\n\n"
                content += f"- Sheet: {p.sheet}\n"
                content += f"- Location: {p.location}\n"
                if p.row_fields:
                    content += f"- Row Fields: {', '.join(p.row_fields)}\n"
                if p.column_fields:
                    content += f"- Column Fields: {', '.join(p.column_fields)}\n"
                if p.data_fields:
                    content += f"- Data Fields: {', '.join(p.data_fields)}\n"
                if p.filter_fields:
                    content += f"- Filter Fields: {', '.join(p.filter_fields)}\n"
                content += "\n"

            self._write_file("features/pivot_tables.md", content)

        # Comments
        if self.analysis.comments:
            content = "# Comments\n\n"
            content += f"Total comments: {len(self.analysis.comments)}\n\n"

            for c in self.analysis.comments:
                content += f"## {c.location.address}\n\n"
                if c.author:
                    content += f"**Author**: {c.author}\n\n"
                content += f"{c.text}\n\n"
                if c.replies:
                    content += f"### Replies ({len(c.replies)})\n\n"
                    for r in c.replies:
                        author = r.author or "Unknown"
                        content += f"- **{author}**: {r.text}\n"
                    content += "\n"

            self._write_file("features/comments.md", content)

        # Hyperlinks
        if self.analysis.hyperlinks:
            content = "# Hyperlinks\n\n"
            content += f"Total hyperlinks: {len(self.analysis.hyperlinks)}\n\n"

            content += "| Location | Target | Display Text |\n"
            content += "|----------|--------|-------------|\n"
            for h in self.analysis.hyperlinks:
                display = h.display_text or "-"
                target_preview = h.target[:50] + "..." if len(h.target) > 50 else h.target
                content += f"| {h.location.address} | {target_preview} | {display} |\n"

            self._write_file("features/hyperlinks.md", content)

        # Controls
        if self.analysis.controls:
            content = "# Form Controls\n\n"
            content += f"Total controls: {len(self.analysis.controls)}\n\n"

            for c in self.analysis.controls:
                content += f"## {c.name}\n\n"
                content += f"- Type: {c.control_type}\n"
                content += f"- Sheet: {c.sheet}\n"
                if c.linked_cell:
                    content += f"- Linked Cell: {c.linked_cell}\n"
                if c.macro:
                    content += f"- Macro: {c.macro}\n"
                content += "\n"

            self._write_file("features/controls.md", content)

        # Connections
        if self.analysis.connections:
            content = "# Data Connections\n\n"
            content += f"Total connections: {len(self.analysis.connections)}\n\n"

            for c in self.analysis.connections:
                content += f"## {c.name}\n\n"
                content += f"- Type: {c.connection_type}\n"
                if c.connection_string:
                    content += f"- Connection: `{c.connection_string}`\n"
                if c.command_text:
                    content += f"- Command: `{c.command_text}`\n"
                content += "\n"

            self._write_file("features/connections.md", content)

        # Protection
        if self.analysis.protection:
            p = self.analysis.protection
            content = "# Protection Settings\n\n"

            content += "## Workbook Level\n\n"
            content += f"- Protected: {'Yes' if p.workbook_protected else 'No'}\n"
            if p.workbook_protected:
                content += f"- Structure: {'Protected' if p.workbook_structure else 'Not protected'}\n"
                content += f"- Windows: {'Protected' if p.workbook_windows else 'Not protected'}\n"

            content += "\n## Sheet Level\n\n"
            for sheet_name, details in p.sheets.items():
                protected = details.get("protected", False)
                content += f"### {sheet_name}\n\n"
                content += f"- Protected: {'Yes' if protected else 'No'}\n"
                if protected:
                    content += f"- Password Protected: {'Yes' if details.get('password_protected') else 'No'}\n"
                content += "\n"

            self._write_file("features/protection.md", content)

    def _write_issues(self) -> None:
        """Write issue documentation."""
        # Errors
        if self.analysis.error_cells:
            content = "# Error Cells\n\n"
            content += f"Total errors: {len(self.analysis.error_cells)}\n\n"

            # Group by error type
            by_type = {}
            for e in self.analysis.error_cells:
                t = e.error_type.value
                if t not in by_type:
                    by_type[t] = []
                by_type[t].append(e)

            for error_type, errors in by_type.items():
                content += f"## {error_type} ({len(errors)})\n\n"
                content += "| Location | Formula |\n"
                content += "|----------|--------|\n"
                for e in errors[:20]:
                    formula = e.formula or "-"
                    content += f"| {e.location.address} | `{formula}` |\n"
                if len(errors) > 20:
                    content += f"\n*...and {len(errors) - 20} more*\n"
                content += "\n"

            self._write_file("issues/errors.md", content)

        # External References
        if self.analysis.external_refs:
            content = "# External References\n\n"
            content += f"Total external references: {len(self.analysis.external_refs)}\n\n"

            # Group by target workbook
            by_workbook = {}
            for ref in self.analysis.external_refs:
                wb = ref.target_workbook
                if wb not in by_workbook:
                    by_workbook[wb] = []
                by_workbook[wb].append(ref)

            for workbook, refs in by_workbook.items():
                broken = any(r.is_broken for r in refs)
                status = " (BROKEN)" if broken else ""
                content += f"## {workbook}{status}\n\n"

                for ref in refs[:10]:
                    if ref.source_cell.cell:
                        content += f"- {ref.source_cell.address}"
                        if ref.target_sheet:
                            content += f" -> {ref.target_sheet}"
                        if ref.target_range:
                            content += f"!{ref.target_range}"
                        content += "\n"
                if len(refs) > 10:
                    content += f"\n*...and {len(refs) - 10} more references*\n"
                content += "\n"

            self._write_file("issues/external_refs.md", content)

    def _write_vba(self) -> None:
        """Write VBA documentation."""
        if not self.analysis.vba_modules:
            return

        # Index file
        content = "# VBA Modules\n\n"
        if self.analysis.vba_project_name:
            content += f"Project: {self.analysis.vba_project_name}\n\n"

        content += f"Total modules: {len(self.analysis.vba_modules)}\n\n"

        content += "| Module | Type | Lines | Procedures |\n"
        content += "|--------|------|-------|------------|\n"
        for m in self.analysis.vba_modules:
            procs = ", ".join(m.procedures[:5])
            if len(m.procedures) > 5:
                procs += f" (+{len(m.procedures) - 5} more)"
            content += f"| [{m.name}]({self._sanitize_filename(m.name)}.md) | {m.module_type} | {m.line_count} | {procs} |\n"

        self._write_file("vba/_index.md", content)

        # Individual module files
        for m in self.analysis.vba_modules:
            module_content = f"# {m.name}\n\n"
            module_content += f"**Type**: {m.module_type}\n\n"
            module_content += f"**Lines**: {m.line_count}\n\n"

            if m.procedures:
                module_content += "## Procedures\n\n"
                for p in m.procedures:
                    module_content += f"- {p}\n"
                module_content += "\n"

            module_content += "## Code\n\n```vb\n"
            module_content += m.code
            module_content += "\n```\n"

            filename = f"vba/{self._sanitize_filename(m.name)}.md"
            self._write_file(filename, module_content)

    def _write_power_query(self) -> None:
        """Write Power Query documentation."""
        if not self.analysis.power_queries:
            return

        # Index file
        content = "# Power Query\n\n"
        content += f"Total queries: {len(self.analysis.power_queries)}\n\n"

        content += "| Query | Description |\n"
        content += "|-------|-------------|\n"
        for q in self.analysis.power_queries:
            desc = q.description or "-"
            content += f"| [{q.name}]({self._sanitize_filename(q.name)}.md) | {desc} |\n"

        self._write_file("power_query/_index.md", content)

        # Individual query files
        for q in self.analysis.power_queries:
            query_content = f"# {q.name}\n\n"
            if q.description:
                query_content += f"{q.description}\n\n"

            query_content += "## M Code\n\n```powerquery\n"
            query_content += q.formula
            query_content += "\n```\n"

            filename = f"power_query/{self._sanitize_filename(q.name)}.md"
            self._write_file(filename, query_content)

    def _write_screenshots_index(self) -> None:
        """Write screenshots index."""
        if not self.analysis.screenshots:
            return

        content = "# Screenshots\n\n"
        content += f"Total screenshots: {len(self.analysis.screenshots)}\n\n"

        for s in self.analysis.screenshots:
            content += f"## {s.sheet}\n\n"
            content += f"![{s.sheet}]({s.path.name})\n\n"
            if s.captured_at:
                content += f"*Captured: {s.captured_at}*\n\n"

        self._write_file("screenshots/_index.md", content)

    def _write_file(self, relative_path: str, content: str) -> None:
        """Write content to a file."""
        path = self.output_dir / relative_path
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(content, encoding="utf-8")

    def _sanitize_filename(self, name: str) -> str:
        """Sanitize a string for use as a filename."""
        invalid_chars = '<>:"/\\|?*'
        result = name
        for char in invalid_chars:
            result = result.replace(char, "_")
        if len(result) > 100:
            result = result[:100]
        return result

    def _format_size(self, size_bytes: int) -> str:
        """Format file size in human-readable form."""
        for unit in ["B", "KB", "MB", "GB"]:
            if size_bytes < 1024:
                return f"{size_bytes:.1f} {unit}"
            size_bytes /= 1024
        return f"{size_bytes:.1f} TB"
