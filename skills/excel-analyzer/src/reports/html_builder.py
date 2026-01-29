"""Rich HTML report generator - Sheet-centric multi-page layout."""

from __future__ import annotations

import re
from pathlib import Path
from datetime import datetime
from collections import defaultdict

from pygments import highlight
from pygments.lexers import VbNetLexer, get_lexer_by_name
from pygments.formatters import HtmlFormatter

from ..models import SheetVisibility, WorkbookAnalysis, SheetInfo


class HTMLReportBuilder:
    """Generates a multi-page HTML report centered around sheets."""

    # Common prefixes to detect for grouping
    GROUP_PREFIXES = [
        ("pbi-", "Power BI Data"),
        ("ref-", "Reference Data"),
        ("config-", "Configuration"),
        ("cfg-", "Configuration"),
        ("data-", "Data"),
        ("raw-", "Raw Data"),
        ("src-", "Source Data"),
        ("calc-", "Calculations"),
        ("tmp-", "Temporary"),
        ("test-", "Test"),
    ]

    def __init__(self, analysis: WorkbookAnalysis, output_dir: Path):
        self.analysis = analysis
        self.output_dir = output_dir

        # Build cross-reference maps
        self._build_cross_references()

    def _build_cross_references(self):
        """Build maps of what each sheet contains and what references what."""
        a = self.analysis

        # Map sheet -> items it contains
        self.sheet_formulas = defaultdict(list)
        self.sheet_charts = defaultdict(list)
        self.sheet_pivots = defaultdict(list)
        self.sheet_tables = defaultdict(list)
        self.sheet_cfs = defaultdict(list)
        self.sheet_dvs = defaultdict(list)
        self.sheet_comments = defaultdict(list)
        self.sheet_hyperlinks = defaultdict(list)
        self.sheet_errors = defaultdict(list)
        self.sheet_controls = defaultdict(list)

        for f in a.formulas:
            self.sheet_formulas[f.location.sheet].append(f)
        for c in a.charts:
            self.sheet_charts[c.sheet].append(c)
        for p in a.pivot_tables:
            self.sheet_pivots[p.sheet].append(p)
        for t in a.tables:
            self.sheet_tables[t.sheet].append(t)
        for cf in a.conditional_formats:
            # CF range format is "Sheet!Range" or just "Range"
            sheet = self._extract_sheet_from_range(cf.range, a.sheets[0].name if a.sheets else "Sheet1")
            self.sheet_cfs[sheet].append(cf)
        for dv in a.data_validations:
            sheet = self._extract_sheet_from_range(dv.range, a.sheets[0].name if a.sheets else "Sheet1")
            self.sheet_dvs[sheet].append(dv)
        for c in a.comments:
            self.sheet_comments[c.location.sheet].append(c)
        for h in a.hyperlinks:
            self.sheet_hyperlinks[h.location.sheet].append(h)
        for e in a.error_cells:
            self.sheet_errors[e.location.sheet].append(e)
        for ctrl in a.controls:
            self.sheet_controls[ctrl.sheet].append(ctrl)

        # Map VBA module -> sheets that might use it (by searching for Sub/Function calls in formulas)
        self.vba_to_sheets = defaultdict(set)
        self.sheet_to_vba = defaultdict(set)

        if a.vba_modules:
            # Get all procedure names from VBA
            vba_procs = {}
            for m in a.vba_modules:
                for proc in (m.procedures or []):
                    vba_procs[proc.lower()] = m.name

            # Search formulas for procedure calls
            for f in a.formulas:
                formula_lower = f.formula_clean.lower()
                for proc, module in vba_procs.items():
                    if proc in formula_lower:
                        self.vba_to_sheets[module].add(f.location.sheet)
                        self.sheet_to_vba[f.location.sheet].add(module)

            # Also check controls for macro assignments
            for ctrl in a.controls:
                if ctrl.macro:
                    macro_name = ctrl.macro.split(".")[-1].lower()
                    for proc, module in vba_procs.items():
                        if proc == macro_name:
                            self.vba_to_sheets[module].add(ctrl.sheet)
                            self.sheet_to_vba[ctrl.sheet].add(module)

    def _extract_sheet_from_range(self, range_str: str, default_sheet: str) -> str:
        """Extract sheet name from a range like 'Sheet1!A1:B2'."""
        if "!" in range_str:
            sheet_part = range_str.split("!")[0]
            # Remove quotes if present
            return sheet_part.strip("'\"")
        return default_sheet

    def build(self) -> Path:
        """Generate the complete multi-page HTML report."""
        # Create directories
        self.output_dir.mkdir(parents=True, exist_ok=True)
        (self.output_dir / "sheets").mkdir(exist_ok=True)
        (self.output_dir / "workbook").mkdir(exist_ok=True)

        # Write shared CSS
        self._write_styles()

        # Generate index page
        index_path = self._generate_index()

        # Generate individual sheet pages
        for sheet in self.analysis.sheets:
            self._generate_sheet_page(sheet)

        # Generate workbook-wide pages
        if self.analysis.vba_modules:
            self._generate_vba_page()
        if self.analysis.power_queries:
            self._generate_power_query_page()
        if self.analysis.connections or self.analysis.external_refs:
            self._generate_connections_page()
        if self.analysis.named_ranges:
            self._generate_named_ranges_page()

        return index_path

    def _write_styles(self):
        """Write the shared CSS file."""
        css = self._get_styles()
        (self.output_dir / "styles.css").write_text(css, encoding="utf-8")

    def _group_sheets(self) -> dict[str, list[SheetInfo]]:
        """Group sheets by detected prefix patterns."""
        groups = defaultdict(list)
        ungrouped = []

        for sheet in self.analysis.sheets:
            name_lower = sheet.name.lower()
            matched = False

            for prefix, group_name in self.GROUP_PREFIXES:
                if name_lower.startswith(prefix):
                    groups[group_name].append(sheet)
                    matched = True
                    break

            if not matched:
                ungrouped.append(sheet)

        # Add ungrouped as "Main Sheets" or similar
        if ungrouped:
            groups["Main Sheets"] = ungrouped

        return dict(groups)

    def _generate_index(self) -> Path:
        """Generate the main index.html page."""
        a = self.analysis
        groups = self._group_sheets()

        # Build sheet groups HTML
        groups_html = ""
        for group_name, sheets in groups.items():
            cards = ""
            for s in sheets:
                features = self._get_sheet_feature_badges(s)
                vis_class = s.visibility.value.replace("_", "-")
                vis_badge = "" if s.visibility == SheetVisibility.VISIBLE else f'<span class="visibility-badge {vis_class}">{s.visibility.value}</span>'

                color_dot = ""
                if s.tab_color and s.tab_color.startswith("#") and len(s.tab_color) <= 9:
                    color_dot = f'<span class="color-dot" style="background-color: {s.tab_color}"></span>'

                cards += f"""
                <a href="sheets/{self._sheet_filename(s.name)}" class="sheet-card">
                    <div class="sheet-card-header">
                        {color_dot}
                        <span class="sheet-name">{self._escape(s.name)}</span>
                        {vis_badge}
                    </div>
                    <div class="sheet-card-meta">
                        {s.row_count} rows × {s.col_count} cols
                    </div>
                    <div class="sheet-card-features">
                        {features}
                    </div>
                </a>
                """

            groups_html += f"""
            <div class="sheet-group">
                <h3>{self._escape(group_name)} ({len(sheets)})</h3>
                <div class="sheet-cards">
                    {cards}
                </div>
            </div>
            """

        # Build workbook-wide nav
        workbook_nav = ""
        if a.vba_modules:
            workbook_nav += f'<a href="workbook/vba.html" class="workbook-link"><span class="icon">📜</span> VBA Modules ({len(a.vba_modules)})</a>'
        if a.power_queries:
            workbook_nav += f'<a href="workbook/power-query.html" class="workbook-link"><span class="icon">🔄</span> Power Query ({len(a.power_queries)})</a>'
        if a.connections or a.external_refs:
            conn_count = len(a.connections) + len(a.external_refs)
            workbook_nav += f'<a href="workbook/connections.html" class="workbook-link"><span class="icon">🔗</span> Connections ({conn_count})</a>'
        if a.named_ranges:
            workbook_nav += f'<a href="workbook/named-ranges.html" class="workbook-link"><span class="icon">📛</span> Named Ranges ({len(a.named_ranges)})</a>'

        # Quick stats
        stats = f"""
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-value">{len(a.sheets)}</div>
                <div class="stat-label">Sheets</div>
            </div>
            <div class="stat-card">
                <div class="stat-value">{len(a.formulas)}</div>
                <div class="stat-label">Formulas</div>
            </div>
            <div class="stat-card">
                <div class="stat-value">{len(a.charts)}</div>
                <div class="stat-label">Charts</div>
            </div>
            <div class="stat-card">
                <div class="stat-value">{len(a.pivot_tables)}</div>
                <div class="stat-label">Pivot Tables</div>
            </div>
            <div class="stat-card">
                <div class="stat-value">{len(a.tables)}</div>
                <div class="stat-label">Tables</div>
            </div>
        </div>
        """

        # Warnings block
        warnings = self._generate_warnings_block()

        html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{self._escape(a.file_name)} - Excel Analysis</title>
    <link rel="stylesheet" href="styles.css">
</head>
<body>
    <header class="page-header">
        <h1>{self._escape(a.file_name)}</h1>
        <p class="subtitle">Excel Workbook Analysis</p>
        <p class="meta">{self._format_size(a.file_size)} · Generated {datetime.now().strftime("%Y-%m-%d %H:%M")}</p>
    </header>

    <main>
        <section class="summary-section">
            <h2>Overview</h2>
            {stats}
            {warnings}
        </section>

        <section class="sheets-section">
            <h2>Sheets</h2>
            {groups_html}
        </section>

        {f'''<section class="workbook-section">
            <h2>Workbook-Wide</h2>
            <div class="workbook-links">
                {workbook_nav}
            </div>
        </section>''' if workbook_nav else ""}
    </main>

    <footer>
        <p>Generated by Excel Analyzer for Claude Code</p>
    </footer>
</body>
</html>"""

        path = self.output_dir / "index.html"
        path.write_text(html, encoding="utf-8")
        return path

    def _generate_sheet_page(self, sheet: SheetInfo):
        """Generate an individual sheet page."""
        a = self.analysis
        name = sheet.name

        # Gather all items for this sheet
        formulas = self.sheet_formulas.get(name, [])
        charts = self.sheet_charts.get(name, [])
        pivots = self.sheet_pivots.get(name, [])
        tables = self.sheet_tables.get(name, [])
        cfs = self.sheet_cfs.get(name, [])
        dvs = self.sheet_dvs.get(name, [])
        comments = self.sheet_comments.get(name, [])
        hyperlinks = self.sheet_hyperlinks.get(name, [])
        errors = self.sheet_errors.get(name, [])
        controls = self.sheet_controls.get(name, [])
        vba_refs = self.sheet_to_vba.get(name, set())

        # Screenshots for this sheet
        screenshots = [s for s in a.screenshots if s.sheet == name]

        # Build sections
        sections = []

        # Screenshot section
        if screenshots:
            sections.append(self._build_screenshot_section(screenshots))

        # Charts
        if charts:
            sections.append(self._build_charts_section(charts, name))

        # Pivot Tables
        if pivots:
            sections.append(self._build_pivots_section(pivots))

        # Tables
        if tables:
            sections.append(self._build_tables_section(tables))

        # Formulas
        if formulas:
            sections.append(self._build_formulas_section(formulas))

        # Conditional Formatting
        if cfs:
            sections.append(self._build_cf_section(cfs))

        # Data Validations
        if dvs:
            sections.append(self._build_dv_section(dvs))

        # Comments
        if comments:
            sections.append(self._build_comments_section(comments))

        # Controls
        if controls:
            sections.append(self._build_controls_section(controls))

        # Errors
        if errors:
            sections.append(self._build_errors_section(errors))

        # VBA References
        if vba_refs:
            sections.append(self._build_vba_refs_section(vba_refs))

        # No content message
        if not sections:
            sections.append('<div class="empty-state">This sheet has no special features to display.</div>')

        # Sheet metadata
        vis_class = sheet.visibility.value.replace("_", "-")
        visibility_html = f'<span class="visibility-badge {vis_class}">{sheet.visibility.value}</span>'

        color_html = ""
        if sheet.tab_color and sheet.tab_color.startswith("#") and len(sheet.tab_color) <= 9:
            color_html = f'<span class="color-dot large" style="background-color: {sheet.tab_color}"></span>'

        # Navigation for sections
        nav_items = []
        if screenshots:
            nav_items.append('<a href="#screenshots">Screenshots</a>')
        if charts:
            nav_items.append(f'<a href="#charts">Charts ({len(charts)})</a>')
        if pivots:
            nav_items.append(f'<a href="#pivots">Pivot Tables ({len(pivots)})</a>')
        if tables:
            nav_items.append(f'<a href="#tables">Tables ({len(tables)})</a>')
        if formulas:
            nav_items.append(f'<a href="#formulas">Formulas ({len(formulas)})</a>')
        if cfs:
            nav_items.append(f'<a href="#cf">Conditional Formatting ({len(cfs)})</a>')
        if dvs:
            nav_items.append(f'<a href="#dv">Data Validation ({len(dvs)})</a>')
        if comments:
            nav_items.append(f'<a href="#comments">Comments ({len(comments)})</a>')
        if controls:
            nav_items.append(f'<a href="#controls">Controls ({len(controls)})</a>')
        if errors:
            nav_items.append(f'<a href="#errors">Errors ({len(errors)})</a>')
        if vba_refs:
            nav_items.append(f'<a href="#vba">VBA ({len(vba_refs)})</a>')

        nav_html = " · ".join(nav_items) if nav_items else ""

        html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{self._escape(name)} - {self._escape(a.file_name)}</title>
    <link rel="stylesheet" href="../styles.css">
</head>
<body>
    <nav class="breadcrumb">
        <a href="../index.html">← Back to Index</a>
    </nav>

    <header class="page-header sheet-header">
        <div class="sheet-title">
            {color_html}
            <h1>{self._escape(name)}</h1>
            {visibility_html}
        </div>
        <div class="sheet-meta">
            {sheet.row_count} rows × {sheet.col_count} cols
        </div>
        {f'<div class="sheet-nav">{nav_html}</div>' if nav_html else ''}
    </header>

    <main>
        {"".join(sections)}
    </main>

    <footer>
        <p>Generated by Excel Analyzer for Claude Code</p>
    </footer>
</body>
</html>"""

        path = self.output_dir / "sheets" / self._sheet_filename(name)
        path.write_text(html, encoding="utf-8")

    def _build_screenshot_section(self, screenshots) -> str:
        """Build screenshots section for a sheet."""
        import re

        views = {}
        birdseye_zoom = None
        for s in screenshots:
            filename = s.path.name
            if "_100." in filename:
                views["normal"] = filename
            else:
                # Any other zoom level is bird's eye
                views["birdseye"] = filename
                # Extract zoom level from filename (e.g., "sheet_25.png" -> 25)
                match = re.search(r'_(\d+)\.png$', filename)
                if match:
                    birdseye_zoom = match.group(1)

        imgs = ""
        if views.get("normal"):
            imgs += f'''
            <div class="screenshot-view">
                <span class="zoom-label">100% (Normal)</span>
                <a href="../screenshots/{views["normal"]}" target="_blank">
                    <img src="../screenshots/{views["normal"]}" alt="Normal View" loading="lazy" />
                </a>
            </div>'''
        if views.get("birdseye"):
            zoom_label = f"{birdseye_zoom}% (Fit All)" if birdseye_zoom else "Bird's Eye"
            imgs += f'''
            <div class="screenshot-view">
                <span class="zoom-label">{zoom_label}</span>
                <a href="../screenshots/{views["birdseye"]}" target="_blank">
                    <img src="../screenshots/{views["birdseye"]}" alt="Bird's Eye View" loading="lazy" />
                </a>
            </div>'''

        return f'''
        <section id="screenshots" class="content-section">
            <h2>Screenshots</h2>
            <div class="screenshot-views">{imgs}</div>
        </section>'''

    def _build_charts_section(self, charts, sheet_name: str = "") -> str:
        """Build charts section for a sheet."""
        cards = ""
        for c in charts:
            title_html = f"<p><strong>Title:</strong> {self._escape(c.title)}</p>" if c.title else ""
            data_html = f"<p><strong>Data:</strong> <code>{self._escape(c.data_range)}</code></p>" if c.data_range else ""

            # Check for chart image
            chart_img = ""
            if sheet_name:
                # Look for chart image in screenshots/charts/
                safe_sheet = self._sanitize_for_path(sheet_name)
                safe_chart = self._sanitize_for_path(c.name)
                img_path = self.output_dir / "screenshots" / "charts" / f"{safe_sheet}_{safe_chart}.png"
                if img_path.exists():
                    rel_path = f"../screenshots/charts/{safe_sheet}_{safe_chart}.png"
                    chart_img = f'''
                    <div class="chart-image">
                        <a href="{rel_path}" target="_blank">
                            <img src="{rel_path}" alt="{self._escape(c.name)}" loading="lazy" />
                        </a>
                    </div>'''

            cards += f'''
            <div class="item-card chart-card">
                <div class="item-header">
                    <span class="item-type">{self._escape(c.chart_type)}</span>
                    <span class="item-name">{self._escape(c.name)}</span>
                </div>
                {chart_img}
                <div class="item-details">
                    {title_html}
                    {data_html}
                </div>
            </div>'''

        return f'''
        <section id="charts" class="content-section">
            <h2>Charts ({len(charts)})</h2>
            <div class="item-grid">{cards}</div>
        </section>'''

    def _sanitize_for_path(self, name: str) -> str:
        """Sanitize a string for use in file paths."""
        import re
        result = re.sub(r'[<>:"/\\|?*]', '_', name)
        if len(result) > 100:
            result = result[:100]
        return result

    def _build_pivots_section(self, pivots) -> str:
        """Build pivot tables section for a sheet."""
        cards = ""
        for p in pivots:
            rows = f"<p><strong>Rows:</strong> {', '.join(p.row_fields)}</p>" if p.row_fields else ""
            cols = f"<p><strong>Columns:</strong> {', '.join(p.column_fields)}</p>" if p.column_fields else ""
            vals = f"<p><strong>Values:</strong> {', '.join(p.data_fields)}</p>" if p.data_fields else ""
            source = f"<p><strong>Source:</strong> <code>{self._escape(p.source_range)}</code></p>" if p.source_range else ""

            cards += f'''
            <div class="item-card pivot-card">
                <div class="item-header">
                    <span class="item-name">{self._escape(p.name)}</span>
                    <span class="item-location">{self._escape(p.location)}</span>
                </div>
                <div class="item-details">
                    {source}
                    {rows}
                    {cols}
                    {vals}
                </div>
            </div>'''

        return f'''
        <section id="pivots" class="content-section">
            <h2>Pivot Tables ({len(pivots)})</h2>
            <div class="item-grid">{cards}</div>
        </section>'''

    def _build_tables_section(self, tables) -> str:
        """Build structured tables section for a sheet."""
        cards = ""
        for t in tables:
            cols = ", ".join(t.columns[:6]) if t.columns else "-"
            if len(t.columns) > 6:
                cols += f" (+{len(t.columns) - 6} more)"

            cards += f'''
            <div class="item-card table-card">
                <div class="item-header">
                    <span class="item-name">{self._escape(t.display_name)}</span>
                </div>
                <div class="item-details">
                    <p><strong>Range:</strong> <code>{self._escape(t.range)}</code></p>
                    <p><strong>Columns:</strong> {self._escape(cols)}</p>
                </div>
            </div>'''

        return f'''
        <section id="tables" class="content-section">
            <h2>Structured Tables ({len(tables)})</h2>
            <div class="item-grid">{cards}</div>
        </section>'''

    def _build_formulas_section(self, formulas) -> str:
        """Build formulas section for a sheet."""
        # Filter out empty formulas
        formulas = [f for f in formulas if f.formula_clean.strip() not in ("=", "")]

        if not formulas:
            return ""

        # Group by category
        by_cat = defaultdict(list)
        for f in formulas:
            by_cat[f.category.value].append(f)

        content = ""
        for cat, cat_formulas in sorted(by_cat.items(), key=lambda x: -len(x[1])):
            rows = ""
            for f in cat_formulas[:30]:  # Limit per category
                formula_escaped = self._escape(f.formula_clean)
                # Show preview with expand button for long formulas
                if len(f.formula_clean) > 50:
                    preview = self._escape(f.formula_clean[:50]) + "..."
                    rows += f'''
                <tr class="formula-row">
                    <td class="cell-ref">{f.location.cell}</td>
                    <td class="formula-cell">
                        <code class="formula-preview">{preview}</code>
                        <div class="formula-full collapsed"><code>{formula_escaped}</code></div>
                        <button class="expand-btn" onclick="toggleFormula(this)">Show full</button>
                    </td>
                </tr>'''
                else:
                    rows += f'''
                <tr class="formula-row">
                    <td class="cell-ref">{f.location.cell}</td>
                    <td class="formula-cell"><code>{formula_escaped}</code></td>
                </tr>'''

            more = f"<p class='more-note'>...and {len(cat_formulas) - 30} more</p>" if len(cat_formulas) > 30 else ""

            content += f'''
            <div class="formula-category">
                <h4>{cat} ({len(cat_formulas)})</h4>
                <table class="data-table formula-table">
                    <thead><tr><th>Cell</th><th>Formula</th></tr></thead>
                    <tbody>{rows}</tbody>
                </table>
                {more}
            </div>'''

        return f'''
        <section id="formulas" class="content-section">
            <h2>Formulas ({len(formulas)})</h2>
            {content}
            <script>
            function toggleFormula(btn) {{
                const row = btn.closest('.formula-cell');
                const preview = row.querySelector('.formula-preview');
                const full = row.querySelector('.formula-full');
                if (full.classList.contains('collapsed')) {{
                    full.classList.remove('collapsed');
                    if (preview) preview.style.display = 'none';
                    btn.textContent = 'Show less';
                }} else {{
                    full.classList.add('collapsed');
                    if (preview) preview.style.display = 'inline';
                    btn.textContent = 'Show full';
                }}
            }}
            </script>
        </section>'''

    def _build_cf_section(self, cfs) -> str:
        """Build conditional formatting section."""
        rows = ""
        for cf in cfs[:30]:
            rows += f'''
            <tr>
                <td>{self._escape(cf.range.split("!")[-1] if "!" in cf.range else cf.range)}</td>
                <td>{cf.rule_type.value}</td>
                <td>{self._escape(cf.description)}</td>
            </tr>'''

        more = f"<p class='more-note'>...and {len(cfs) - 30} more rules</p>" if len(cfs) > 30 else ""

        return f'''
        <section id="cf" class="content-section">
            <h2>Conditional Formatting ({len(cfs)})</h2>
            <table class="data-table">
                <thead><tr><th>Range</th><th>Type</th><th>Description</th></tr></thead>
                <tbody>{rows}</tbody>
            </table>
            {more}
        </section>'''

    def _build_dv_section(self, dvs) -> str:
        """Build data validation section."""
        rows = ""
        for dv in dvs[:30]:
            formula = dv.formula1 or "-"
            if len(formula) > 40:
                formula = formula[:40] + "..."
            rows += f'''
            <tr>
                <td>{self._escape(dv.range.split("!")[-1] if "!" in dv.range else dv.range)}</td>
                <td>{dv.type}</td>
                <td><code>{self._escape(formula)}</code></td>
            </tr>'''

        more = f"<p class='more-note'>...and {len(dvs) - 30} more rules</p>" if len(dvs) > 30 else ""

        return f'''
        <section id="dv" class="content-section">
            <h2>Data Validation ({len(dvs)})</h2>
            <table class="data-table">
                <thead><tr><th>Range</th><th>Type</th><th>Formula/List</th></tr></thead>
                <tbody>{rows}</tbody>
            </table>
            {more}
        </section>'''

    def _build_comments_section(self, comments) -> str:
        """Build comments section."""
        items = ""
        for c in comments[:20]:
            text = c.text[:100] + "..." if len(c.text) > 100 else c.text
            items += f'''
            <div class="comment-item">
                <span class="comment-cell">{c.location.cell}</span>
                <span class="comment-author">{self._escape(c.author or "Unknown")}</span>
                <p class="comment-text">{self._escape(text)}</p>
            </div>'''

        more = f"<p class='more-note'>...and {len(comments) - 20} more comments</p>" if len(comments) > 20 else ""

        return f'''
        <section id="comments" class="content-section">
            <h2>Comments ({len(comments)})</h2>
            <div class="comments-list">{items}</div>
            {more}
        </section>'''

    def _build_controls_section(self, controls) -> str:
        """Build form controls section."""
        rows = ""
        for c in controls:
            macro = c.macro or "-"
            rows += f'''
            <tr>
                <td>{self._escape(c.name)}</td>
                <td>{c.control_type}</td>
                <td><code>{self._escape(macro)}</code></td>
            </tr>'''

        return f'''
        <section id="controls" class="content-section">
            <h2>Form Controls ({len(controls)})</h2>
            <table class="data-table">
                <thead><tr><th>Name</th><th>Type</th><th>Macro</th></tr></thead>
                <tbody>{rows}</tbody>
            </table>
        </section>'''

    def _build_errors_section(self, errors) -> str:
        """Build errors section."""
        rows = ""
        for e in errors[:30]:
            formula = e.formula or "-"
            if len(formula) > 50:
                formula = formula[:50] + "..."
            rows += f'''
            <tr>
                <td>{e.location.cell}</td>
                <td class="error-type">{e.error_type.value}</td>
                <td><code>{self._escape(formula)}</code></td>
            </tr>'''

        more = f"<p class='more-note'>...and {len(errors) - 30} more errors</p>" if len(errors) > 30 else ""

        return f'''
        <section id="errors" class="content-section">
            <h2>Errors ({len(errors)})</h2>
            <table class="data-table">
                <thead><tr><th>Cell</th><th>Error</th><th>Formula</th></tr></thead>
                <tbody>{rows}</tbody>
            </table>
            {more}
        </section>'''

    def _build_vba_refs_section(self, vba_refs) -> str:
        """Build VBA references section."""
        links = ""
        for module in sorted(vba_refs):
            links += f'<a href="../workbook/vba.html#{self._slug(module)}" class="vba-link">{self._escape(module)}</a>'

        return f'''
        <section id="vba" class="content-section">
            <h2>VBA Modules Used</h2>
            <p>This sheet references the following VBA modules:</p>
            <div class="vba-links">{links}</div>
        </section>'''

    def _generate_vba_page(self):
        """Generate the VBA modules workbook-wide page."""
        a = self.analysis

        modules_html = ""
        for m in a.vba_modules:
            # Get sheets that use this module
            using_sheets = self.vba_to_sheets.get(m.name, set())
            sheets_html = ""
            if using_sheets:
                sheet_links = ", ".join(
                    f'<a href="../sheets/{self._sheet_filename(s)}">{self._escape(s)}</a>'
                    for s in sorted(using_sheets)
                )
                sheets_html = f'<p><strong>Used by sheets:</strong> {sheet_links}</p>'

            # Syntax highlight
            try:
                highlighted = highlight(m.code, VbNetLexer(), HtmlFormatter(nowrap=True))
            except Exception:
                highlighted = self._escape(m.code)

            procs = ", ".join(m.procedures) if m.procedures else "None"

            modules_html += f'''
            <div id="{self._slug(m.name)}" class="vba-module">
                <div class="module-header" onclick="toggleModule(this)">
                    <h3>{self._escape(m.name)}</h3>
                    <span class="module-info">{m.module_type} · {m.line_count} lines · {len(m.procedures or [])} procedures</span>
                </div>
                <div class="module-content collapsed">
                    <p><strong>Procedures:</strong> {self._escape(procs)}</p>
                    {sheets_html}
                    <pre class="code-block"><code>{highlighted}</code></pre>
                </div>
            </div>'''

        html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>VBA Modules - {self._escape(a.file_name)}</title>
    <link rel="stylesheet" href="../styles.css">
</head>
<body>
    <nav class="breadcrumb">
        <a href="../index.html">← Back to Index</a>
    </nav>

    <header class="page-header">
        <h1>VBA Modules</h1>
        <p class="subtitle">{len(a.vba_modules)} modules in {self._escape(a.file_name)}</p>
    </header>

    <main>
        {modules_html}
    </main>

    <footer>
        <p>Generated by Excel Analyzer for Claude Code</p>
    </footer>

    <script>
        function toggleModule(header) {{
            header.nextElementSibling.classList.toggle('collapsed');
        }}
    </script>
</body>
</html>"""

        (self.output_dir / "workbook" / "vba.html").write_text(html, encoding="utf-8")

    def _generate_power_query_page(self):
        """Generate the Power Query workbook-wide page."""
        a = self.analysis

        queries_html = ""
        for q in a.power_queries:
            try:
                lexer = get_lexer_by_name("text")
                highlighted = highlight(q.formula, lexer, HtmlFormatter(nowrap=True))
            except Exception:
                highlighted = self._escape(q.formula)

            queries_html += f'''
            <div class="pq-query">
                <div class="query-header" onclick="toggleModule(this)">
                    <h3>{self._escape(q.name)}</h3>
                </div>
                <div class="query-content collapsed">
                    {f"<p>{self._escape(q.description)}</p>" if q.description else ""}
                    <pre class="code-block"><code>{highlighted}</code></pre>
                </div>
            </div>'''

        html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Power Query - {self._escape(a.file_name)}</title>
    <link rel="stylesheet" href="../styles.css">
</head>
<body>
    <nav class="breadcrumb">
        <a href="../index.html">← Back to Index</a>
    </nav>

    <header class="page-header">
        <h1>Power Query</h1>
        <p class="subtitle">{len(a.power_queries)} queries in {self._escape(a.file_name)}</p>
    </header>

    <main>
        {queries_html}
    </main>

    <footer>
        <p>Generated by Excel Analyzer for Claude Code</p>
    </footer>

    <script>
        function toggleModule(header) {{
            header.nextElementSibling.classList.toggle('collapsed');
        }}
    </script>
</body>
</html>"""

        (self.output_dir / "workbook" / "power-query.html").write_text(html, encoding="utf-8")

    def _generate_connections_page(self):
        """Generate the connections workbook-wide page."""
        a = self.analysis

        content = ""

        if a.connections:
            conn_cards = ""
            for c in a.connections:
                # Connection string (truncated)
                conn_str = c.connection_string or "-"
                conn_str_display = conn_str[:80] + "..." if len(conn_str) > 80 else conn_str

                # Command type badge
                cmd_type_badge = ""
                if c.command_type:
                    badge_class = "dax" if c.is_dax else ""
                    cmd_type_badge = f'<span class="cmd-type-badge {badge_class}">{c.command_type}</span>'

                # DAX query section
                dax_section = ""
                if c.is_dax and c.dax_query:
                    dax_section = f'''
                    <div class="dax-query-section">
                        <h5>DAX Query</h5>
                        <pre class="code-block dax-code"><code>{self._escape(c.dax_query)}</code></pre>
                    </div>'''

                # Command text (if not DAX, show as SQL/command)
                cmd_section = ""
                if c.command_text and not c.is_dax:
                    cmd_section = f'''
                    <div class="command-section">
                        <h5>Command</h5>
                        <pre class="code-block"><code>{self._escape(c.command_text)}</code></pre>
                    </div>'''

                # Used by pivot caches
                pivot_section = ""
                if c.used_by_pivot_caches:
                    pivot_links = ", ".join(c.used_by_pivot_caches)
                    pivot_section = f'''
                    <div class="used-by-section">
                        <strong>Used by:</strong> {pivot_links}
                    </div>'''

                conn_cards += f'''
                <div class="connection-card">
                    <div class="connection-header">
                        <h4>{self._escape(c.name)}</h4>
                        <div class="connection-badges">
                            <span class="conn-type-badge">{c.connection_type}</span>
                            {cmd_type_badge}
                        </div>
                    </div>
                    <div class="connection-details">
                        {f'<p class="conn-desc">{self._escape(c.description)}</p>' if c.description else ''}
                        <p><strong>Connection:</strong> <code>{self._escape(conn_str_display)}</code></p>
                        {pivot_section}
                        {dax_section}
                        {cmd_section}
                    </div>
                </div>'''

            content += f'''
            <section class="content-section">
                <h2>Data Connections ({len(a.connections)})</h2>
                <div class="connections-grid">
                    {conn_cards}
                </div>
            </section>'''

        if a.external_refs:
            rows = ""
            for ref in a.external_refs:
                status = '<span class="badge error">Broken</span>' if ref.is_broken else ""
                source = ref.source_cell.address if ref.source_cell.cell else "-"
                rows += f'''
                <tr>
                    <td>{self._escape(source)}</td>
                    <td>{self._escape(ref.target_workbook)}</td>
                    <td>{self._escape(ref.target_sheet or "-")}</td>
                    <td>{status}</td>
                </tr>'''

            content += f'''
            <section class="content-section">
                <h2>External References ({len(a.external_refs)})</h2>
                <table class="data-table">
                    <thead><tr><th>Source</th><th>Workbook</th><th>Sheet</th><th>Status</th></tr></thead>
                    <tbody>{rows}</tbody>
                </table>
            </section>'''

        html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Connections - {self._escape(a.file_name)}</title>
    <link rel="stylesheet" href="../styles.css">
</head>
<body>
    <nav class="breadcrumb">
        <a href="../index.html">← Back to Index</a>
    </nav>

    <header class="page-header">
        <h1>Connections & External References</h1>
        <p class="subtitle">{self._escape(a.file_name)}</p>
    </header>

    <main>
        {content}
    </main>

    <footer>
        <p>Generated by Excel Analyzer for Claude Code</p>
    </footer>
</body>
</html>"""

        (self.output_dir / "workbook" / "connections.html").write_text(html, encoding="utf-8")

    def _generate_named_ranges_page(self):
        """Generate the named ranges workbook-wide page."""
        a = self.analysis

        lambdas = [n for n in a.named_ranges if n.is_lambda]
        regular = [n for n in a.named_ranges if not n.is_lambda]

        content = ""

        if lambdas:
            items = ""
            for n in lambdas:
                items += f'''
                <div class="lambda-def">
                    <h4>{self._escape(n.name)}</h4>
                    <pre><code>{self._escape(n.value)}</code></pre>
                </div>'''

            content += f'''
            <section class="content-section">
                <h2>LAMBDA Functions ({len(lambdas)})</h2>
                {items}
            </section>'''

        if regular:
            rows = ""
            for n in regular[:100]:
                scope = n.scope or "Global"
                value = n.value[:60] + "..." if len(n.value) > 60 else n.value
                rows += f'''
                <tr>
                    <td>{self._escape(n.name)}</td>
                    <td><code>{self._escape(value)}</code></td>
                    <td>{self._escape(scope)}</td>
                </tr>'''

            more = f"<p class='more-note'>...and {len(regular) - 100} more</p>" if len(regular) > 100 else ""

            content += f'''
            <section class="content-section">
                <h2>Named Ranges ({len(regular)})</h2>
                <table class="data-table">
                    <thead><tr><th>Name</th><th>Value</th><th>Scope</th></tr></thead>
                    <tbody>{rows}</tbody>
                </table>
                {more}
            </section>'''

        html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Named Ranges - {self._escape(a.file_name)}</title>
    <link rel="stylesheet" href="../styles.css">
</head>
<body>
    <nav class="breadcrumb">
        <a href="../index.html">← Back to Index</a>
    </nav>

    <header class="page-header">
        <h1>Named Ranges</h1>
        <p class="subtitle">{len(a.named_ranges)} definitions in {self._escape(a.file_name)}</p>
    </header>

    <main>
        {content}
    </main>

    <footer>
        <p>Generated by Excel Analyzer for Claude Code</p>
    </footer>
</body>
</html>"""

        (self.output_dir / "workbook" / "named-ranges.html").write_text(html, encoding="utf-8")

    def _generate_warnings_block(self) -> str:
        """Generate warnings/errors block if any."""
        a = self.analysis
        if not a.errors and not a.warnings:
            return ""

        items = ""
        for e in a.errors:
            items += f'<li class="error">{self._escape(e.extractor)}: {self._escape(e.message)}</li>'
        for w in a.warnings:
            items += f'<li class="warning">{self._escape(w.extractor)}: {self._escape(w.message)}</li>'

        return f'''
        <div class="warnings-block">
            <h4>Extraction Notes</h4>
            <ul>{items}</ul>
        </div>'''

    def _get_sheet_feature_badges(self, sheet: SheetInfo) -> str:
        """Get HTML badges for sheet features."""
        badges = []
        if sheet.has_formulas:
            badges.append('<span class="badge">Formulas</span>')
        if sheet.has_charts:
            badges.append('<span class="badge">Charts</span>')
        if sheet.has_pivots:
            badges.append('<span class="badge">Pivots</span>')
        if sheet.has_tables:
            badges.append('<span class="badge">Tables</span>')
        if sheet.has_conditional_formatting:
            badges.append('<span class="badge">CF</span>')
        if sheet.has_data_validation:
            badges.append('<span class="badge">DV</span>')
        return " ".join(badges) if badges else ""

    def _sheet_filename(self, name: str) -> str:
        """Convert sheet name to safe filename."""
        # Replace problematic characters
        safe = re.sub(r'[<>:"/\\|?*]', '_', name)
        safe = re.sub(r'\s+', '-', safe)
        return f"{safe}.html"

    def _slug(self, text: str) -> str:
        """Convert text to URL-safe slug."""
        slug = re.sub(r'[^a-zA-Z0-9]+', '-', text.lower())
        return slug.strip('-')

    def _escape(self, text: str) -> str:
        """Escape HTML special characters."""
        if not text:
            return ""
        return (
            str(text)
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;")
            .replace("'", "&#39;")
        )

    def _format_size(self, size_bytes: int) -> str:
        """Format file size in human-readable form."""
        for unit in ["B", "KB", "MB", "GB"]:
            if size_bytes < 1024:
                return f"{size_bytes:.1f} {unit}"
            size_bytes /= 1024
        return f"{size_bytes:.1f} TB"

    def _get_styles(self) -> str:
        """Get the shared CSS styles."""
        return """
:root {
    --bg: #f8f9fa;
    --card-bg: #ffffff;
    --text: #212529;
    --text-muted: #6c757d;
    --border: #dee2e6;
    --primary: #0d6efd;
    --success: #198754;
    --warning: #ffc107;
    --error: #dc3545;
}

* {
    box-sizing: border-box;
    margin: 0;
    padding: 0;
}

body {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    background: var(--bg);
    color: var(--text);
    line-height: 1.6;
    padding: 2rem;
    max-width: 1400px;
    margin: 0 auto;
}

.breadcrumb {
    margin-bottom: 1rem;
}

.breadcrumb a {
    color: var(--primary);
    text-decoration: none;
}

.breadcrumb a:hover {
    text-decoration: underline;
}

.page-header {
    margin-bottom: 2rem;
    padding-bottom: 1rem;
    border-bottom: 2px solid var(--border);
}

.page-header h1 {
    font-size: 2rem;
    margin-bottom: 0.25rem;
}

.subtitle {
    color: var(--text-muted);
    font-size: 1.1rem;
}

.meta {
    color: var(--text-muted);
    font-size: 0.9rem;
}

/* Sheet header */
.sheet-header .sheet-title {
    display: flex;
    align-items: center;
    gap: 0.75rem;
}

.sheet-meta {
    color: var(--text-muted);
    margin-top: 0.5rem;
}

.sheet-nav {
    margin-top: 1rem;
    padding-top: 1rem;
    border-top: 1px solid var(--border);
}

.sheet-nav a {
    color: var(--primary);
    text-decoration: none;
}

.sheet-nav a:hover {
    text-decoration: underline;
}

/* Stats grid */
.stats-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(120px, 1fr));
    gap: 1rem;
    margin-bottom: 1.5rem;
}

.stat-card {
    background: var(--card-bg);
    border-radius: 8px;
    padding: 1rem;
    text-align: center;
    box-shadow: 0 1px 3px rgba(0,0,0,0.1);
}

.stat-value {
    font-size: 2rem;
    font-weight: bold;
    color: var(--primary);
}

.stat-label {
    color: var(--text-muted);
    font-size: 0.9rem;
}

/* Sheet groups */
.sheet-group {
    margin-bottom: 2rem;
}

.sheet-group h3 {
    margin-bottom: 1rem;
    color: var(--text-muted);
    font-size: 1rem;
    text-transform: uppercase;
    letter-spacing: 0.05em;
}

.sheet-cards {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(280px, 1fr));
    gap: 1rem;
}

.sheet-card {
    display: block;
    background: var(--card-bg);
    border-radius: 8px;
    padding: 1rem;
    text-decoration: none;
    color: var(--text);
    box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    transition: transform 0.2s, box-shadow 0.2s;
}

.sheet-card:hover {
    transform: translateY(-2px);
    box-shadow: 0 4px 12px rgba(0,0,0,0.15);
}

.sheet-card-header {
    display: flex;
    align-items: center;
    gap: 0.5rem;
    margin-bottom: 0.5rem;
}

.sheet-name {
    font-weight: 600;
    font-size: 1.1rem;
}

.sheet-card-meta {
    color: var(--text-muted);
    font-size: 0.85rem;
    margin-bottom: 0.5rem;
}

.sheet-card-features {
    display: flex;
    flex-wrap: wrap;
    gap: 0.25rem;
}

/* Workbook links */
.workbook-links {
    display: flex;
    flex-wrap: wrap;
    gap: 1rem;
}

.workbook-link {
    display: flex;
    align-items: center;
    gap: 0.5rem;
    background: var(--card-bg);
    border-radius: 8px;
    padding: 1rem 1.5rem;
    text-decoration: none;
    color: var(--text);
    box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    transition: transform 0.2s;
}

.workbook-link:hover {
    transform: translateY(-2px);
}

.workbook-link .icon {
    font-size: 1.5rem;
}

/* Badges */
.badge {
    display: inline-block;
    padding: 0.15rem 0.4rem;
    border-radius: 4px;
    font-size: 0.7rem;
    font-weight: 500;
    background: var(--bg);
    color: var(--text-muted);
}

.badge.error {
    background: #ffeef0;
    color: var(--error);
}

.visibility-badge {
    padding: 0.2rem 0.5rem;
    border-radius: 4px;
    font-size: 0.75rem;
    font-weight: 500;
}

.visibility-badge.hidden {
    background: #fff3cd;
    color: #856404;
}

.visibility-badge.very-hidden {
    background: #ffeef0;
    color: var(--error);
}

/* Color dot */
.color-dot {
    display: inline-block;
    width: 12px;
    height: 12px;
    border-radius: 50%;
}

.color-dot.large {
    width: 16px;
    height: 16px;
}

/* Content sections */
.content-section {
    background: var(--card-bg);
    border-radius: 8px;
    padding: 1.5rem;
    margin-bottom: 1.5rem;
    box-shadow: 0 1px 3px rgba(0,0,0,0.1);
}

.content-section h2 {
    margin-bottom: 1rem;
    padding-bottom: 0.5rem;
    border-bottom: 1px solid var(--border);
}

/* Item cards (charts, pivots, tables) */
.item-grid {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(300px, 1fr));
    gap: 1rem;
}

.item-card {
    background: var(--bg);
    border-radius: 8px;
    overflow: hidden;
}

.item-header {
    padding: 0.75rem 1rem;
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    color: white;
    display: flex;
    justify-content: space-between;
    align-items: center;
    flex-wrap: wrap;
    gap: 0.5rem;
}

.pivot-card .item-header {
    background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);
}

.table-card .item-header {
    background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
}

.item-type {
    background: rgba(255,255,255,0.2);
    padding: 0.2rem 0.5rem;
    border-radius: 4px;
    font-size: 0.75rem;
}

.item-name {
    font-weight: 600;
}

.item-location {
    font-size: 0.8rem;
    opacity: 0.9;
}

.item-details {
    padding: 1rem;
}

.item-details p {
    margin: 0.25rem 0;
    font-size: 0.9rem;
}

/* Chart images */
.chart-image {
    padding: 0.5rem;
    background: var(--bg);
    border-bottom: 1px solid var(--border);
}

.chart-image img {
    max-width: 100%;
    height: auto;
    display: block;
    border-radius: 4px;
}

/* Data tables */
.data-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 0.9rem;
}

.data-table th,
.data-table td {
    padding: 0.75rem;
    text-align: left;
    border-bottom: 1px solid var(--border);
}

.data-table th {
    background: var(--bg);
    font-weight: 600;
}

.data-table tr:hover {
    background: var(--bg);
}

/* Formula categories */
.formula-category {
    margin-bottom: 1.5rem;
}

.formula-category h4 {
    margin-bottom: 0.5rem;
    color: var(--text-muted);
}

/* Formula table */
.formula-table .cell-ref {
    width: 60px;
    white-space: nowrap;
    font-weight: 500;
}

.formula-cell {
    position: relative;
}

.formula-cell code {
    word-break: break-all;
    white-space: pre-wrap;
}

.formula-full {
    margin-top: 0.5rem;
    padding: 0.5rem;
    background: var(--card-bg);
    border-radius: 4px;
    border: 1px solid var(--border);
}

.formula-full.collapsed {
    display: none;
}

.expand-btn {
    margin-top: 0.25rem;
    padding: 0.2rem 0.5rem;
    font-size: 0.75rem;
    background: var(--bg);
    border: 1px solid var(--border);
    border-radius: 3px;
    cursor: pointer;
    color: var(--primary);
}

.expand-btn:hover {
    background: var(--border);
}

/* Screenshots */
.screenshot-views {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(400px, 1fr));
    gap: 1rem;
}

.screenshot-view {
    background: var(--bg);
    border-radius: 6px;
    overflow: hidden;
}

.zoom-label {
    display: block;
    padding: 0.5rem;
    font-size: 0.85rem;
    color: var(--text-muted);
    text-align: center;
    background: var(--card-bg);
    border-bottom: 1px solid var(--border);
}

.screenshot-view img {
    width: 100%;
    height: auto;
    display: block;
}

/* Comments */
.comments-list {
    display: flex;
    flex-direction: column;
    gap: 0.75rem;
}

.comment-item {
    background: var(--bg);
    padding: 0.75rem;
    border-radius: 6px;
}

.comment-cell {
    font-weight: 600;
    margin-right: 0.5rem;
}

.comment-author {
    color: var(--text-muted);
    font-size: 0.85rem;
}

.comment-text {
    margin-top: 0.25rem;
    font-size: 0.9rem;
}

/* VBA/PQ modules */
.vba-module, .pq-query {
    margin-bottom: 1rem;
}

.module-header, .query-header {
    cursor: pointer;
    padding: 1rem;
    background: var(--bg);
    border-radius: 6px;
    display: flex;
    justify-content: space-between;
    align-items: center;
}

.module-header:hover, .query-header:hover {
    background: var(--border);
}

.module-header h3, .query-header h3 {
    margin: 0;
    font-size: 1rem;
}

.module-info {
    color: var(--text-muted);
    font-size: 0.85rem;
}

.module-content, .query-content {
    padding: 1rem;
}

.collapsed {
    display: none;
}

/* Code blocks */
code {
    font-family: "SF Mono", Monaco, Consolas, monospace;
    font-size: 0.85em;
    background: var(--bg);
    padding: 0.1rem 0.3rem;
    border-radius: 3px;
}

.code-block {
    background: #1e1e1e;
    color: #d4d4d4;
    padding: 1rem;
    border-radius: 6px;
    overflow-x: auto;
    font-size: 0.85rem;
    line-height: 1.5;
}

.code-block code {
    background: none;
    padding: 0;
}

/* Lambda definitions */
.lambda-def {
    background: var(--bg);
    padding: 1rem;
    border-radius: 6px;
    margin-bottom: 1rem;
}

.lambda-def h4 {
    margin-bottom: 0.5rem;
    color: var(--primary);
}

/* VBA links */
.vba-links {
    display: flex;
    flex-wrap: wrap;
    gap: 0.5rem;
}

.vba-link {
    display: inline-block;
    padding: 0.5rem 1rem;
    background: var(--bg);
    border-radius: 6px;
    text-decoration: none;
    color: var(--primary);
}

.vba-link:hover {
    background: var(--border);
}

/* Warnings */
.warnings-block {
    background: #fff3cd;
    border: 1px solid #ffc107;
    border-radius: 6px;
    padding: 1rem;
    margin-top: 1rem;
}

.warnings-block h4 {
    margin-bottom: 0.5rem;
}

.warnings-block ul {
    list-style: none;
}

.warnings-block li.error {
    color: var(--error);
}

.warnings-block li.warning {
    color: #856404;
}

/* Misc */
.error-type {
    color: var(--error);
    font-weight: 500;
}

.more-note {
    color: var(--text-muted);
    font-size: 0.9rem;
    margin-top: 0.5rem;
}

.empty-state {
    color: var(--text-muted);
    text-align: center;
    padding: 2rem;
}

footer {
    margin-top: 2rem;
    padding-top: 1rem;
    border-top: 1px solid var(--border);
    text-align: center;
    color: var(--text-muted);
    font-size: 0.9rem;
}

/* Summary/sheets/workbook sections on index */
.summary-section, .sheets-section, .workbook-section {
    background: var(--card-bg);
    border-radius: 8px;
    padding: 1.5rem;
    margin-bottom: 1.5rem;
    box-shadow: 0 1px 3px rgba(0,0,0,0.1);
}

.summary-section h2, .sheets-section h2, .workbook-section h2 {
    margin-bottom: 1rem;
    padding-bottom: 0.5rem;
    border-bottom: 1px solid var(--border);
}

/* Connection cards */
.connections-grid {
    display: flex;
    flex-direction: column;
    gap: 1rem;
}

.connection-card {
    background: var(--bg);
    border-radius: 8px;
    overflow: hidden;
    border: 1px solid var(--border);
}

.connection-header {
    padding: 1rem;
    background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);
    color: white;
    display: flex;
    justify-content: space-between;
    align-items: center;
    flex-wrap: wrap;
    gap: 0.5rem;
}

.connection-header h4 {
    margin: 0;
    font-size: 1.1rem;
}

.connection-badges {
    display: flex;
    gap: 0.5rem;
}

.conn-type-badge, .cmd-type-badge {
    background: rgba(255,255,255,0.2);
    padding: 0.25rem 0.5rem;
    border-radius: 4px;
    font-size: 0.75rem;
}

.cmd-type-badge.dax {
    background: rgba(255,215,0,0.3);
    color: #fff;
    font-weight: 600;
}

.connection-details {
    padding: 1rem;
}

.connection-details p {
    margin: 0.5rem 0;
    font-size: 0.9rem;
}

.conn-desc {
    color: var(--text-muted);
    font-style: italic;
}

.used-by-section {
    margin: 0.75rem 0;
    padding: 0.5rem;
    background: var(--card-bg);
    border-radius: 4px;
    font-size: 0.9rem;
}

.dax-query-section, .command-section {
    margin-top: 1rem;
}

.dax-query-section h5, .command-section h5 {
    margin: 0 0 0.5rem 0;
    font-size: 0.9rem;
    color: var(--text-muted);
}

.dax-code {
    background: #1a1a2e;
}
"""
