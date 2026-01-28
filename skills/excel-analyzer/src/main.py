#!/usr/bin/env python3
"""Excel Workbook Analyzer - Main entry point."""

from __future__ import annotations

import argparse
import asyncio
import sys
import tempfile
from pathlib import Path
from urllib.parse import urlparse

from openpyxl import load_workbook

from .models import WorkbookAnalysis, ExtractionError, ExtractionWarning
from .auth import SessionStore, SSOHandler
from .download import SharePointDownloader
from .extractors import (
    SheetExtractor,
    FormulaExtractor,
    NamedRangeExtractor,
    ConditionalFormatExtractor,
    DataValidationExtractor,
    PivotTableExtractor,
    ChartExtractor,
    TableExtractor,
    FilterExtractor,
    VBAExtractor,
    PowerQueryExtractor,
    ControlExtractor,
    ConnectionExtractor,
    CommentExtractor,
    HyperlinkExtractor,
    ProtectionExtractor,
    PrintSettingsExtractor,
    ErrorExtractor,
    DAXDetector,
)
from .screenshots import ExcelOnlineScreenshotter
from .reports import HTMLReportBuilder, MarkdownReportBuilder


def is_url(path: str) -> bool:
    """Check if the input is a URL."""
    try:
        result = urlparse(path)
        return result.scheme in ("http", "https")
    except Exception:
        return False


def is_sharepoint_url(url: str) -> bool:
    """Check if URL is a SharePoint/OneDrive URL."""
    sharepoint_domains = [
        "sharepoint.com",
        "sharepoint.us",
        "onedrive.com",
        "office.com",
    ]
    try:
        parsed = urlparse(url)
        return any(domain in parsed.netloc.lower() for domain in sharepoint_domains)
    except Exception:
        return False


async def download_file(url: str, session_store: SessionStore) -> Path:
    """Download file from SharePoint."""
    downloader = SharePointDownloader(session_store)
    return await downloader.download(url)


async def capture_screenshots(
    sharepoint_url: str,
    sheets: list,
    output_dir: Path,
    session_store: SessionStore,
) -> list:
    """Capture screenshots via Excel Online."""
    from playwright.async_api import async_playwright

    sso_handler = SSOHandler(session_store)

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        try:
            context = await sso_handler.get_authenticated_context(sharepoint_url, browser)
            screenshotter = ExcelOnlineScreenshotter(output_dir / "screenshots")
            screenshots = await screenshotter.capture_all_sheets(context, sharepoint_url, sheets)
            await context.close()
            return screenshots
        finally:
            await browser.close()


def run_extractor(extractor_class, workbook, file_path, name: str):
    """Run a single extractor and handle errors."""
    try:
        extractor = extractor_class(workbook, file_path)
        return extractor.extract(), None
    except Exception as e:
        return None, ExtractionError(extractor=name, message=str(e))


def analyze_workbook(file_path: Path) -> WorkbookAnalysis:
    """Run all extractors on a workbook."""
    print(f"Analyzing: {file_path.name}")

    # Load workbook
    # Use data_only=False to get formulas instead of cached values
    workbook = load_workbook(str(file_path), data_only=False, read_only=False)

    # Initialize analysis
    analysis = WorkbookAnalysis(
        file_path=file_path,
        file_name=file_path.name,
        file_size=file_path.stat().st_size,
        is_macro_enabled=file_path.suffix.lower() in (".xlsm", ".xlsb", ".xltm", ".xlam"),
    )

    # Run extractors
    extractors = [
        (SheetExtractor, "sheets"),
        (FormulaExtractor, "formulas"),
        (NamedRangeExtractor, "named_ranges"),
        (ConditionalFormatExtractor, "conditional_formats"),
        (DataValidationExtractor, "data_validations"),
        (PivotTableExtractor, "pivot_tables"),
        (ChartExtractor, "charts"),
        (TableExtractor, "tables"),
        (FilterExtractor, "auto_filters"),
        (CommentExtractor, "comments"),
        (HyperlinkExtractor, "hyperlinks"),
        (ProtectionExtractor, "protection"),
        (PrintSettingsExtractor, "print_settings"),
        (ErrorExtractor, "error_cells"),
    ]

    for extractor_class, attr_name in extractors:
        print(f"  Extracting: {attr_name}...")
        result, error = run_extractor(extractor_class, workbook, file_path, attr_name)
        if error:
            analysis.errors.append(error)
        elif result is not None:
            setattr(analysis, attr_name, result)

    # VBA extraction (only for macro-enabled files)
    if analysis.is_macro_enabled:
        print("  Extracting: VBA modules...")
        result, error = run_extractor(VBAExtractor, workbook, file_path, "vba")
        if error:
            analysis.errors.append(error)
        elif result:
            analysis.vba_modules = result
            # Try to get project name
            try:
                vba_extractor = VBAExtractor(workbook, file_path)
                analysis.vba_project_name = vba_extractor.get_vba_project_name()
            except Exception:
                pass

    # Power Query extraction
    print("  Extracting: Power Query...")
    result, error = run_extractor(PowerQueryExtractor, workbook, file_path, "power_query")
    if error:
        analysis.errors.append(error)
    elif result:
        analysis.power_queries = result

    # Form controls extraction
    print("  Extracting: Form controls...")
    result, error = run_extractor(ControlExtractor, workbook, file_path, "controls")
    if error:
        analysis.errors.append(error)
    elif result:
        analysis.controls = result

    # Data connections extraction
    print("  Extracting: Data connections...")
    try:
        conn_extractor = ConnectionExtractor(workbook, file_path)
        connections, external_refs = conn_extractor.extract()
        analysis.connections = connections
        analysis.external_refs = external_refs
    except Exception as e:
        analysis.errors.append(ExtractionError(extractor="connections", message=str(e)))

    # DAX detection
    print("  Detecting: DAX/Power Pivot...")
    try:
        dax_detector = DAXDetector(workbook, file_path)
        has_dax, note = dax_detector.extract()
        analysis.has_dax = has_dax
        analysis.dax_detection_note = note
        if has_dax:
            analysis.warnings.append(ExtractionWarning(
                extractor="dax",
                message="DAX/Power Pivot detected but cannot be fully extracted",
                details=note,
            ))
    except Exception as e:
        analysis.errors.append(ExtractionError(extractor="dax", message=str(e)))

    workbook.close()

    return analysis


def generate_reports(analysis: WorkbookAnalysis, output_dir: Path) -> None:
    """Generate HTML and Markdown reports."""
    print(f"Generating reports in: {output_dir}")

    # HTML report
    print("  Building HTML report...")
    html_builder = HTMLReportBuilder(analysis, output_dir)
    html_path = html_builder.build()
    print(f"    Created: {html_path}")

    # Markdown reports
    print("  Building Markdown reports...")
    md_builder = MarkdownReportBuilder(analysis, output_dir)
    md_builder.build()
    print(f"    Created: {output_dir}/README.md")


def print_summary(analysis: WorkbookAnalysis) -> None:
    """Print analysis summary to console."""
    print("\n" + "=" * 60)
    print(f"ANALYSIS COMPLETE: {analysis.file_name}")
    print("=" * 60)

    print(f"\nSheets: {len(analysis.sheets)}")
    for s in analysis.sheets[:10]:
        vis = f" [{s.visibility.value}]" if s.visibility.value != "visible" else ""
        print(f"  - {s.name}{vis}")
    if len(analysis.sheets) > 10:
        print(f"  ... and {len(analysis.sheets) - 10} more")

    print(f"\nFormulas: {len(analysis.formulas)}")
    print(f"Named Ranges: {len(analysis.named_ranges)}")
    print(f"  LAMBDA functions: {sum(1 for n in analysis.named_ranges if n.is_lambda)}")

    if analysis.tables:
        print(f"Tables: {len(analysis.tables)}")
    if analysis.pivot_tables:
        print(f"Pivot Tables: {len(analysis.pivot_tables)}")
    if analysis.charts:
        print(f"Charts: {len(analysis.charts)}")
    if analysis.vba_modules:
        print(f"VBA Modules: {len(analysis.vba_modules)}")
    if analysis.power_queries:
        print(f"Power Queries: {len(analysis.power_queries)}")
    if analysis.has_dax:
        print(f"DAX/Power Pivot: Detected ({analysis.dax_detection_note})")

    if analysis.error_cells:
        print(f"\nError Cells: {len(analysis.error_cells)}")
    if analysis.external_refs:
        print(f"External References: {len(analysis.external_refs)}")

    if analysis.errors:
        print(f"\nExtraction Errors: {len(analysis.errors)}")
        for e in analysis.errors[:5]:
            print(f"  - {e.extractor}: {e.message}")

    print()


async def main_async(args: argparse.Namespace) -> int:
    """Async main function."""
    input_path = args.input
    output_dir = Path(args.output) if args.output else None
    skip_screenshots = args.no_screenshots

    session_store = SessionStore()
    sharepoint_url = None
    file_path = None
    temp_dir = None

    try:
        # Handle input (URL or local file)
        if is_url(input_path):
            if not is_sharepoint_url(input_path):
                print(f"Error: URL must be a SharePoint/OneDrive URL")
                return 1

            sharepoint_url = input_path
            print(f"Downloading from SharePoint...")
            file_path = await download_file(input_path, session_store)
            temp_dir = file_path.parent
        else:
            file_path = Path(input_path)
            if not file_path.exists():
                print(f"Error: File not found: {file_path}")
                return 1

        # Determine output directory
        if output_dir is None:
            output_dir = Path.cwd() / f"{file_path.stem}_analysis"

        output_dir.mkdir(parents=True, exist_ok=True)

        # Run analysis
        analysis = analyze_workbook(file_path)

        # Capture screenshots (if SharePoint URL provided and not skipped)
        if sharepoint_url and not skip_screenshots:
            print("Capturing screenshots via Excel Online...")
            try:
                screenshots = await capture_screenshots(
                    sharepoint_url,
                    analysis.sheets,
                    output_dir,
                    session_store,
                )
                analysis.screenshots = screenshots
                print(f"  Captured {len(screenshots)} screenshots")
            except Exception as e:
                print(f"  Warning: Could not capture screenshots: {e}")
                analysis.warnings.append(ExtractionWarning(
                    extractor="screenshots",
                    message=f"Could not capture screenshots: {e}",
                ))

        # Generate reports
        generate_reports(analysis, output_dir)

        # Print summary
        print_summary(analysis)

        print(f"Reports saved to: {output_dir}")
        print(f"  - HTML: {output_dir}/report.html")
        print(f"  - Markdown: {output_dir}/README.md")

        return 0

    finally:
        # Cleanup temp directory
        if temp_dir and temp_dir.exists():
            import shutil
            try:
                shutil.rmtree(temp_dir)
            except Exception:
                pass


def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description="Analyze Excel workbooks from SharePoint or local files",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  analyze-excel https://company.sharepoint.com/:x:/r/sites/Finance/Reports.xlsx
  analyze-excel ~/Documents/workbook.xlsx
  analyze-excel workbook.xlsx -o ./analysis_output
        """,
    )

    parser.add_argument(
        "input",
        help="SharePoint URL or path to local Excel file (.xlsx, .xlsm)",
    )

    parser.add_argument(
        "-o", "--output",
        help="Output directory for reports (default: {filename}_analysis)",
    )

    parser.add_argument(
        "--no-screenshots",
        action="store_true",
        help="Skip screenshot capture (faster, works for local files)",
    )

    parser.add_argument(
        "--clear-session",
        action="store_true",
        help="Clear cached SharePoint session and re-authenticate",
    )

    args = parser.parse_args()

    # Handle session clearing
    if args.clear_session and is_url(args.input):
        session_store = SessionStore()
        from urllib.parse import urlparse
        domain = urlparse(args.input).netloc
        session_store.clear_session(domain)
        print(f"Cleared session for {domain}")

    # Run async main
    try:
        return asyncio.run(main_async(args))
    except KeyboardInterrupt:
        print("\nCancelled.")
        return 130


if __name__ == "__main__":
    sys.exit(main())
