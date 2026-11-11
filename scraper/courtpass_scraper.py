#!/usr/bin/env python3
"""
CourtPASS Scraper — NY Court of Appeals
Downloads all available PDFs from CourtPASS for cases within a date window
and produces an organized output directory + index files.

Requirements:
  pip3 install playwright
  playwright install chromium
  Chrome must be running with: --remote-debugging-port=9222

Usage:
  python3 courtpass_scraper.py
  python3 courtpass_scraper.py --start 04/01/2026 --end 05/20/2026
  python3 courtpass_scraper.py --start 04/01/2026 --end 05/20/2026 --output ./my_output
"""

import asyncio
import argparse
import json
import os
import re
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional

from playwright.async_api import async_playwright, Download, TimeoutError as PlaywrightTimeoutError

try:
    from docx import Document
    from docx.shared import Pt, Inches, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

BASE_URL = "https://courtpass.nycourts.gov"
SEARCH_URL = f"{BASE_URL}/Public_search.aspx"
DOCKET_URL = f"{BASE_URL}/Docket.aspx"
CDP_URL = "http://localhost:9222"

# Filing type substrings that indicate a jurisdictional inquiry was issued
JI_KEYWORDS = ["jurisdict", "500.10"]

# The start and end date fields for each way of searching CourtPASS
SEARCH_FIELDS = {
    "argument": ("#cphMain_dpStartAppealGrantedDate", "#cphMain_dpEndAppealGrantedDate"),
    "decision": ("#cphMain_dpStartDecisionDate", "#cphMain_dpEndDecisionDate"),
}

# Handling for slow or failing pages
MAX_ATTEMPTS = 3
RETRY_WAIT_SECONDS = 20
DELAY_BETWEEN_CASES_SECONDS = 3
CASE_TIME_LIMIT_SECONDS = 30 * 60    # a case taking longer than this is treated as stuck
SEARCH_TIME_LIMIT_SECONDS = 10 * 60
FILE_SAVE_LIMIT_SECONDS = 10 * 60    # one PDF finishing its download
PAGE_TIMEOUT_MS = 60000
DOWNLOAD_TIMEOUT_MS = 120000
PROGRESS_FILE = "progress.json"

# Sanitize a string for use as a directory/file name
def safe_name(s: str) -> str:
    return re.sub(r'[^\w\s\-]', '', s).strip().replace(' ', '_')[:80]


def case_folder_name(case: dict) -> str:
    """Folder for a case: its argument date, or decision date if that's all we have."""
    name = safe_name(case["title"])
    date = (case.get("argument_date") or case.get("decision_date") or "").replace("/", "-")
    return f"{date}_{name}" if date else name


async def wait_and_stabilize(page, timeout=30000):
    try:
        await page.wait_for_load_state("networkidle", timeout=timeout)
    except Exception:
        await page.wait_for_load_state("domcontentloaded", timeout=timeout)


async def click_and_wait(page, target, timeout=PAGE_TIMEOUT_MS):
    """Click something that reloads the page, and wait for the new page to arrive.
    Reading the page before then fails, or returns the old page's contents."""
    try:
        async with page.expect_navigation(timeout=timeout):
            if isinstance(target, str):
                await page.click(target)
            else:
                await target.click()
    except PlaywrightTimeoutError:
        pass  # the click didn't reload the page
    await wait_and_stabilize(page)


async def do_search(page, start_date: str, end_date: str) -> list[dict]:
    """Search by argument date range and collect all case entries across pages."""
    print(f"\n[search] Argument dates {start_date} – {end_date}")
    await page.goto(SEARCH_URL, wait_until="networkidle", timeout=PAGE_TIMEOUT_MS)

    await page.fill("#cphMain_dpStartAppealGrantedDate", start_date)
    await page.fill("#cphMain_dpEndAppealGrantedDate", end_date)
    await page.evaluate("document.querySelectorAll('.flatpickr-calendar').forEach(e => e.classList.remove('open'))")

    # Also select "All Cases" (civil + criminal)
    await page.click("#cphMain_RadioButtonList1_2")

    await click_and_wait(page, "#cphMain_btnFind")

    text = await page.inner_text("body")
    if "No Matching Cases" in text:
        print("  → No cases found for this date range.")
        return []

    cases = []
    page_num = 1
    while True:
        print(f"  → Collecting page {page_num}...")
        rows = await page.query_selector_all("table tbody tr")
        for row in rows:
            cells = await row.query_selector_all("td")
            if len(cells) < 3:
                continue
            select_btn = await cells[0].query_selector("input[value='Select']")
            if not select_btn:
                continue
            aria = await select_btn.get_attribute("aria-label") or ""
            title_cell = await cells[1].inner_text()
            date_cell = await cells[2].inner_text() if len(cells) > 2 else ""
            arg_num_cell = await cells[3].inner_text() if len(cells) > 3 else ""

            btn_name = await select_btn.get_attribute("name")
            cases.append({
                "title": title_cell.strip(),
                "search_type": "argument",
                "argument_date": date_cell.strip(),
                "decision_date": "",
                "argument_number": arg_num_cell.strip(),
                "select_btn_name": btn_name,
                "aria": aria,
            })

        # Check for next page
        next_link = None
        page_links = await page.query_selector_all("a[href*='__doPostBack']")
        for link in page_links:
            t = (await link.inner_text()).strip()
            if t == str(page_num + 1):
                next_link = link
                break

        if not next_link:
            break
        page_num += 1
        await click_and_wait(page, next_link)

    print(f"  → Found {len(cases)} cases.")
    return cases


async def do_decision_date_search(page, start_date: str, end_date: str) -> list[dict]:
    """Search by decision date range and collect all case entries."""
    print(f"\n[search] Decision dates {start_date} – {end_date}")
    await page.goto(SEARCH_URL, wait_until="networkidle", timeout=PAGE_TIMEOUT_MS)

    await page.fill("#cphMain_dpStartDecisionDate", start_date)
    await page.fill("#cphMain_dpEndDecisionDate", end_date)
    await page.evaluate("document.querySelectorAll('.flatpickr-calendar').forEach(e => e.classList.remove('open'))")
    await page.click("#cphMain_RadioButtonList1_2")  # All cases
    await click_and_wait(page, "#cphMain_btnFind")

    text = await page.inner_text("body")
    if "No Matching Cases" in text:
        print("  → No cases found.")
        return []

    cases = []
    page_num = 1
    while True:
        print(f"  → Collecting page {page_num}...")
        rows = await page.query_selector_all("table tbody tr")
        for row in rows:
            cells = await row.query_selector_all("td")
            if len(cells) < 2:
                continue
            select_btn = await cells[0].query_selector("input[value='Select']")
            if not select_btn:
                continue
            title_cell = await cells[1].inner_text()
            date_cell = await cells[2].inner_text() if len(cells) > 2 else ""
            btn_name = await select_btn.get_attribute("name")
            aria = await select_btn.get_attribute("aria-label") or ""
            labelled = re.search(r"Decision date (\d{2}/\d{2}/\d{4})", aria)
            cases.append({
                "title": title_cell.strip(),
                "search_type": "decision",
                "argument_date": "",
                "decision_date": labelled.group(1) if labelled else date_cell.strip(),
                "argument_number": "",
                "select_btn_name": btn_name,
                "aria": aria,
            })

        next_link = None
        for link in await page.query_selector_all("a[href*='__doPostBack']"):
            t = (await link.inner_text()).strip()
            if t == str(page_num + 1):
                next_link = link
                break
        if not next_link:
            break
        page_num += 1
        await click_and_wait(page, next_link)

    print(f"  → Found {len(cases)} cases.")
    return cases


def merge_search_results(arg_cases: list[dict], dec_cases: list[dict]) -> list[dict]:
    """Combine both searches into one entry per case title. A case found by both
    keeps its argument-search details and gains the decision date."""
    merged = {}
    for c in arg_cases + dec_cases:
        key = c["title"].lower().strip()
        if key not in merged:
            merged[key] = c
        elif c.get("decision_date") and not merged[key].get("decision_date"):
            merged[key]["decision_date"] = c["decision_date"]
    return list(merged.values())


async def run_case_search(page, search_type: str, start_date: str, end_date: str):
    """Run one CourtPASS search, using the date fields for that search type."""
    start_field, end_field = SEARCH_FIELDS[search_type]
    await page.goto(SEARCH_URL, wait_until="networkidle", timeout=PAGE_TIMEOUT_MS)
    await page.fill(start_field, start_date)
    await page.fill(end_field, end_date)
    await page.evaluate("document.querySelectorAll('.flatpickr-calendar').forEach(e => e.classList.remove('open'))")
    await page.click("#cphMain_RadioButtonList1_2")
    await click_and_wait(page, "#cphMain_btnFind")


async def find_case_on_results(page, case_title: str) -> bool:
    """Move through the result pages until the case's Select button is showing."""
    title_lower = case_title.lower().strip()
    page_num = 1
    while True:
        for btn in await page.query_selector_all("input[value='Select']"):
            aria = (await btn.get_attribute("aria-label") or "").lower()
            if title_lower in aria:
                return True
        next_link = None
        for link in await page.query_selector_all("a[href*='__doPostBack']"):
            if (await link.inner_text()).strip() == str(page_num + 1):
                next_link = link
                break
        if not next_link:
            return False
        page_num += 1
        await click_and_wait(page, next_link)


async def open_case_results(page, case: dict, start_date: str, end_date: str) -> bool:
    """Bring up the results page holding this case, using the search that found it.
    Tries the case's exact date first, then the full date range."""
    search_type = case.get("search_type") or "argument"
    exact = case.get("decision_date") if search_type == "decision" else case.get("argument_date")
    if exact:
        await run_case_search(page, search_type, exact, exact)
        if await find_case_on_results(page, case["title"]):
            return True
    await run_case_search(page, search_type, start_date, end_date)
    return await find_case_on_results(page, case["title"])


async def get_case_detail(page, case: dict) -> dict:
    """
    Navigate to a case detail page by finding the Select button whose aria-label
    matches the case title, then collect metadata and file list.
    """
    select_btns = await page.query_selector_all("input[value='Select']")
    target_btn = None
    case_title_lower = case["title"].lower().strip()
    for btn in select_btns:
        aria = (await btn.get_attribute("aria-label") or "").lower()
        # Match on case title contained in aria label
        if case_title_lower in aria:
            target_btn = btn
            break

    if not target_btn:
        return {"error": "Select button not found", **case}

    await click_and_wait(page, target_btn)

    # Parse case detail
    body_text = await page.inner_text("body")
    html = await page.content()

    # Extract full case title (first heading)
    full_title = ""
    title_el = await page.query_selector("h1, h2, .case-title")
    if title_el:
        full_title = (await title_el.inner_text()).strip()

    # Extract issues
    issues = ""
    issues_start = body_text.find("Issues")
    if issues_start >= 0:
        next_section = body_text.find("The following documents", issues_start)
        if next_section >= 0:
            issues = body_text[issues_start + len("Issues"):next_section].strip()

    # Extract files
    files = []
    file_rows = await page.query_selector_all("#cphMain_gvFiles tr, table tr")
    for row in file_rows:
        filename_el = await row.query_selector("span[id*='lbFileName']")
        download_btn = await row.query_selector("input[id*='bttnDownload']")
        if not filename_el or not download_btn:
            continue
        filename = (await filename_el.inner_text()).strip()
        btn_value = await download_btn.get_attribute("value") or ""
        btn_name = await download_btn.get_attribute("name") or ""
        aria = await download_btn.get_attribute("aria-label") or ""
        is_available = btn_value == "Download PDF"
        # Extract actual filename from aria label
        pdf_name = aria.replace("Download PDF: ", "").replace("Not Available: ", "").strip()
        files.append({
            "display_name": filename,
            "pdf_name": pdf_name,
            "available": is_available,
            "btn_name": btn_name,
        })

    return {
        **case,
        "full_title": full_title or case["title"],
        "issues": issues,
        "files": files,
        "detail_url": page.url,
    }


async def check_docket_for_ji(page, case_title: str) -> dict:
    """
    Search the CourtPASS docket by case title and return JI-related filing info.

    Returns a dict with:
      ji_found      bool   — True if any filing entry matches JI keywords
      ji_filings    list   — the matching filing row strings
      all_filings   list   — every filing row (for reference)
      apl_number    str    — the APL/CTQ/JCR number if found
    """
    # Build a short search query from the first two meaningful words of the title
    words = [w for w in re.split(r'\W+', case_title) if len(w) > 2 and w.lower() not in
             {"the", "and", "for", "llc", "inc", "ltd", "corp", "matter", "people", "state"}]
    query = " ".join(words[:3])

    try:
        await page.goto(DOCKET_URL, wait_until="networkidle", timeout=PAGE_TIMEOUT_MS)
        await page.fill("#tbPartyNames", query)
        await click_and_wait(page, "#cphMain_btnFind")

        # Pick the right case if multiple results
        select_btns = await page.query_selector_all("input[value='Select']")
        target = None
        title_lower = case_title.lower()
        for btn in select_btns:
            row = await btn.evaluate_handle("el => el.closest('tr')")
            row_text = (await (await row.get_property("innerText")).json_value()).lower()
            if any(w.lower() in row_text for w in words[:2]):
                target = btn
                break
        if not target and select_btns:
            target = select_btns[0]
        if not target:
            return {"ji_found": False, "ji_filings": [], "all_filings": [], "apl_number": ""}

        await click_and_wait(page, target)

        body = await page.inner_text("body")

        # Extract APL number
        apl_number = ""
        apl_match = re.search(r'(APL|CTQ|JCR)-\d{4}-\d+', body)
        if apl_match:
            apl_number = apl_match.group(0)

        # Extract filing rows from the FILINGS table
        all_filings = []
        rows = await page.query_selector_all("table tr")
        in_filings = False
        for row in rows:
            cells = await row.query_selector_all("td, th")
            if not cells:
                continue
            cell_texts = [(await c.inner_text()).strip() for c in cells]
            row_text = "\t".join(cell_texts)
            if "Type" in cell_texts and "Date Due" in cell_texts:
                in_filings = True
                continue
            if in_filings and any("ATTORNEY" in t.upper() for t in cell_texts):
                break
            if in_filings and any(cell_texts):
                all_filings.append(row_text)

        ji_filings = [f for f in all_filings
                      if any(kw in f.lower() for kw in JI_KEYWORDS)]

        return {
            "ji_found": len(ji_filings) > 0,
            "ji_filings": ji_filings,
            "all_filings": all_filings,
            "apl_number": apl_number,
        }

    except Exception as e:
        return {"ji_found": False, "ji_filings": [], "all_filings": [], "apl_number": "", "docket_error": str(e)}


async def download_file(page, context, btn_name: str, output_dir: Path, pdf_name: str) -> Optional[str]:
    """Click a download button and save the resulting file."""
    output_dir.mkdir(parents=True, exist_ok=True)

    # Check if file already exists (by pdf_name or any .pdf with same stem)
    stem = Path(pdf_name).stem
    existing = list(output_dir.glob(f"{stem}*"))
    if existing:
        print(f"      [skip] {pdf_name} already downloaded")
        return str(existing[0])

    async with page.expect_download(timeout=DOWNLOAD_TIMEOUT_MS) as dl_info:
        await page.evaluate(f"""
            (function() {{
                var btn = document.querySelector('input[name="{btn_name}"]');
                if (btn) btn.click();
            }})();
        """)

    download: Download = await dl_info.value
    # Use browser's suggested filename as the authoritative name
    filename = download.suggested_filename or pdf_name
    dest = output_dir / filename
    await asyncio.wait_for(download.save_as(str(dest)), FILE_SAVE_LIMIT_SECONDS)
    size = dest.stat().st_size
    print(f"      [saved] {filename} ({size:,} bytes)")
    return str(dest)


async def process_case(page, context, case_detail: dict, output_base: Path) -> dict:
    """Download all available PDFs for a case and return enriched metadata."""
    case_dir = output_base / case_folder_name(case_detail)
    case_dir.mkdir(parents=True, exist_ok=True)

    downloaded = []
    unavailable = []
    failed_downloads = []

    for f in case_detail.get("files", []):
        if f["available"]:
            try:
                path = await download_file(page, context, f["btn_name"], case_dir, f["pdf_name"])
            except (PlaywrightTimeoutError, asyncio.TimeoutError):
                # Keep the case and its other files; a rerun tries this file again
                print(f"      [failed] {f['pdf_name']}: download didn't start in time")
                failed_downloads.append(f["pdf_name"])
                continue
            if path:
                downloaded.append({"file": f["pdf_name"], "path": path})
        else:
            unavailable.append(f["pdf_name"])

    return {
        **case_detail,
        "output_dir": str(case_dir),
        "downloaded_files": downloaded,
        "unavailable_files": unavailable,
        "failed_downloads": failed_downloads,
    }


def write_case_issues_docx(case_detail: dict, case_dir: Path):
    """Write a single-case issues.docx into the case's folder."""
    if not DOCX_AVAILABLE:
        return
    issues = (case_detail.get("issues") or "").strip()
    if not issues:
        return

    doc = Document()
    sec = doc.sections[0]
    sec.page_width    = Inches(8.5)
    sec.page_height   = Inches(11)
    sec.left_margin   = Inches(1)
    sec.right_margin  = Inches(1)
    sec.top_margin    = Inches(1)
    sec.bottom_margin = Inches(1)

    doc.styles["Normal"].font.name = "Arial"
    doc.styles["Normal"].font.size = Pt(11)
    h1 = doc.styles["Heading 1"]
    h1.font.name = "Arial"
    h1.font.size = Pt(14)
    h1.font.bold = True
    h1.font.color.rgb = RGBColor(0x1F, 0x49, 0x6E)

    title = case_detail.get("title") or case_detail.get("full_title", "Unknown")
    ji = case_detail.get("ji_found", False)
    doc.add_heading(("⚠ [JI]  " if ji else "") + title, level=1)

    # Metadata
    meta_parts = []
    if case_detail.get("argument_date"):
        meta_parts.append(f"Argued: {case_detail['argument_date']}")
    if case_detail.get("decision_date"):
        meta_parts.append(f"Decided: {case_detail['decision_date']}")
    if case_detail.get("argument_number"):
        meta_parts.append(f"No. {case_detail['argument_number']}")
    if case_detail.get("apl_number"):
        meta_parts.append(case_detail["apl_number"])
    if ji:
        meta_parts.append("⚠ Jurisdictional Inquiry")

    if meta_parts:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(8)
        mr = p.add_run("  |  ".join(meta_parts))
        mr.font.size = Pt(9)
        mr.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
        mr.italic = True

    if ji:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(6)
        wr = p.add_run("⚠ Jurisdictional Inquiry Detected")
        wr.bold = True
        wr.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
        for entry in case_detail.get("ji_filings", []):
            bp = doc.add_paragraph(style="List Bullet")
            bp.add_run(entry).font.size = Pt(10)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.add_run("Issues Presented:").bold = True

    p = doc.add_paragraph(issues)
    p.paragraph_format.left_indent = Inches(0.25)
    p.paragraph_format.space_after = Pt(6)

    out_path = case_dir / "issues.docx"
    doc.save(str(out_path))


def write_issues_docx(results: list, output_base: Path, start_date: str, end_date: str):
    """Write a Word document containing only case names and issues."""
    if not DOCX_AVAILABLE:
        return

    doc = Document()

    sec = doc.sections[0]
    sec.page_width    = Inches(8.5)
    sec.page_height   = Inches(11)
    sec.left_margin   = Inches(1)
    sec.right_margin  = Inches(1)
    sec.top_margin    = Inches(1)
    sec.bottom_margin = Inches(1)

    doc.styles["Normal"].font.name = "Arial"
    doc.styles["Normal"].font.size = Pt(11)

    for h, size in [("Heading 1", 16), ("Heading 2", 13)]:
        s = doc.styles[h]
        s.font.name = "Arial"
        s.font.size = Pt(size)
        s.font.bold = True
        s.font.color.rgb = RGBColor(0x1F, 0x49, 0x6E)

    # Title
    doc.add_heading("NY Court of Appeals — Issues Presented", level=1)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    p.add_run("Date range: ").bold = True
    p.add_run(f"{start_date} – {end_date}")

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(12)
    p.add_run("Generated: ").bold = True
    p.add_run(datetime.now().strftime("%Y-%m-%d %H:%M"))

    cases_with_issues = [r for r in results if (r.get("issues") or "").strip()]

    for i, r in enumerate(cases_with_issues, 1):
        title = r.get("title") or r.get("full_title", "Unknown")
        ji = r.get("ji_found", False)

        h = doc.add_heading(f"{i}. {'[JI] ' if ji else ''}{title}", level=2)
        h.paragraph_format.space_before = Pt(14)
        h.paragraph_format.space_after  = Pt(2)

        # Metadata line
        meta_parts = []
        if r.get("argument_date"):
            meta_parts.append(f"Argued: {r['argument_date']}")
        if r.get("decision_date"):
            meta_parts.append(f"Decided: {r['decision_date']}")
        if r.get("argument_number"):
            meta_parts.append(f"No. {r['argument_number']}")
        if r.get("apl_number"):
            meta_parts.append(r["apl_number"])
        if ji:
            meta_parts.append("⚠ Jurisdictional Inquiry")

        if meta_parts:
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(4)
            mr = p.add_run("  |  ".join(meta_parts))
            mr.font.size = Pt(9)
            mr.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
            mr.italic = True

        # Issues text
        issues = r["issues"].strip()
        p = doc.add_paragraph(issues)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.left_indent = Inches(0.25)
        for run in p.runs:
            run.font.size = Pt(11)

    out_path = output_base / "issues.docx"
    doc.save(str(out_path))
    print(f"[index] Wrote {out_path}")


def write_docx(results: list, output_base: Path, start_date: str, end_date: str):
    """Write a formatted Word document index."""
    if not DOCX_AVAILABLE:
        print("[docx] python-docx not installed — skipping Word output. Run: pip3 install python-docx")
        return

    doc = Document()

    # Page setup: US Letter, 1" margins
    sec = doc.sections[0]
    sec.page_width    = Inches(8.5)
    sec.page_height   = Inches(11)
    sec.left_margin   = Inches(1)
    sec.right_margin  = Inches(1)
    sec.top_margin    = Inches(1)
    sec.bottom_margin = Inches(1)

    # Base font
    doc.styles["Normal"].font.name = "Arial"
    doc.styles["Normal"].font.size = Pt(11)

    # Heading colours
    for h, size in [("Heading 1", 16), ("Heading 2", 13)]:
        s = doc.styles[h]
        s.font.name = "Arial"
        s.font.size = Pt(size)
        s.font.bold = True
        s.font.color.rgb = RGBColor(0x1F, 0x49, 0x6E)

    def add_field(label, value):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(2)
        r = p.add_run(label + ": ")
        r.bold = True
        r.font.size = Pt(10)
        v = p.add_run(value)
        v.font.size = Pt(10)

    def add_bullet(text, gray=False):
        p = doc.add_paragraph(style="List Bullet")
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(1)
        r = p.add_run(text)
        r.font.name = "Arial"
        r.font.size = Pt(9)
        if gray:
            r.font.color.rgb = RGBColor(0x88, 0x88, 0x88)

    def add_rule():
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after  = Pt(4)
        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "6")
        bottom.set(qn("w:space"), "1")
        bottom.set(qn("w:color"), "CCCCCC")
        pBdr.append(bottom)
        pPr.append(pBdr)

    # Title block
    doc.add_heading("NY Court of Appeals — CourtPASS", level=1)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    p.add_run("Date range: ").bold = True
    p.add_run(f"{start_date} – {end_date}")

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    p.add_run("Generated: ").bold = True
    p.add_run(datetime.now().strftime("%Y-%m-%d %H:%M"))

    total_dl = sum(len(r.get("downloaded_files", [])) for r in results)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(10)
    p.add_run("Total cases: ").bold = True
    p.add_run(f"{len(results)}     ")
    p.add_run("Files downloaded: ").bold = True
    p.add_run(str(total_dl))

    add_rule()

    # Cases
    for r in results:
        title = r.get("title") or r.get("full_title", "Unknown")
        ji = r.get("ji_found", False)

        h = doc.add_heading(("⚠ [JI]  " if ji else "") + title, level=2)
        h.paragraph_format.space_before = Pt(12)
        h.paragraph_format.space_after  = Pt(4)

        if ji:
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(4)
            wr = p.add_run("⚠ Jurisdictional Inquiry Detected")
            wr.bold = True
            wr.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
            for entry in r.get("ji_filings", []):
                add_bullet(entry)

        if r.get("apl_number"):
            add_field("APL", r["apl_number"])
        if r.get("argument_date"):
            add_field("Argument Date", r["argument_date"])
        if r.get("decision_date"):
            add_field("Decision Date", r["decision_date"])
        if r.get("argument_number"):
            add_field("Argument No.", r["argument_number"])

        issues = (r.get("issues") or "").strip()
        if issues:
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after  = Pt(2)
            p.add_run("Issues: ").bold = True
            doc.add_paragraph(issues).paragraph_format.space_after = Pt(4)

        if r.get("error"):
            p = doc.add_paragraph()
            p.add_run("Note: ").bold = True
            p.add_run(r["error"]).italic = True

        downloaded = r.get("downloaded_files", [])
        if downloaded:
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after  = Pt(2)
            p.add_run(f"Downloaded ({len(downloaded)} files):").bold = True
            for d in downloaded:
                add_bullet(d["file"])

        unavailable = r.get("unavailable_files", [])
        if unavailable:
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after  = Pt(2)
            rl = p.add_run(f"Not yet available ({len(unavailable)} files):")
            rl.bold = True
            rl.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
            for u in unavailable:
                add_bullet(u, gray=True)

        add_rule()

    docx_path = output_base / "index.docx"
    doc.save(str(docx_path))
    print(f"[index] Wrote {docx_path}")


def write_index(results: list[dict], output_base: Path, start_date: str = "", end_date: str = ""):
    """Write JSON and Markdown index files."""
    # JSON
    json_path = output_base / "index.json"
    with open(json_path, "w") as f:
        json.dump(results, f, indent=2)
    print(f"\n[index] Wrote {json_path}")

    # Markdown
    md_path = output_base / "index.md"
    with open(md_path, "w") as f:
        f.write(f"# NY Court of Appeals — CourtPASS Scrape\n\n")
        f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n\n")
        f.write(f"Total cases: {len(results)}\n\n---\n\n")
        for r in results:
            title = r.get('title') or r.get('full_title', 'Unknown')
            ji = r.get("ji_found", False)
            f.write(f"## {'⚠️ [JI] ' if ji else ''}{title}\n\n")
            if r.get("apl_number"):
                f.write(f"**APL:** {r['apl_number']}  \n")
            if r.get("argument_date"):
                f.write(f"**Argument Date:** {r['argument_date']}  \n")
            if r.get("decision_date"):
                f.write(f"**Decision Date:** {r['decision_date']}  \n")
            if r.get("argument_number"):
                f.write(f"**Argument No.:** {r['argument_number']}  \n")
            if ji:
                f.write(f"\n> ⚠️ **Jurisdictional Inquiry detected**\n")
                for ji_entry in r.get("ji_filings", []):
                    f.write(f"> - {ji_entry}\n")
                f.write("\n")
            if r.get("issues"):
                f.write(f"**Issues:** {r['issues']}  \n")
            f.write(f"**Output folder:** `{r.get('output_dir', '')}` \n\n")

            downloaded = r.get("downloaded_files", [])
            if downloaded:
                f.write(f"**Downloaded ({len(downloaded)} files):**\n\n")
                for d in downloaded:
                    f.write(f"- `{d['file']}`\n")
                f.write("\n")

            unavailable = r.get("unavailable_files", [])
            if unavailable:
                f.write(f"**Not yet available ({len(unavailable)} files):**\n\n")
                for u in unavailable:
                    f.write(f"- `{u}`\n")
                f.write("\n")

            if r.get("error"):
                f.write(f"**Error:** {r['error']}\n\n")

            f.write("---\n\n")

    print(f"[index] Wrote {md_path}")

    write_docx(results, output_base, start_date, end_date)
    write_issues_docx(results, output_base, start_date, end_date)


async def scrape_one_case(page, context, case: dict, start_date: str, end_date: str, output_base: Path) -> dict:
    """Open one case, download its files, and check its docket."""
    # Re-run the search that found this case, then let get_case_detail
    # find the right button by aria-label (title match).
    await open_case_results(page, case, start_date, end_date)

    case_detail = await get_case_detail(page, case)
    if "error" in case_detail:
        print(f"  [error] {case_detail['error']}")
        return case_detail

    n_available = sum(1 for f in case_detail["files"] if f["available"])
    print(f"  Files: {len(case_detail['files'])} total, {n_available} available for download")

    if n_available == 0:
        print("  [skip] No downloadable files.")
        result = {**case_detail, "output_dir": "", "downloaded_files": [], "unavailable_files": [f["pdf_name"] for f in case_detail["files"]]}
    else:
        # Download while still on the case detail page
        result = await process_case(page, context, case_detail, output_base)

    # Check docket for JI after downloads (navigates away from case detail)
    docket_info = await check_docket_for_ji(page, case["title"])
    result.update(docket_info)
    if docket_info["ji_found"]:
        print(f"  *** JURISDICTIONAL INQUIRY DETECTED ***")
        for ji in docket_info["ji_filings"]:
            print(f"      {ji}")
    if docket_info.get("apl_number"):
        print(f"  APL: {docket_info['apl_number']}")

    # Write per-case issues.docx (needs docket info, so called after update)
    if result.get("output_dir"):
        write_case_issues_docx(result, Path(result["output_dir"]))
    elif (result.get("issues") or "").strip():
        # No downloadable files but issues text exists — create folder anyway
        case_dir = output_base / case_folder_name(result)
        case_dir.mkdir(parents=True, exist_ok=True)
        result["output_dir"] = str(case_dir)
        write_case_issues_docx(result, case_dir)

    return result


def load_progress(output_base: Path, start_date: str, end_date: str) -> dict:
    """Cases already finished by an interrupted run of this same date range."""
    try:
        with open(output_base / PROGRESS_FILE) as f:
            data = json.load(f)
    except (OSError, ValueError):
        return {}
    if data.get("start") != start_date or data.get("end") != end_date:
        return {}
    return {r["title"].lower().strip(): r for r in data.get("results", [])}


def save_progress(output_base: Path, start_date: str, end_date: str, saved: dict):
    """Save after every case, so a crash never loses finished work."""
    tmp = output_base / (PROGRESS_FILE + ".tmp")
    with open(tmp, "w") as f:
        json.dump({"start": start_date, "end": end_date, "results": list(saved.values())}, f, indent=2)
    os.replace(tmp, output_base / PROGRESS_FILE)


async def fresh_page(context, old_page):
    """Replace a page that may have crashed or be stuck."""
    try:
        await old_page.close()
    except Exception:
        pass
    try:
        return await context.new_page()
    except Exception as e:
        print(f"\nERROR: Chrome is no longer reachable ({e}).\n"
              "Finished cases are saved. Reopen Chrome and run the same command to continue.")
        sys.exit(1)


async def scrape(start_date: str, end_date: str, output_dir: str):
    output_base = Path(output_dir)
    output_base.mkdir(parents=True, exist_ok=True)
    print(f"\nOutput directory: {output_base.resolve()}")

    async with async_playwright() as p:
        print(f"\nConnecting to Chrome at {CDP_URL}...")
        try:
            browser = await p.chromium.connect_over_cdp(CDP_URL)
        except Exception as e:
            print(f"\nERROR: Could not connect to Chrome.\n"
                  f"Make sure Chrome is running with:\n"
                  f'  open -a "Google Chrome" --args --remote-debugging-port=9222 --user-data-dir=/tmp/chrome_debug\n'
                  f"Details: {e}")
            sys.exit(1)

        context = browser.contexts[0]
        page = await context.new_page()

        # Search by argument date, then decision date
        for attempt in range(1, MAX_ATTEMPTS + 1):
            try:
                arg_cases = await asyncio.wait_for(do_search(page, start_date, end_date), SEARCH_TIME_LIMIT_SECONDS)
                dec_cases = await asyncio.wait_for(do_decision_date_search(page, start_date, end_date), SEARCH_TIME_LIMIT_SECONDS)
                break
            except Exception as e:
                print(f"  [retry] Search attempt {attempt} of {MAX_ATTEMPTS} failed: {str(e).splitlines()[0] if str(e) else e!r}")
                if attempt == MAX_ATTEMPTS:
                    print("\nERROR: CourtPASS search kept failing. Wait a few minutes and run the same command again.")
                    sys.exit(1)
                await asyncio.sleep(RETRY_WAIT_SECONDS * attempt)
                page = await fresh_page(context, page)

        # One entry per case, keeping the dates found by each search
        unique_cases = merge_search_results(arg_cases, dec_cases)

        if not unique_cases:
            print("\nNo cases found. Try broadening the date range.")
            print("The court sits in sessions, so some months have no arguments or decisions.")
            await page.close()
            return

        print(f"\n{'='*60}")
        print(f"Found {len(unique_cases)} unique cases. Processing each...")
        print(f"{'='*60}")

        results = []
        saved = load_progress(output_base, start_date, end_date)
        if saved:
            print(f"Continuing an earlier run: {len(saved)} cases already finished.")
        for i, case in enumerate(unique_cases, 1):
            shown_date = case.get("argument_date") or case.get("decision_date") or ""
            print(f"\n[{i}/{len(unique_cases)}] {datetime.now():%H:%M} {case['title']} ({shown_date})", flush=True)

            key = case["title"].lower().strip()
            if key in saved:
                print("  [skip] Already finished in an earlier run.")
                results.append(saved[key])
                continue

            result = None
            for attempt in range(1, MAX_ATTEMPTS + 1):
                try:
                    result = await asyncio.wait_for(
                        scrape_one_case(page, context, case, start_date, end_date, output_base),
                        CASE_TIME_LIMIT_SECONDS)
                    break
                except Exception as e:
                    print(f"  [retry] Attempt {attempt} of {MAX_ATTEMPTS} failed: {str(e).splitlines()[0] if str(e) else e!r}")
                    if attempt == MAX_ATTEMPTS:
                        result = {**case, "error": f"Failed after {MAX_ATTEMPTS} attempts: {e}"}
                    else:
                        await asyncio.sleep(RETRY_WAIT_SECONDS * attempt)
                    page = await fresh_page(context, page)

            results.append(result)
            if not result.get("error") and not result.get("failed_downloads"):
                saved[key] = result
                save_progress(output_base, start_date, end_date, saved)

            # Small delay between cases to be polite
            await asyncio.sleep(DELAY_BETWEEN_CASES_SECONDS)

        await page.close()

        write_index(results, output_base, start_date, end_date)

        # Summary
        total_downloaded = sum(len(r.get("downloaded_files", [])) for r in results)
        print(f"\n{'='*60}")
        failed = [r for r in results if r.get("error")]
        print(f"Done! Scraped {len(results) - len(failed)} of {len(results)} cases, {total_downloaded} files downloaded.")
        if failed:
            print(f"{len(failed)} cases failed. Run the same command again to retry only those:")
            for r in failed:
                print(f"  - {r['title']}: {r['error'].splitlines()[0]}")
        missing = [(r["title"], name) for r in results for name in r.get("failed_downloads", [])]
        if missing:
            print(f"{len(missing)} PDFs did not download. Run the same command again to retry only those cases:")
            for title, name in missing:
                print(f"  - {title}: {name}")
        if not failed and not missing:
            (output_base / PROGRESS_FILE).unlink(missing_ok=True)
        print(f"Output: {output_base.resolve()}")
        print(f"Index:  {output_base.resolve()}/index.md")
        print(f"{'='*60}")


def main():
    parser = argparse.ArgumentParser(
        description="Scrape NY Court of Appeals CourtPASS for case materials"
    )
    default_end = datetime.today().strftime("%m/%d/%Y")
    default_start = (datetime.today() - timedelta(days=30)).strftime("%m/%d/%Y")

    parser.add_argument("--start", default=default_start,
                        help=f"Start date MM/DD/YYYY (default: {default_start})")
    parser.add_argument("--end", default=default_end,
                        help=f"End date MM/DD/YYYY (default: {default_end})")
    parser.add_argument("--output", default="./courtpass_output",
                        help="Output directory (default: ./courtpass_output)")
    args = parser.parse_args()

    print(f"CourtPASS Scraper")
    print(f"  Date range: {args.start} – {args.end}")
    print(f"  Output:     {args.output}")

    asyncio.run(scrape(args.start, args.end, args.output))


if __name__ == "__main__":
    main()
