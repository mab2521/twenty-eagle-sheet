#!/usr/bin/env python3
"""
sheet_sync.py — send CourtPASS scraper results to the review Google Sheet.

Reads the scraper's index.json, reshapes each case into one Sheet row, and
POSTs it to the Apps Script web app attached to the Sheet (see Code.gs).
Uses only the Python standard library.

Setup (once):
  export SHEET_WEBAPP_URL="https://script.google.com/macros/s/XXXX/exec"
  export SHEET_SYNC_TOKEN="the same value you saved as SYNC_TOKEN in Apps Script"

Usage:
  # Preview the row without sending anything
  python3 sheet_sync.py courtpass_output/index.json --case "Mann v Mezuyon" --dry-run

  # Send one case (pilot)
  python3 sheet_sync.py courtpass_output/index.json --case "Mann v Mezuyon"

  # Send every successfully scraped case
  python3 sheet_sync.py courtpass_output/index.json --all

Links: CourtPASS has no linkable addresses for its files, so the decision,
argument video and transcript links are looked up on the court's public site
(www.nycourts.gov), matched by the filenames CourtPASS reports. Add
--no-url-lookup to skip that step.

From the scraper (later): `from sheet_sync import push_cases; push_cases(results)`
"""

from __future__ import annotations  # keeps this working on older Python 3

import argparse
import html
import json
import os
import re
import sys
import urllib.request
from datetime import datetime
from urllib.parse import quote, unquote, urljoin, urlsplit, urlunsplit

# Invisible formatting characters CourtPASS sometimes embeds in issues text
_INVISIBLE = re.compile(r"[\u200b-\u200f\u202a-\u202e\u2066-\u2069\ufeff]")
_TRANSCRIPT = re.compile(r"^(\d{2})(\d{2})(\d{2})-(.+?)-Oral-Argument-Transcript", re.I)
_DECISION_FILE = re.compile(r"(opn|mem)\d{2}-Decision", re.I)


def clean(text: str) -> str:
    text = _INVISIBLE.sub("", text or "").replace("\u00a0", " ")
    return re.sub(r"[ \t]+", " ", text).strip()


def split_issues(raw: str) -> tuple[str, str]:
    """CourtPASS issues come as 'Category heading', blank line, issue text
    (sometimes repeated). Return (categories, issue_text)."""
    categories, texts = [], []
    for para in re.split(r"\n\s*\n", clean(raw)):
        para = clean(para.replace("\n", " "))
        if not para:
            continue
        is_heading = (
            len(para) < 100
            and not re.search(r"[.;?:]$", para)
            and not para.lower().startswith("whether")
        )
        (categories if is_heading else texts).append(para)
    return "; ".join(categories), "\n\n".join(texts)


def blog_summary_from(issue_text: str) -> str:
    """First-pass summary in the blog's 'Whether ...; whether ...' style.
    This only pre-fills the human column; the reviewer rewrites it."""
    s = " ".join(issue_text.split())
    s = re.sub(r";\s+Whether\b", "; whether", s)
    s = re.sub(r"[;,]\s*$", "", s)
    if s and s[-1] not in ".?":
        s += "."
    return s


# ───────────── Links on the court's public website ─────────────

COURT_SITE = "https://www.nycourts.gov/CTAPPs"
_MONTHS = ["January", "February", "March", "April", "May", "June", "July",
           "August", "September", "October", "November", "December"]
_page_cache = {}


def _page_links(url: str):
    """All links on a court web page, or None if the page could not be read."""
    if url not in _page_cache:
        try:
            req = urllib.request.Request(url, headers={
                "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                              "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36"})
            with urllib.request.urlopen(req, timeout=30) as resp:
                page = resp.read().decode("utf-8", "replace")
            hrefs = re.findall(r"""href\s*=\s*["']([^"']+)["']""", page, flags=re.I)
            _page_cache[url] = [urljoin(url, html.unescape(h).strip()) for h in hrefs]
        except Exception:
            _page_cache[url] = None
    return _page_cache[url]


def _basename(url: str) -> str:
    return unquote(urlsplit(url).path).rsplit("/", 1)[-1]


def _norm(name: str) -> str:
    return re.sub(r"[^a-z0-9]", "", unquote(name).lower())


def _tidy(url: str) -> str:
    """Encode spaces and '&' so the link works when pasted anywhere."""
    p = urlsplit(url)
    return urlunsplit((p.scheme, p.netloc, quote(unquote(p.path), safe="/"), p.query, ""))


def _month_folder(month: int, year: int) -> str:
    return f"{year}/{_MONTHS[month - 1][:3]}{year % 100:02d}"


def court_urls(row: dict, transcript_file: str, decision_file: str) -> dict:
    """Find the decision, video and transcript links on nycourts.gov.

    A link is 'verified' when it is actually listed on the court's own index
    page. If that page cannot be read, the link is built from the court's
    naming pattern and marked 'not checked' so the reviewer knows to click it.
    """
    out = {"decision_url": "", "video_url": "", "transcript_url": ""}
    notes = []

    # Argument video + transcript: listed on the session's oral-argument page
    m = re.match(r"(\d{2})/(\d{2})/(\d{4})$", row.get("argument_date", ""))
    if m and transcript_file:
        month, year = int(m.group(1)), int(m.group(3))
        numbers = re.findall(r"\d+", row.get("argument_no", ""))
        # A session's folder is normally the argument month; try neighbours too
        candidates = [(month, year),
                      (month % 12 + 1, year + (month == 12)),
                      ((month - 2) % 12 + 1, year - (month == 1))]
        reachable, found = False, False
        for mo, yr in candidates:
            folder = f"{COURT_SITE}/arguments/{_month_folder(mo, yr)}"
            links = _page_links(f"{folder}/{folder.rsplit('/', 1)[-1]}_OA.html")
            if links is None:
                continue
            reachable = True
            tr = [u for u in links if _norm(_basename(u)) == _norm(transcript_file)]
            if not tr:
                continue
            found = True
            out["transcript_url"] = _tidy(tr[0])
            vid = [u for u in links if "/video/" in u.lower()
                   and re.findall(r"\d+", _basename(u)) == numbers]
            if vid:
                out["video_url"] = _tidy(vid[0])
            notes.append("transcript: verified")
            notes.append("video: verified" if vid else "video: not found on court site")
            break
        if not reachable:
            folder = f"{COURT_SITE}/arguments/{_month_folder(month, year)}"
            out["transcript_url"] = _tidy(f"{folder}/Transcripts/{transcript_file}")
            if numbers:
                out["video_url"] = f"{folder}/Video/{'-'.join(numbers)}.html"
            notes.append("transcript + video: NOT CHECKED (built from pattern)")
        elif not found:
            notes.append("transcript + video: not found on court site")

    # Decision: listed on the month's decisions page
    m = re.match(r"(\d{2})/(\d{2})/(\d{4})$", row.get("decision_date", ""))
    if m and decision_file:
        month, year = int(m.group(1)), int(m.group(3))
        folder = f"{COURT_SITE}/Decisions/{_month_folder(month, year)}"
        links = _page_links(f"{folder}/{_MONTHS[month - 1]}{year % 100:02d}.html")
        if links is None:
            out["decision_url"] = _tidy(f"{folder}/{decision_file}")
            notes.append("decision: NOT CHECKED (built from pattern)")
        else:
            dec = [u for u in links if _norm(_basename(u)) == _norm(decision_file)]
            if dec:
                out["decision_url"] = _tidy(dec[0])
                notes.append("decision: verified")
            else:
                notes.append("decision: not found on court site")

    out["url_check"] = "; ".join(notes)
    return out


def session_label(argument_date: str) -> str:
    """'09/11/2024' -> 'September Session 24', the blog's name for a session."""
    m = re.match(r"(\d{2})/\d{2}/\d{2}(\d{2})$", argument_date)
    return f"{_MONTHS[int(m.group(1)) - 1]} Session {m.group(2)}" if m else ""


def case_to_row(case: dict, lookup_urls: bool = True) -> dict:
    title = clean(case.get("title", ""))
    aria = case.get("aria", "")
    files = case.get("files", [])
    names = [f.get("pdf_name", "") for f in files]

    decision_date = clean(case.get("decision_date") or "")
    argument_date = clean(case.get("argument_date") or "")
    argument_no = clean(case.get("argument_number", ""))
    # In index.json files from older scraper runs, "argument_date" held the
    # decision date for cases found by decision date (the aria label says which).
    m = re.search(r"Decision date (\d{2}/\d{2}/\d{4})", aria)
    if m:
        decision_date = decision_date or m.group(1)
        if argument_date == m.group(1):
            argument_date = ""

    # The transcript filename carries the real argument date and number,
    # e.g. "041426-39-Oral-Argument-Transcript.pdf" -> 04/14/2026, No. 39
    transcript_file, decision_file = "", ""
    for n in names:
        t = _TRANSCRIPT.match(n)
        if t:
            transcript_file = n
            argument_date = f"{t.group(1)}/{t.group(2)}/20{t.group(3)}"
            argument_no = argument_no or t.group(4).strip()
            break

    decision_type = ""
    for n in names:
        d = _DECISION_FILE.search(n)
        if d:
            decision_file = n
            decision_type = "Opinion" if d.group(1).lower() == "opn" else "Memorandum"
            break

    category, issue_text = split_issues(case.get("issues", ""))

    # An empty docket result means the check did not run, not that there is no JI.
    if case.get("ji_found"):
        ji_status = "YES"
    elif case.get("all_filings") or case.get("apl_number"):
        ji_status = "No"
    else:
        ji_status = "Not checked"

    row = {
        "case_key": re.sub(r"[^a-z0-9]+", "-", title.lower()).strip("-"),
        "title": title,
        "decision_date": decision_date,
        "decision_type": decision_type,
        "argument_date": argument_date,
        "argument_no": argument_no,
        "session": session_label(argument_date),
        "apl_number": clean(case.get("apl_number") or ""),
        "category": category,
        # Exactly what the scraper collected. Do not clean or reformat this field.
        "issues_courtpass": case.get("issues") or "",
        "ji_status": ji_status,
        "ji_filings": "\n".join(case.get("ji_filings") or []),
        "files_available": "\n".join(f["pdf_name"] for f in files if f.get("available")),
        "files_unavailable": "\n".join(f["pdf_name"] for f in files if not f.get("available")),
        "last_synced": datetime.now().strftime("%Y-%m-%d %H:%M"),
        # Pre-fills the human "Blog summary" column on first insert only
        "blog_summary": blog_summary_from(issue_text),
        "decision_url": "", "video_url": "", "transcript_url": "", "url_check": "",
    }
    if lookup_urls:
        # These also only pre-fill blank human cells; the reviewer confirms them
        row.update(court_urls(row, transcript_file, decision_file))
    return row


def push_cases(cases: list[dict], url: str = "", token: str = "", lookup_urls: bool = True) -> dict:
    """Send scraped cases to the Sheet. Cases the scraper failed on are skipped."""
    url = url or os.environ.get("SHEET_WEBAPP_URL", "")
    token = token or os.environ.get("SHEET_SYNC_TOKEN", "")
    if not url or not token:
        raise SystemExit("Set SHEET_WEBAPP_URL and SHEET_SYNC_TOKEN first (see top of this file).")

    rows = [case_to_row(c, lookup_urls) for c in cases if not c.get("error")]
    body = json.dumps({"token": token, "cases": rows}).encode("utf-8")
    req = urllib.request.Request(url, data=body, headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=360) as resp:  # a month can take minutes to write
        text = resp.read().decode("utf-8")
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        raise SystemExit(
            "The web app did not return JSON. Check that it is deployed with access "
            "'Anyone' and that the URL ends in /exec.\n" + text[:300]
        )


def main():
    ap = argparse.ArgumentParser(description="Send CourtPASS scraper results to the review Sheet")
    ap.add_argument("index_json", help="Path to the scraper's index.json")
    ap.add_argument("--case", help="Send only cases whose title contains this text")
    ap.add_argument("--all", action="store_true", help="Send every successfully scraped case")
    ap.add_argument("--no-url-lookup", action="store_true",
                    help="Do not look up links on nycourts.gov")
    ap.add_argument("--dry-run", action="store_true", help="Print the rows instead of sending")
    args = ap.parse_args()

    if not args.case and not args.all:
        ap.error("choose --case \"title text\" or --all")

    with open(args.index_json, encoding="utf-8") as f:
        cases = json.load(f)

    failed = [c for c in cases if c.get("error")]
    good = [c for c in cases if not c.get("error")]
    if args.case:
        good = [c for c in good if args.case.lower() in c.get("title", "").lower()]
    if not good:
        raise SystemExit("No successfully scraped case matched.")

    if args.dry_run:
        print(json.dumps([case_to_row(c, not args.no_url_lookup) for c in good], indent=2, ensure_ascii=False))
    else:
        result = push_cases(good, lookup_urls=not args.no_url_lookup)
        print(json.dumps(result, indent=2))
        if not result.get("ok"):
            sys.exit(1)

    if failed and args.all:
        print(f"\nSkipped {len(failed)} cases the scraper could not open:", file=sys.stderr)
        for c in failed:
            print(f"  - {c.get('title')}: {c.get('error')}", file=sys.stderr)


if __name__ == "__main__":
    main()