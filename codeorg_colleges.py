"""Open code.org, sign in with Google, then capture student progress for every college.

Each college is a classroom section on the code.org home page. For every section
the script opens "View progress" and records each student's display name and
whether their lesson cell is filled in (done) or blank (not done).

Google sign-in is completed by hand in the browser window the first time. The
session is stored in .browser_profile/, so later runs skip the login step.

Usage:
    python codeorg_colleges.py                    # every college
    python codeorg_colleges.py --section "Agra"   # only colleges whose name matches

Writes output/colleges.csv (name + section code) and output/progress.csv. A full
run also saves output/snapshots/progress_<date>.csv and an Excel report comparing
it with the previous snapshot (see progress_report.py).
"""

import argparse
import csv
import re
import sys
from datetime import date, datetime
from pathlib import Path
from urllib.parse import urlparse

from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import TimeoutError as PlaywrightTimeout
from playwright.sync_api import sync_playwright

import progress_report

ROOT = Path(__file__).parent
PROFILE_DIR = ROOT / ".browser_profile"
OUTPUT_DIR = ROOT / "output"

HOME_URL = "https://studio.code.org/home"
SIGN_IN_URL = "https://studio.code.org/users/sign_in"
SECTIONS_API_URL = "https://studio.code.org/dashboardapi/sections"
LOGIN_TIMEOUT_MS = 5 * 60 * 1000
TABLE_TIMEOUT_MS = 90 * 1000

STUDENT_ROWS = '#ui-test-progress-table-v2 button[id^="ui-test-student-row"]'

# Student names and lesson cells are separate columns, matched up by position.
# A cell is "lesson-data-cell-<lessonId>-<studentId>"; only the first lesson is read.
READ_TABLE_JS = """() => {
    const names = [...document.querySelectorAll('%s')].map(b => b.innerText.trim());
    const cells = [...document.querySelectorAll('[data-testid^="lesson-data-cell-"]')];
    const lessonOf = cell => cell.dataset.testid.split('-')[3];
    const lessons = [...new Set(cells.map(lessonOf))];
    const statuses = cells.filter(cell => lessonOf(cell) === lessons[0]).map(cell => {
        const labelled = cell.querySelector('[aria-label]');
        if (labelled) return labelled.getAttribute('aria-label');
        return cell.querySelector('i, svg') || cell.innerText.trim() ? 'Yes' : '';
    });
    return {names, statuses, lessons: lessons.length};
}""" % STUDENT_ROWS


def launch(playwright):
    """Start a real installed browser; Google rejects sign-in from bundled Chromium."""
    last_error = None
    for channel in ("chrome", "msedge"):
        try:
            return playwright.chromium.launch_persistent_context(
                str(PROFILE_DIR),
                channel=channel,
                headless=False,
                no_viewport=True,
                args=["--disable-blink-features=AutomationControlled", "--start-maximized"],
                ignore_default_args=["--enable-automation"],
            )
        except PlaywrightError as error:
            last_error = error
    raise SystemExit(f"Could not start Chrome or Edge: {last_error}")


def is_signed_in_url(url):
    # Compare the hostname: Google's sign-in URL carries studio.code.org in its query string.
    parts = urlparse(url)
    return parts.hostname == "studio.code.org" and not parts.path.startswith("/users/")


def settle(page):
    """Wait for the page to finish loading, without failing on pages that never go idle."""
    try:
        page.wait_for_load_state("networkidle", timeout=15000)
    except PlaywrightTimeout:
        pass


def login(page):
    page.goto(HOME_URL, wait_until="domcontentloaded")
    settle(page)
    if is_signed_in_url(page.url):
        print("Already signed in.")
        return

    if urlparse(page.url).path != "/users/sign_in":
        page.goto(SIGN_IN_URL, wait_until="domcontentloaded")

    try:
        page.locator(
            'form[action*="google_oauth2"] button, a[href*="google_oauth2"]'
        ).or_(page.get_by_role("button", name=re.compile("google", re.I))).first.click(timeout=10000)
    except PlaywrightError:
        print('Could not find the Google button - click "Continue with Google" yourself.')

    print("Finish the Google sign-in in the browser window (waiting up to 5 minutes)...")
    try:
        page.wait_for_url(is_signed_in_url, timeout=LOGIN_TIMEOUT_MS)
    except PlaywrightTimeout:
        raise SystemExit("Timed out waiting for sign-in.")
    except PlaywrightError:
        raise SystemExit("The browser window was closed before sign-in finished.")
    settle(page)
    print("Signed in.")


def fetch_sections(page):
    """Return the signed-in teacher's visible sections, in the order the home page shows them."""
    response = page.request.get(SECTIONS_API_URL)
    if not response.ok:
        raise SystemExit(f"Could not load the section list (HTTP {response.status}).")
    sections = [section for section in response.json() if not section.get("hidden")]
    sections.reverse()
    return sections


def read_progress(page, section):
    """Open a section's progress page and return [(display name, status)], or None on failure."""
    expected = section.get("numberOfStudents") or 0
    page.goto(section["linkToProgress"], wait_until="domcontentloaded")
    if expected == 0:
        return []

    try:
        page.wait_for_function(
            "([selector, count]) => document.querySelectorAll(selector).length >= count",
            arg=[STUDENT_ROWS, expected],
            timeout=TABLE_TIMEOUT_MS,
        )
    except PlaywrightTimeout:
        pass
    settle(page)
    page.evaluate("window.scrollTo(0, document.body.scrollHeight)")

    # Cells fill in after the rows appear; read until two consecutive reads agree.
    table = page.evaluate(READ_TABLE_JS)
    for _ in range(15):
        page.wait_for_timeout(2000)
        latest = page.evaluate(READ_TABLE_JS)
        if latest == table and len(latest["names"]) == len(latest["statuses"]):
            break
        table = latest

    names, statuses = table["names"], table["statuses"]
    if not names or len(names) != len(statuses):
        print(f"      could not read table: {len(names)} students, {len(statuses)} lesson cells")
        return None
    if len(names) != expected:
        print(f"      warning: section lists {expected} students, table shows {len(names)}")
    if table["lessons"] > 1:
        print(f"      warning: {table['lessons']} lessons shown, only the first is recorded")
    return list(zip(names, statuses))


def write_csv(path, header, rows):
    with path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(header)
        writer.writerows(rows)


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument(
        "--section", help="only colleges whose name matches this case-insensitive regex"
    )
    args = parser.parse_args()
    print(f"Run started {datetime.now():%Y-%m-%d %H:%M}")

    with sync_playwright() as playwright:
        context = launch(playwright)
        page = context.pages[0] if context.pages else context.new_page()
        try:
            login(page)
            sections = fetch_sections(page)
            OUTPUT_DIR.mkdir(exist_ok=True)
            write_csv(
                OUTPUT_DIR / "colleges.csv",
                ["College Name", "Section Code"],
                [(section["name"], section["code"]) for section in sections],
            )
            if args.section:
                pattern = re.compile(args.section, re.I)
                sections = [section for section in sections if pattern.search(section["name"])]
            print(f"Found {len(sections)} colleges.")

            rows = []
            failed = []
            for index, section in enumerate(sections, start=1):
                print(f"  [{index}/{len(sections)}] {section['name']} ({section['code']})")
                students = None
                for _ in range(2):
                    try:
                        students = read_progress(page, section)
                    except PlaywrightError as error:
                        print(f"      failed: {error}")
                    if students is not None:
                        break
                if students is None:
                    failed.append(section["name"])
                    continue
                done = sum(1 for _, status in students if status)
                print(f"      {len(students)} students, {done} done")
                for name, status in students:
                    rows.append(
                        (section["name"], section["code"], name, "Yes" if status else "No", status)
                    )

            progress_path = OUTPUT_DIR / "progress.csv"
            write_csv(progress_path, progress_report.SNAPSHOT_HEADER, rows)
            print(f"Saved {len(rows)} students to {progress_path}")
            if failed:
                print(f"Failed colleges ({len(failed)}): {', '.join(failed)}")

            # A filtered run is not a full day's snapshot, so it is left out of the comparison.
            if not args.section and rows:
                today = date.today()
                progress_report.write_snapshot(today, rows)
                print(f"Saved report to {progress_report.build_report_for(today)}")
            return 1 if failed else 0
        finally:
            context.close()


if __name__ == "__main__":
    sys.exit(main())
