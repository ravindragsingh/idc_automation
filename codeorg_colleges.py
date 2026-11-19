"""Open code.org, sign in with Google, then capture student progress for every college.

Each college is a classroom section on the code.org home page. For every section
the script opens "View progress" and records each student's display name and
whether their lesson cell is filled in (done) or blank (not done).

Google sign-in is completed by hand in the browser window the first time. The
session is stored in .browser_profile/, so later runs skip the login step.

Usage:
    python codeorg_colleges.py                    # every college
    python codeorg_colleges.py --section "Agra"   # only colleges whose name matches
    python codeorg_colleges.py --full             # also re-open colleges already 100% done

A student counts as done when lesson 1 is filled in, except in "How AI Makes
Decisions", which needs lessons 2 and 3 (see REQUIRED_LESSONS).

Colleges that were 100% done in the newest snapshot, and still have the same number
of students, are not opened again; their rows are copied from that snapshot.

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

# Lessons that must all be filled in for a student to count as done, by course
# name as code.org reports it. Any other course uses lesson 1 (AI for Oceans has only one).
REQUIRED_LESSONS = {"how ai makes decisions": (2, 3)}

# Student names and lesson cells are separate columns, matched up by position.
# A cell is "lesson-data-cell-<lessonId>-<studentId>"; the lesson columns appear in
# lesson order (1, 2, 3, ...), so `lessons` holds one list of statuses per lesson.
READ_TABLE_JS = """() => {
    const names = [...document.querySelectorAll('%s')].map(b => b.innerText.trim());
    const cells = [...document.querySelectorAll('[data-testid^="lesson-data-cell-"]')];
    const byLesson = new Map();
    for (const cell of cells) {
        const lesson = cell.dataset.testid.split('-')[3];
        if (!byLesson.has(lesson)) byLesson.set(lesson, []);
        const labelled = cell.querySelector('[aria-label]');
        byLesson.get(lesson).push(
            labelled ? labelled.getAttribute('aria-label')
                : cell.querySelector('i, svg') || cell.innerText.trim() ? 'Yes' : ''
        );
    }
    return {names, lessons: [...byLesson.values()]};
}""" % STUDENT_ROWS


def course_of(section):
    """The course assigned to a section, e.g. "AI for Oceans"."""
    return (section.get("assignedTitle") or "").strip()


def required_lessons(section):
    return REQUIRED_LESSONS.get(course_of(section).lower(), (1,))


def matches_rule(rows, lessons):
    """Whether snapshot rows were recorded with the section's current lesson rule.

    Snapshots taken before a course had its own rule recorded only lesson 1.
    """
    return lessons == (1,) or all(row[4].startswith(f"L{lessons[0]} ") for row in rows)


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
    """Open a section's progress page and return [(display name, done, status)], or None on failure."""
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
        if latest == table and complete_table(latest):
            break
        table = latest

    names, columns = table["names"], table["lessons"]
    if not complete_table(table):
        counts = ", ".join(str(len(column)) for column in columns) or "no"
        print(f"      could not read table: {len(names)} students, {counts} lesson cells")
        return None
    if len(names) != expected:
        print(f"      warning: section lists {expected} students, table shows {len(names)}")

    lessons = required_lessons(section)
    if max(lessons) > len(columns):
        print(f"      could not read table: needs lesson {max(lessons)}, page shows {len(columns)}")
        return None
    if lessons == (1,):
        return [(name, bool(status), status) for name, status in zip(names, columns[0])]

    students = []
    for index, name in enumerate(names):
        statuses = [(number, columns[number - 1][index]) for number in lessons]
        summary = ", ".join(f"L{number} {status}" for number, status in statuses if status)
        students.append((name, all(status for _, status in statuses), summary))
    return students


def complete_table(table):
    names = table["names"]
    return bool(names) and bool(table["lessons"]) and all(
        len(column) == len(names) for column in table["lessons"]
    )


def write_csv(path, header, rows):
    with path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(header)
        writer.writerows(rows)


def completed_sections():
    """Return ({section code: snapshot rows}, snapshot date) for colleges 100% done in the newest snapshot.

    Students do not lose a finished lesson, so these colleges need not be opened again.
    """
    days = progress_report.snapshot_days()
    if not days:
        return {}, None
    by_section = {}
    for row in progress_report.read_snapshot(progress_report.snapshot_path(days[-1])):
        by_section.setdefault(row[1], []).append(row)
    complete = {
        code: rows for code, rows in by_section.items() if all(row[3] == "Yes" for row in rows)
    }
    return complete, days[-1]


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument(
        "--section", help="only colleges whose name matches this case-insensitive regex"
    )
    parser.add_argument(
        "--full", action="store_true", help="also re-read colleges that were already 100%% done"
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
            complete, complete_day = ({}, None) if args.full else completed_sections()

            rows = []
            failed = []
            skipped = 0
            for index, section in enumerate(sections, start=1):
                if page.is_closed():
                    print("The browser window was closed; stopping.")
                    failed.extend(sections[index - 1 :])
                    break
                print(f"  [{index}/{len(sections)}] {section['name']} ({section['code']})")
                earlier = complete.get(section["code"])
                if (
                    earlier
                    and len(earlier) == section.get("numberOfStudents")
                    and matches_rule(earlier, required_lessons(section))
                ):
                    print(f"      skipped: all {len(earlier)} students done on {complete_day}")
                    rows += [(section["name"], *row[1:5], course_of(section)) for row in earlier]
                    skipped += 1
                    continue
                students = None
                for _ in range(2):
                    try:
                        students = read_progress(page, section)
                    except PlaywrightError as error:
                        print(f"      failed: {error}")
                    if students is not None or page.is_closed():
                        break
                if students is None:
                    failed.append(section)
                    continue
                done = sum(1 for _, is_done, _ in students if is_done)
                print(f"      {len(students)} students, {done} done")
                for name, is_done, status in students:
                    rows.append(
                        (
                            section["name"],
                            section["code"],
                            name,
                            "Yes" if is_done else "No",
                            status,
                            course_of(section),
                        )
                    )

            progress_path = OUTPUT_DIR / "progress.csv"
            write_csv(progress_path, progress_report.SNAPSHOT_HEADER, rows)
            print(f"Saved {len(rows)} students to {progress_path}")
            if skipped:
                print(f"Skipped {skipped} colleges that were already 100% done")
            if failed:
                names = ", ".join(section["name"] for section in failed)
                print(f"Failed colleges ({len(failed)}): {names}")

            # A filtered run is not a full day's snapshot, so it is left out of the comparison.
            if not args.section and rows:
                today = date.today()
                snapshot = progress_report.snapshot_path(today)
                if failed and snapshot.exists():
                    # Keep an earlier run's data for the colleges this run could not read.
                    failed_codes = {section["code"] for section in failed}
                    kept = [
                        row
                        for row in progress_report.read_snapshot(snapshot)
                        if row[1] in failed_codes
                    ]
                    rows += kept
                    print(f"Kept {len(kept)} students from today's earlier run for failed colleges")
                progress_report.write_snapshot(today, rows)
                print(f"Saved report to {progress_report.build_report_for(today)}")
            return 1 if failed else 0
        finally:
            context.close()


if __name__ == "__main__":
    sys.exit(main())
