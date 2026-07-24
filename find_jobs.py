#!/usr/bin/env python3
"""
Salesforce Job Finder
Searches LinkedIn, Indeed, Naukri, Dice, Glassdoor, and Wellfound
for jobs posted in the last 24 hours in India and saves them to an Excel sheet.
"""
import asyncio
import re
from datetime import datetime, date, timedelta
from pathlib import Path
from playwright.async_api import async_playwright, Page, TimeoutError as PWT
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

OUTPUT_FILE  = Path(__file__).parent / "salesforce_jobs.xlsx"
JOBS_JSON    = Path(__file__).parent / "jobs.json"

UA = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/124.0.0.0 Safari/537.36"
)

KEYWORDS = [
    "Senior Salesforce Developer",
    "Salesforce Consultant",
    "Salesforce Technical Consultant",
    "Salesforce LWC Developer",
]

# ── helpers ────────────────────────────────────────────────────────────────────

async def safe_goto(page: Page, url: str, timeout=20000):
    try:
        await page.goto(url, wait_until="domcontentloaded", timeout=timeout)
        await page.wait_for_timeout(2000)
    except Exception:
        pass

def parse_age(text: str) -> int:
    """Return age in days. Returns 0 for unknown/empty (assume recent)."""
    text = (text or "").lower().strip()
    if not text:
        return 0
    # ISO date like 2026-07-18
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", text)
    if m:
        try:
            d = date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            return (date.today() - d).days
        except Exception:
            pass
    if any(t in text for t in ("just now", "today", "moments ago", "new")):
        return 0
    m = re.search(r"(\d+)\s*(hour|hr|minute|min|sec)", text)
    if m:
        return 0
    m = re.search(r"(\d+)\s*day", text)
    if m:
        return int(m.group(1))
    m = re.search(r"(\d+)\s*week", text)
    if m:
        return int(m.group(1)) * 7
    m = re.search(r"(\d+)\s*month", text)
    if m:
        return int(m.group(1)) * 30
    # unrecognised — assume recent so we don't drop it
    return 0

def is_recent(age_days: int) -> bool:
    return age_days <= 3

# ── scrapers ───────────────────────────────────────────────────────────────────

async def scrape_linkedin(page: Page, keyword: str) -> list[dict]:
    jobs = []
    kw = keyword.replace(" ", "%20")
    # f_TPR=r259200 = past 3 days, geoId=102713980 = India
    url = f"https://www.linkedin.com/jobs/search/?keywords={kw}&f_TPR=r259200&geoId=102713980&sortBy=DD"
    await safe_goto(page, url)

    cards = await page.query_selector_all(".base-card, .job-search-card")
    for card in cards:
        try:
            title_el   = await card.query_selector(".base-search-card__title, h3")
            company_el = await card.query_selector(".base-search-card__subtitle, h4")
            link_el    = await card.query_selector("a.base-card__full-link, a")
            date_el    = await card.query_selector("time, .job-search-card__listdate")
            loc_el     = await card.query_selector(".job-search-card__location, .base-search-card__metadata span")

            if not (title_el and link_el):
                continue

            title    = (await title_el.inner_text()).strip()
            company  = (await company_el.inner_text()).strip() if company_el else ""
            href     = await link_el.get_attribute("href") or ""
            location = (await loc_el.inner_text()).strip() if loc_el else "India"
            age_txt  = ""
            if date_el:
                age_txt = await date_el.get_attribute("datetime") or await date_el.inner_text()

            age = parse_age(age_txt) if age_txt else 0
            if not is_recent(age):
                continue

            jobs.append({
                "title":    title,
                "company":  company,
                "platform": "LinkedIn",
                "location": location,
                "url":      href,
                "posted":   age_txt[:20] if age_txt else "Recent",
                "age_days": age,
            })
        except Exception:
            continue

    print(f"  LinkedIn '{keyword}': {len(jobs)} recent jobs")
    return jobs


async def scrape_indeed(page: Page, keyword: str) -> list[dict]:
    jobs = []
    kw = keyword.replace(" ", "+")
    # fromage=3 = last 3 days, l=India
    url = f"https://in.indeed.com/jobs?q={kw}&l=India&sort=date&fromage=3"
    await safe_goto(page, url)

    for sel in ["a.jcs-JobTitle", "[data-testid='job-title'] a", "h2.jobTitle a", ".job_seen_beacon h2 a"]:
        links = await page.query_selector_all(sel)
        if not links:
            continue
        for link in links:
            try:
                title = (await link.inner_text()).strip()
                href  = await link.get_attribute("href") or ""
                if not href.startswith("http"):
                    href = "https://www.indeed.com" + href

                # company
                parent = await link.evaluate_handle("el => el.closest('[data-jk], .job_seen_beacon, .resultContent')")
                company = ""
                try:
                    co = await parent.query_selector("[data-testid='company-name'], .companyName, span.company")
                    if co:
                        company = (await co.inner_text()).strip()
                except Exception:
                    pass

                # date
                age_txt = ""
                try:
                    dt = await parent.query_selector("[data-testid='myJobsStateDate'], .date, span.date")
                    if dt:
                        age_txt = (await dt.inner_text()).strip()
                except Exception:
                    pass

                # location
                location = "India"
                try:
                    loc = await parent.query_selector("[data-testid='text-location'], .companyLocation")
                    if loc:
                        location = (await loc.inner_text()).strip()
                except Exception:
                    pass

                age = parse_age(age_txt) if age_txt else 0
                if title and href:
                    jobs.append({
                        "title":    title,
                        "company":  company or "Unknown",
                        "platform": "Indeed",
                        "location": location,
                        "url":      href,
                        "posted":   age_txt or "Recent",
                        "age_days": age,
                    })
            except Exception:
                continue
        if jobs:
            break

    print(f"  Indeed '{keyword}': {len(jobs)} jobs")
    return jobs


async def scrape_naukri(page: Page, keyword: str) -> list[dict]:
    jobs = []
    kw = keyword.lower().replace(" ", "-")
    # jobAge=1 = last 24 hours, India is default for naukri.com
    url = f"https://www.naukri.com/{kw}-jobs-in-india?experience=2&jobAge=1"
    await safe_goto(page, url)

    cards = await page.query_selector_all(".jobTupleHeader, article.jobTuple")
    for card in cards:
        try:
            title_el   = await card.query_selector("a.title, .jobTitle a")
            company_el = await card.query_selector(".companyInfo a, .comp-name, span.comp-name")
            date_el    = await card.query_selector(".job-post-day, .fleft.grey-text.fs12.fw500")

            if not title_el:
                continue

            title   = (await title_el.inner_text()).strip()
            href    = await title_el.get_attribute("href") or ""
            company = (await company_el.inner_text()).strip() if company_el else ""
            age_txt = (await date_el.inner_text()).strip() if date_el else ""

            age = parse_age(age_txt)
            if age > 1:
                continue

            jobs.append({
                "title":    title,
                "company":  company,
                "platform": "Naukri",
                "url":      href,
                "posted":   age_txt or "Recent",
            })
        except Exception:
            continue

    print(f"  Naukri '{keyword}': {len(jobs)} recent jobs")
    return jobs


async def scrape_dice(page: Page, keyword: str) -> list[dict]:
    jobs = []
    kw = keyword.replace(" ", "%20")
    url = f"https://www.dice.com/jobs?q={kw}&countryCode=IN&filters.postedDate=ONE_DAY&language=en"
    await safe_goto(page, url)
    await page.wait_for_timeout(3000)  # Dice is JS-heavy

    cards = await page.query_selector_all("dhi-search-card, .card-title-container")
    for card in cards:
        try:
            title_el   = await card.query_selector("a[data-cy='card-title-link'], h5 a, .card-title a")
            company_el = await card.query_selector("[data-cy='search-result-company-name'], .card-company")
            date_el    = await card.query_selector("[data-cy='card-posted-date'], .posted-date")

            if not title_el:
                continue

            title   = (await title_el.inner_text()).strip()
            href    = await title_el.get_attribute("href") or ""
            if not href.startswith("http"):
                href = "https://www.dice.com" + href
            company = (await company_el.inner_text()).strip() if company_el else ""
            age_txt = (await date_el.inner_text()).strip() if date_el else ""

            jobs.append({
                "title":    title,
                "company":  company,
                "platform": "Dice",
                "url":      href,
                "posted":   age_txt or "Recent",
            })
        except Exception:
            continue

    print(f"  Dice '{keyword}': {len(jobs)} jobs")
    return jobs


async def scrape_wellfound(page: Page, keyword: str) -> list[dict]:
    jobs = []
    kw = keyword.replace(" ", "+")
    url = f"https://wellfound.com/jobs?keywords={kw}&location=India"
    await safe_goto(page, url)

    cards = await page.query_selector_all("[class*='JobListing'], [data-test*='job']")
    for card in cards[:30]:
        try:
            title_el   = await card.query_selector("a[class*='title'], h2 a, h3 a")
            company_el = await card.query_selector("[class*='company'], [class*='startup']")
            date_el    = await card.query_selector("[class*='date'], [class*='time'], time")

            if not title_el:
                continue

            title   = (await title_el.inner_text()).strip()
            href    = await title_el.get_attribute("href") or ""
            if not href.startswith("http"):
                href = "https://wellfound.com" + href
            company = (await company_el.inner_text()).strip() if company_el else ""
            age_txt = (await date_el.inner_text()).strip() if date_el else ""

            age = parse_age(age_txt)
            if age > 1:
                continue

            jobs.append({
                "title":    title,
                "company":  company,
                "platform": "Wellfound",
                "url":      href,
                "posted":   age_txt or "Recent",
            })
        except Exception:
            continue

    print(f"  Wellfound '{keyword}': {len(jobs)} recent jobs")
    return jobs


async def scrape_lever_api(keyword: str) -> list[dict]:
    """Use Lever's public JSON API — no browser needed."""
    import httpx
    jobs = []
    companies = [
        "salesforce", "slalom", "deloitte", "accenture", "capgemini",
        "cognizant", "publicissapient", "virtusa", "concentrix",
    ]
    kw_lower = keyword.lower()
    try:
        async with httpx.AsyncClient(timeout=10) as client:
            for company in companies:
                try:
                    r = await client.get(f"https://api.lever.co/v0/postings/{company}?mode=json")
                    if r.status_code != 200:
                        continue
                    for posting in r.json():
                        title = posting.get("text", "")
                        if not any(k in title.lower() for k in ["salesforce", "lwc", "apex", "crm"]):
                            continue
                        created_ms = posting.get("createdAt", 0)
                        if created_ms:
                            created = datetime.fromtimestamp(created_ms / 1000)
                            age = (datetime.now() - created).days
                            if age > 1:
                                continue
                            age_txt = f"{age}d ago" if age > 0 else "Today"
                        else:
                            age_txt = "Recent"
                        jobs.append({
                            "title":    title,
                            "company":  company.capitalize(),
                            "platform": "Lever (ATS)",
                            "url":      posting.get("hostedUrl", ""),
                            "posted":   age_txt,
                        })
                except Exception:
                    continue
    except Exception as e:
        print(f"  Lever API error: {e}")
    print(f"  Lever API '{keyword}': {len(jobs)} recent jobs")
    return jobs


# ── JSON writer ────────────────────────────────────────────────────────────────

def write_json(jobs: list[dict]):
    import json
    seen = set()
    unique = []
    for j in jobs:
        key = (j["title"].lower().strip(), j["company"].lower().strip())
        if j["url"] and key not in seen:
            seen.add(key)
            unique.append(j)
    unique.sort(key=lambda j: j["platform"])
    payload = {
        "jobs":      unique,
        "total":     len(unique),
        "scraped_at": datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"),
        "platforms": sorted({j["platform"] for j in unique}),
    }
    JOBS_JSON.write_text(json.dumps(payload, indent=2))


# ── Excel writer ────────────────────────────────────────────────────────────────

def write_excel(jobs: list[dict]):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Salesforce Jobs"

    # Header style
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    headers = ["#", "Job Title", "Company", "Platform", "Posted", "Apply Link"]
    col_widths = [5, 45, 30, 15, 12, 80]

    for col, (h, w) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = w

    ws.row_dimensions[1].height = 20

    # Alternate row colors
    fill_even = PatternFill("solid", fgColor="EBF3FB")
    fill_odd  = PatternFill("solid", fgColor="FFFFFF")

    # Deduplicate by title+company (URLs vary per platform)
    seen = set()
    unique = []
    for j in jobs:
        key = (j["title"].lower().strip(), j["company"].lower().strip())
        if j["url"] and key not in seen:
            seen.add(key)
            unique.append(j)

    unique.sort(key=lambda j: j["platform"])

    for i, job in enumerate(unique, 1):
        row = i + 1
        fill = fill_even if i % 2 == 0 else fill_odd
        for col in range(1, 7):
            ws.cell(row=row, column=col).fill = fill

        ws.cell(row=row, column=1, value=i)
        ws.cell(row=row, column=2, value=job["title"])
        ws.cell(row=row, column=3, value=job["company"])
        ws.cell(row=row, column=4, value=job["platform"])
        ws.cell(row=row, column=5, value=job["posted"])

        # Clickable hyperlink
        url = job["url"]
        link_cell = ws.cell(row=row, column=6, value=url)
        if url.startswith("http"):
            link_cell.hyperlink = url
            link_cell.font = Font(color="0563C1", underline="single")

        ws.row_dimensions[row].height = 15

    # Freeze header row
    ws.freeze_panes = "A2"

    wb.save(OUTPUT_FILE)
    return len(unique)


# ── main ───────────────────────────────────────────────────────────────────────

async def main():
    print(f"\n{'='*55}")
    print("  Salesforce Job Finder — last 24 hours | India only")
    print(f"{'='*55}\n")

    all_jobs = []

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=True)

        for keyword in KEYWORDS:
            print(f"\n[{keyword}]")

            # Each scraper gets its own page to avoid navigation conflicts
            async def run_scraper(scraper_fn, kw=keyword):
                ctx = await browser.new_context(user_agent=UA)
                pg  = await ctx.new_page()
                try:
                    return await scraper_fn(pg, kw)
                except Exception as e:
                    print(f"  {scraper_fn.__name__} error: {e}")
                    return []
                finally:
                    await ctx.close()

            results = await asyncio.gather(
                run_scraper(scrape_linkedin),
                run_scraper(scrape_indeed),
                run_scraper(scrape_naukri),
                run_scraper(scrape_dice),
                run_scraper(scrape_wellfound),
                scrape_lever_api(keyword),
            )
            for r in results:
                all_jobs.extend(r)

        await browser.close()

    total = write_excel(all_jobs)
    write_json(all_jobs)
    print(f"\n{'='*55}")
    print(f"  Done! {total} unique jobs saved to:")
    print(f"  {OUTPUT_FILE}")
    print(f"  {JOBS_JSON}")
    print(f"{'='*55}\n")

if __name__ == "__main__":
    asyncio.run(main())
