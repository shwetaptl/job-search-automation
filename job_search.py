"""
Job Search Automation

Usage:
  python job_search.py --init path/to/profile.docx   build profile_data.md from your resume doc
  python job_search.py --update "fact to add"         merge a new fact into profile_data.md
  python job_search.py --show                         print the current profile_data.md
  python job_search.py                                run the job search using profile_data.md

Outputs: jobs_YYYYMMDD_HHMMSS.csv sorted by match score (highest first)
"""

import sys
import csv
import re
import time
import os
import argparse
from datetime import datetime

import requests
from bs4 import BeautifulSoup
from docx import Document
try:
    from playwright.sync_api import sync_playwright
    PLAYWRIGHT_AVAILABLE = True
except ImportError:
    PLAYWRIGHT_AVAILABLE = False

# ── Config ────────────────────────────────────────────────────────────────────

GITHUB_SOURCES = [
    {
        "name": "speedyapply",
        "raw_url": "https://raw.githubusercontent.com/speedyapply/2026-SWE-College-Jobs/main/NEW_GRAD_USA.md",
        "format": "markdown-html",
    },
    {
        "name": "SimplifyJobs",
        "raw_url": "https://raw.githubusercontent.com/SimplifyJobs/New-Grad-Positions/dev/README.md",
        "format": "html-table",
    },
    {
        "name": "jobright-ai",
        "raw_url": "https://raw.githubusercontent.com/jobright-ai/2026-Software-Engineer-New-Grad/master/README.md",
        "format": "markdown-html",
    },
    {
        "name": "cvrve",
        "raw_url": "https://raw.githubusercontent.com/cvrve/New-Grad/main/README.md",
        "format": "markdown-html",
    },
]

# Companies that publish open job boards via the Greenhouse API.
# Add or remove slugs freely — the API is public and requires no key.
GREENHOUSE_COMPANIES = [
    "stripe", "airbnb", "figma", "discord", "reddit",
    "cloudflare", "databricks", "robinhood", "coinbase", "brex",
    "gusto", "lyft", "pinterest", "twitch", "dropbox",
    "hubspot", "asana", "okta", "mongodb",
    "amplitude", "lattice", "anthropic",
]

# Titles containing these words are skipped — they signal senior/management roles
# not suitable for a new-grad candidate.
SENIOR_TITLE_KEYWORDS = [
    "senior", "staff", "principal", "manager", "director", "lead",
    "head of", "vp ", "vice president", "distinguished", "fellow",
    "architect", "cto", "ceo", "ciso",
]

GEMINI_MODEL = "gemini-2.0-flash"
GROQ_MODEL   = "llama-3.3-70b-versatile"
MIN_SCORE = 60          # jobs below this score are skipped in output
MAX_JOBS = 200          # safety cap — process at most this many jobs
REQUEST_DELAY = 0.5     # seconds between HTTP requests (be polite)
REQUEST_TIMEOUT = 10    # seconds per HTTP request

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )
}

PROFILE_DOC_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "profile_data.md")

# ── Helpers ───────────────────────────────────────────────────────────────────

def read_docx(path: str) -> str:
    """Extract all text from a .docx file."""
    doc = Document(path)
    paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
    # also grab tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell.text.strip():
                    paragraphs.append(cell.text.strip())
    return "\n".join(paragraphs)


def summarize_profile(provider: str, client, raw_text: str) -> str:
    """
    One-time call: ask Gemini to distill the raw .docx text into a clean,
    scoring-optimized summary. Strips meta-instructions, interview coaching
    notes, and conditional resume rules — keeps only factual career info.
    """
    prompt = f"""You are preparing a candidate profile for automated job matching against software engineering job descriptions.

Below is a detailed career profile document. Extract ONLY the factual career information relevant to job matching.

EXCLUDE:
- Meta-instructions or notes about how to use the document
- Interview coaching advice
- Conditional rules about resume writing (e.g. "only list X if JD requires it")
- "Do not claim" / "not skilled in" disclaimers (just omit those skills entirely)
- Internal notes about what to say vs not say

INCLUDE with all details:
- Authorization status (visa type, OPT start date, STEM extension eligibility, relocation openness)
- Education (degrees, schools, GPAs, graduation dates)
- Target roles (what the candidate is open to)
- Work experience (company, title, dates, technologies used, key achievements and metrics)
- Personal projects (tech stack, what it does, quantified outcomes)
- Technical skills (verified skills only — languages, frameworks, cloud services, databases, tools)
- Certifications

Output a clean structured summary. Keep all real metrics and tech stack details — they are critical for accurate job matching.

PROFILE DOCUMENT:
{raw_text}"""

    try:
        return call_llm(provider, client, prompt, max_tokens=2048)
    except Exception as e:
        print(f"  [warn] Profile summarization failed: {e}. Using raw text.")
        return raw_text


def load_profile_doc() -> str:
    """Load the persistent profile doc. Errors out with guidance if it doesn't exist yet."""
    if not os.path.exists(PROFILE_DOC_PATH):
        print(f"Error: no profile doc found at {PROFILE_DOC_PATH}")
        print("Run this first: python job_search.py --init path/to/your_profile.docx")
        sys.exit(1)
    with open(PROFILE_DOC_PATH, "r", encoding="utf-8") as f:
        return f.read()


def save_profile_doc(text: str) -> None:
    """Write the persistent profile doc to disk."""
    with open(PROFILE_DOC_PATH, "w", encoding="utf-8") as f:
        f.write(text)


def update_profile_doc(provider: str, client, current_doc: str, new_fact: str) -> str:
    """
    Merge a new fact (e.g. "I learned Kubernetes", "I completed an AWS cert")
    into the existing structured profile doc. Keeps everything else unchanged
    unless the new fact directly contradicts something already there.
    """
    prompt = f"""You maintain a structured career profile document used for automated job matching.

Below is the CURRENT profile document, followed by a NEW FACT the candidate wants added.

Update the document to incorporate the new fact:
- Add it to the most relevant existing section (e.g. a new skill goes under Technical Skills, a new project goes under Projects)
- If the new fact contradicts something already in the document, update that part rather than duplicating it
- Keep everything else in the document unchanged
- Keep the same structured format (same section headers) as the current document
- Do not add commentary, meta-notes, or explanations — output only the updated profile document

CURRENT PROFILE DOCUMENT:
{current_doc}

NEW FACT TO ADD:
{new_fact}

Output the complete updated profile document:"""

    return call_llm(provider, client, prompt, max_tokens=2048)


def fetch_text(url: str) -> str:
    """Fetch URL and return raw text. Returns empty string on failure."""
    try:
        r = requests.get(url, headers=HEADERS, timeout=REQUEST_TIMEOUT)
        r.raise_for_status()
        return r.text
    except Exception as e:
        print(f"  [warn] Could not fetch {url}: {e}")
        return ""


def looks_js_rendered(html: str) -> bool:
    """Return True if the page HTML looks like a JS-rendered shell with no real content."""
    markers = [
        "__NEXT_DATA__", "ng-version", "data-reactroot",
        '<div id="root">', '<div id="app">', "window.__INITIAL_STATE__",
        "__nuxt__", "data-v-app",
    ]
    return any(m in html for m in markers)


def fetch_with_browser(url: str) -> str:
    """Fetch a URL using a headless Chromium browser and return the rendered HTML."""
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page = browser.new_page()
            page.goto(url, timeout=20000, wait_until="domcontentloaded")
            page.wait_for_timeout(2500)  # let JS render
            html = page.content()
            browser.close()
            return html
    except Exception as e:
        print(f"  [warn] Browser fetch failed for {url}: {e}")
        return ""


def _parse_html_to_text(html: str) -> str:
    """Shared BeautifulSoup extraction used by both fetch paths."""
    soup = BeautifulSoup(html, "html.parser")
    for tag in soup(["script", "style", "nav", "header", "footer", "aside"]):
        tag.decompose()
    text = soup.get_text(separator="\n")
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    return "\n".join(lines)


def extract_job_description(url: str) -> str:
    """
    Fetch a job posting and extract visible text.
    Tries plain HTTP first; falls back to headless Chromium if the page
    looks JS-rendered (empty shell returned by requests).
    """
    html = fetch_text(url)
    if not html:
        return ""

    text = _parse_html_to_text(html)

    if len(text) < 200 and looks_js_rendered(html):
        if PLAYWRIGHT_AVAILABLE:
            print("  [browser] JS-rendered page detected, using headless browser...")
            html = fetch_with_browser(url)
            if html:
                text = _parse_html_to_text(html)
        else:
            print("  [warn] JS-rendered page detected but Playwright not installed.")
            print("         Run: pip install playwright && playwright install chromium")

    return text[:8000]


def _extract_urls(text: str) -> list[str]:
    """Extract all https URLs from text (markdown or HTML link format)."""
    urls = []
    urls += re.findall(r'href=["\']?(https?://[^"\'>\s]+)', text)
    urls += [u for _, u in re.findall(r"\[([^\]]*)\]\((https?://[^)]+)\)", text)]
    return urls


def _clean_text(text: str) -> str:
    """Strip HTML and markdown link syntax, keep visible text only."""
    text = re.sub(r"<a[^>]*>(.*?)</a>", r"\1", text, flags=re.DOTALL)
    text = re.sub(r"</?(?:strong|em|b|i|div|td|tr|th|tbody|thead)[^>]*>", "", text)
    text = re.sub(r"<[^>]+>", "", text)
    text = re.sub(r"\[([^\]]*)\]\([^)]*\)", r"\1", text)
    text = re.sub(r":[a-z_]+:", "", text)  # :emoji: shortcodes
    text = re.sub(r"\*+([^*]+)\*+", r"\1", text)  # **bold** and *italic*
    return text.strip()


def _pick_apply_url(urls: list[str]) -> str:
    """
    From a list of URLs in a job row, pick the best apply link.
    Prefer non-github, non-company-homepage URLs (i.e. ATS / job board links).
    Simplify links are good apply proxies if nothing better is found.
    """
    ats_keywords = (
        "myworkday", "greenhouse", "lever", "ashby", "smartrecruiters",
        "icims", "jobvite", "workday", "taleo", "ultipro", "bamboohr",
        "amazon.jobs", "careers.", "jobs.", "job-boards", "apply",
        "simplify.jobs/p/",  # simplify direct apply pages
    )
    # tier 1: real ATS links
    for url in urls:
        if any(kw in url.lower() for kw in ats_keywords):
            return url
    # tier 2: any non-github link
    for url in urls:
        if "github.com" not in url and "simplify.jobs/c/" not in url:
            return url
    # tier 3: simplify company page (still useful)
    for url in urls:
        if "simplify.jobs" in url:
            return url
    return urls[-1] if urls else ""


def is_posted_today(posted: str) -> bool:
    """
    Decide whether a job's 'posted' string means it went up within the last 24 hours.
    Handles two formats seen across the GitHub source repos:
      - relative age: "0d", "5h", "30m", "new" -> today
      - absolute date: "Jun 30" -> compare month/day to today's date
    Unparseable/missing values are treated as NOT today (conservative).
    """
    if not posted:
        return False
    p = posted.strip().lower()

    if p in ("new", "today"):
        return True

    m = re.match(r"^(\d+)\s*([hmd])\b", p)
    if m:
        value, unit = int(m.group(1)), m.group(2)
        if unit == "d":
            return value == 0
        return True  # hours or minutes ago is always within 24h

    # ISO timestamp from Greenhouse: "2026-06-30T17:05:44-04:00"
    try:
        parsed = datetime.fromisoformat(posted.strip())
        today = datetime.now()
        return parsed.date() == today.date()
    except ValueError:
        pass

    # absolute date like "Jun 30"
    try:
        parsed = datetime.strptime(posted.strip(), "%b %d")
        today = datetime.now()
        return parsed.month == today.month and parsed.day == today.day
    except ValueError:
        return False


def parse_jobs(content: str, fmt: str) -> list[dict]:
    """
    Dispatch to the right parser based on fmt:
      'markdown-html' — pipe-delimited markdown table with optional HTML in cells
      'html-table'    — raw HTML <table><tr><td> embedded in the file
    """
    if fmt == "html-table":
        return _parse_html_table(content)
    return _parse_markdown_table(content)


def _parse_markdown_table(markdown: str) -> list[dict]:
    jobs = []
    for line in markdown.splitlines():
        if not line.startswith("|"):
            continue
        if re.match(r"^\|[\s\-:|]+\|$", line):
            continue
        cells = [c.strip() for c in line.split("|")]
        cells = [c for c in cells if c]
        if len(cells) < 3:
            continue

        all_urls = []
        for cell in cells:
            all_urls.extend(_extract_urls(cell))

        if not all_urls:
            continue

        apply_url = _pick_apply_url(all_urls)
        if not apply_url:
            continue

        company = _clean_text(cells[0])
        role = _clean_text(cells[1]) if len(cells) > 1 else ""
        location = _clean_text(cells[2]) if len(cells) > 2 else ""
        posted = _clean_text(cells[-1]) if len(cells) > 3 else ""

        if company.lower() in ("company", "name", ""):
            continue

        jobs.append({"company": company, "role": role, "location": location, "url": apply_url, "posted": posted})

    return jobs


def _parse_html_table(content: str) -> list[dict]:
    """Parse jobs from HTML <tr><td> tables embedded in the file."""
    soup = BeautifulSoup(content, "html.parser")
    jobs = []
    for row in soup.find_all("tr"):
        cells = row.find_all("td")
        if len(cells) < 3:
            continue

        # column 0 = company, 1 = role, 2 = location, 3 = apply link(s), last = age
        company = _clean_text(cells[0].get_text())
        role = _clean_text(cells[1].get_text())
        location = _clean_text(cells[2].get_text())
        posted = _clean_text(cells[-1].get_text()) if len(cells) > 3 else ""

        if company.lower() in ("company", "name", ""):
            continue

        # gather all URLs from the row
        all_urls = _extract_urls(str(row))
        if not all_urls:
            continue

        apply_url = _pick_apply_url(all_urls)
        if not apply_url:
            continue

        jobs.append({"company": company, "role": role, "location": location, "url": apply_url, "posted": posted})

    return jobs


def fetch_json(url: str) -> dict:
    """Fetch URL and return parsed JSON. Returns {} on failure."""
    try:
        r = requests.get(url, headers=HEADERS, timeout=REQUEST_TIMEOUT)
        r.raise_for_status()
        return r.json()
    except Exception as e:
        print(f"  [warn] Could not fetch JSON from {url}: {e}")
        return {}


SWE_TITLE_KEYWORDS = [
    "engineer", "developer", "software", "sde", "swe", "backend", "frontend",
    "full stack", "fullstack", "infrastructure", "security engineer",
    "machine learning", "data engineer", "devops", "cloud engineer",
    "systems", "compiler", "firmware", "embedded", "mobile", "android", "ios",
]

NON_SWE_TITLE_KEYWORDS = [
    "account executive", "account manager", "sales", "recruiter", "recruiting",
    "finance", "legal", "marketing", "design", "customer success", "customer support",
    "operations", "analyst", "consultant", "business development", "partnership",
    "product manager", "program manager", "project manager", "hr ", "human resources",
]

def is_entry_level_title(title: str) -> bool:
    """Return True if the title is an entry-level SWE role (not senior/management, not non-SWE)."""
    t = title.lower()
    if any(kw in t for kw in SENIOR_TITLE_KEYWORDS):
        return False
    if any(kw in t for kw in NON_SWE_TITLE_KEYWORDS):
        return False
    return any(kw in t for kw in SWE_TITLE_KEYWORDS)


def fetch_greenhouse_jobs(companies: list[str]) -> list[dict]:
    """
    Fetch entry-level SWE jobs from Greenhouse company boards.
    The Greenhouse API returns the full job description in the `content` field,
    so no separate page scrape is needed for these jobs.
    """
    import html as html_module
    jobs = []
    for slug in companies:
        url = f"https://boards-api.greenhouse.io/v1/boards/{slug}/jobs?content=true"
        data = fetch_json(url)
        for job in data.get("jobs", []):
            title = job.get("title", "")
            if not is_entry_level_title(title):
                continue
            # extract plain text from HTML job description
            raw_html = html_module.unescape(job.get("content", ""))
            description = _parse_html_to_text(raw_html)[:8000]
            jobs.append({
                "company": job.get("company_name", slug),
                "role": title,
                "location": job.get("location", {}).get("name", ""),
                "url": job.get("absolute_url", ""),
                "posted": job.get("updated_at", ""),
                "description": description,
                "source": "greenhouse",
            })
        time.sleep(REQUEST_DELAY)
    return jobs


def score_job(provider: str, client, profile: str, job: dict, description: str) -> dict:
    """
    Ask Gemini to score the job match 0-100 and give a short reason.
    Returns dict with 'score' (int) and 'reason' (str).
    """
    prompt = f"""You are a career matching engine. Score how well this candidate matches this job on a scale of 0 to 100.

CANDIDATE PROFILE:
{profile}

JOB DETAILS:
Company: {job['company']}
Role: {job['role']}
Location: {job['location']}

JOB DESCRIPTION:
{description if description else "(no description available — score based on job title and role only)"}

SCORING CRITERIA (weight in order):
1. Tech stack overlap — languages, frameworks, cloud platforms, databases the candidate actually has
2. Role level fit — candidate is a new grad (MS graduating Aug 2026, ~3 years prior industry experience as SWE) — penalize senior/staff roles, reward new grad / entry-level / 0-2 YOE roles
3. Domain fit — backend engineering, distributed systems, cloud application development, APIs
4. Authorization — candidate is on F-1 OPT (no sponsorship needed for first 3 years); penalize only if job explicitly says "no visa sponsorship" AND OPT is not accepted
5. Location — candidate is open to anywhere in the US; penalize only if job is outside the US or explicitly remote-excluded

PENALIZE HEAVILY (score below 35) if the job primarily requires:
- Terraform / Infrastructure as Code
- Kubernetes orchestration
- Cloud networking (VPC/VNet design, Private Link, DNS)
- IAM/Entra ID platform engineering
- Hardware or embedded systems

Respond with ONLY a JSON object — no markdown, no text outside the JSON:
{{"score": <integer 0-100>, "reason": "<one concise sentence: the strongest match signal or the key mismatch>"}}"""

    try:
        text = call_llm(provider, client, prompt, max_tokens=256)

        # strip markdown code fences if present
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```$", "", text)

        import json
        match = re.search(r"\{.*\}", text, re.DOTALL)
        if match:
            data = json.loads(match.group())
            return {
                "score": int(data.get("score", 0)),
                "reason": str(data.get("reason", "")),
            }
    except Exception as e:
        print(f"  [warn] Scoring failed for {job['company']} – {job['role']}: {e}")

    return {"score": 0, "reason": "scoring error"}


# ── LLM provider ──────────────────────────────────────────────────────────────

def call_llm(provider: str, client, prompt: str, max_tokens: int = 1024) -> str:
    """Unified LLM call — works with either Gemini or Groq client."""
    if provider == "gemini":
        response = client.models.generate_content(model=GEMINI_MODEL, contents=prompt)
        return response.text.strip()
    else:
        response = client.chat.completions.create(
            model=GROQ_MODEL,
            messages=[{"role": "user", "content": prompt}],
            max_tokens=max_tokens,
        )
        return response.choices[0].message.content.strip()


def get_llm_client() -> tuple:
    """
    Auto-detect which LLM provider to use.
    Checks GEMINI_API_KEY first (primary), falls back to GROQ_API_KEY.
    Returns (provider_name, client_object).
    """
    gemini_key = os.environ.get("GEMINI_API_KEY", "")
    groq_key = os.environ.get("GROQ_API_KEY", "")

    if gemini_key:
        from google import genai
        print("Using Google Gemini API")
        return ("gemini", genai.Client(api_key=gemini_key))

    if groq_key:
        from groq import Groq
        print("Using Groq API (local fallback)")
        return ("groq", Groq(api_key=groq_key))

    print("Error: No API key found. Set one of:")
    print("  Primary:  export GEMINI_API_KEY='your-key'  # https://aistudio.google.com/app/apikey")
    print("  Fallback: export GROQ_API_KEY='your-key'    # https://console.groq.com/keys")
    sys.exit(1)


# ── Main ──────────────────────────────────────────────────────────────────────


def main():
    parser = argparse.ArgumentParser(description="Job Search Automation")
    parser.add_argument("--init", metavar="PROFILE_DOCX", help="Build profile_data.md from a .docx resume/profile")
    parser.add_argument("--update", metavar="FACT", help="Merge a new fact into profile_data.md")
    parser.add_argument("--show", action="store_true", help="Print the current profile_data.md")
    parser.add_argument("--today", action="store_true", help="Only include jobs posted within the last 24 hours")
    args = parser.parse_args()

    if args.show:
        print(load_profile_doc())
        return

    if args.init:
        if not os.path.exists(args.init):
            print(f"Error: file not found: {args.init}")
            sys.exit(1)
        print(f"Reading profile from: {args.init}")
        raw_profile = read_docx(args.init)
        if not raw_profile.strip():
            print("Error: could not extract text from the .docx file.")
            sys.exit(1)
        print(f"  Profile loaded ({len(raw_profile)} chars)")

        provider, client = get_llm_client()
        print("\nSummarizing profile...")
        profile_text = summarize_profile(provider, client, raw_profile)
        save_profile_doc(profile_text)

        print("\n" + "─" * 60)
        print(f"Saved profile doc to: {PROFILE_DOC_PATH}")
        print("─" * 60)
        print(profile_text)
        print("─" * 60)
        return

    if args.update:
        provider, client = get_llm_client()
        current_doc = load_profile_doc()
        print("Updating profile doc with new fact...")
        updated_doc = update_profile_doc(provider, client, current_doc, args.update)
        save_profile_doc(updated_doc)

        print("\n" + "─" * 60)
        print(f"Updated profile doc saved to: {PROFILE_DOC_PATH}")
        print("─" * 60)
        print(updated_doc)
        print("─" * 60)
        return

    # ── Normal run: use the existing profile doc directly ──
    if not PLAYWRIGHT_AVAILABLE:
        print("Tip: install Playwright to scrape JS-rendered job pages:")
        print("     pip install playwright && playwright install chromium")
        print()

    profile_text = load_profile_doc()
    provider, client = get_llm_client()

    # ── Fetch jobs from all sources ──
    all_jobs = []
    for source in GITHUB_SOURCES:
        print(f"\nFetching jobs from {source['name']}...")
        md = fetch_text(source["raw_url"])
        if not md:
            print(f"  [skip] Could not fetch {source['name']}")
            continue
        jobs = parse_jobs(md, source.get("format", "markdown-html"))
        print(f"  Found {len(jobs)} open positions")
        for job in jobs:
            job["source"] = source["name"]
        all_jobs.extend(jobs)
        time.sleep(REQUEST_DELAY)

    print(f"\nFetching jobs from Greenhouse ({len(GREENHOUSE_COMPANIES)} companies)...")
    gh_jobs = fetch_greenhouse_jobs(GREENHOUSE_COMPANIES)
    print(f"  Found {len(gh_jobs)} entry-level positions across Greenhouse boards")
    all_jobs.extend(gh_jobs)

    # deduplicate by URL
    seen_urls = set()
    unique_jobs = []
    for job in all_jobs:
        if job["url"] not in seen_urls:
            seen_urls.add(job["url"])
            unique_jobs.append(job)

    print(f"\nTotal unique jobs found: {len(unique_jobs)}")

    if args.today:
        before = len(unique_jobs)
        unique_jobs = [j for j in unique_jobs if is_posted_today(j.get("posted", ""))]
        print(f"  Filtered to jobs posted in the last 24 hours: {len(unique_jobs)} (of {before})")

    print(f"\nTotal unique jobs to process: {len(unique_jobs)}")
    if len(unique_jobs) > MAX_JOBS:
        print(f"Capping at {MAX_JOBS} jobs. Edit MAX_JOBS in the script to change this.")
        unique_jobs = unique_jobs[:MAX_JOBS]

    # ── Score each job ──
    results = []

    for i, job in enumerate(unique_jobs, 1):
        print(f"\n[{i}/{len(unique_jobs)}] {job['company']} — {job['role']}")
        # Greenhouse jobs already carry the description from the API — skip page scrape
        if job.get("description"):
            print(f"  Description pre-loaded from Greenhouse API")
            description = job["description"]
        else:
            print(f"  Fetching job description from: {job['url']}")
            description = extract_job_description(job["url"])
            time.sleep(REQUEST_DELAY)

        print(f"  Scoring...")
        scored = score_job(provider, client, profile_text, job, description)
        job["score"] = scored["score"]
        job["reason"] = scored["reason"]
        print(f"  Score: {job['score']}/100 — {job['reason']}")

        if job["score"] >= MIN_SCORE:
            results.append(job)

    # ── Write CSV ──
    results.sort(key=lambda x: x["score"], reverse=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = f"jobs_{timestamp}.csv"

    fieldnames = ["score", "company", "role", "location", "posted", "reason", "source", "url"]
    with open(output_file, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for row in results:
            writer.writerow({k: row.get(k, "") for k in fieldnames})

    print(f"\n{'='*60}")
    print(f"Done! {len(results)} matching jobs (score >= {MIN_SCORE}) saved to:")
    print(f"  {os.path.abspath(output_file)}")
    print(f"{'='*60}")
    if not results:
        print("No jobs met the minimum score threshold. Try lowering MIN_SCORE in the script.")


if __name__ == "__main__":
    main()
