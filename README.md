# Job Search Automation Pipeline

An end-to-end Python automation pipeline that scrapes 3,700+ live software engineering job postings daily, scores each one against a personal career profile using the Google Gemini API, and outputs a ranked CSV of the best matches — eliminating hours of manual job hunting.

Built as a demonstration of practical AI integration in a real-world engineering workflow.

---

## The Problem

Manually searching for new-grad SWE jobs means:
- Checking 4+ GitHub repos and dozens of company career pages daily
- Opening each job link, reading the description, judging fit
- Doing this for hundreds of postings — taking hours every day

## The Solution

This pipeline automates the entire process:

1. Pulls job listings from 4 community-maintained GitHub repos + Greenhouse company APIs
2. Filters to jobs posted within the last 24 hours (optional)
3. Scrapes each job's full description — with JS-rendered page support via Playwright
4. Scores each job 0–100 against your career profile using Gemini LLM
5. Outputs a sorted CSV with only the best matches, ready to apply

---

## Architecture

```
┌─────────────────────────────────────────────────────────┐
│                    JOB SOURCES (5 total)                 │
│  GitHub Repos ×4          Greenhouse API ×22 companies   │
│  (markdown/HTML tables)   (JSON, pre-loaded descriptions)│
└────────────────────────┬────────────────────────────────┘
                         │ ~3,700 raw listings
                         ▼
              ┌──────────────────┐
              │  Parse + Dedup   │  Extract company, role,
              │                  │  location, URL, date posted
              └────────┬─────────┘
                       │ optional --today filter (last 24h)
                       ▼
              ┌──────────────────┐
              │  Fetch Job Desc  │  requests → BeautifulSoup
              │                  │  JS fallback → Playwright
              └────────┬─────────┘
                       │
                       ▼
              ┌──────────────────┐
              │  Score via LLM   │  Gemini API
              │                  │  profile + JD → score 0-100
              └────────┬─────────┘
                       │ filter score >= MIN_SCORE
                       ▼
              ┌──────────────────┐
              │   CSV Output     │  sorted by score (desc)
              │  jobs_DATE.csv   │  with apply links
              └──────────────────┘
```

---

## Key Features

- **Multi-source aggregation** — 4 GitHub repos + Greenhouse boards from 22 top tech companies (Stripe, Airbnb, Cloudflare, Databricks, Discord, Reddit, and more)
- **JS-rendered page support** — automatically detects JavaScript-heavy job pages and falls back to headless Chromium via Playwright
- **LLM-powered matching** — Gemini scores each job against your profile with explicit criteria: tech stack overlap, role level fit, domain fit, location/authorization
- **Persistent profile doc** — your career profile is summarized once and stored locally; update it with natural-language commands
- **24-hour filter** — `--today` flag filters to jobs posted in the last 24 hours across all sources
- **Smart deduplication** — same job appearing across multiple sources is processed only once
- **Greenhouse pre-loading** — Greenhouse API returns full job descriptions in the API response, skipping the page scrape entirely for those jobs

---

## Tech Stack

| Layer | Technology |
|---|---|
| Language | Python 3.11+ |
| LLM / Scoring | Google Gemini API (`google-genai`) |
| Web scraping | `requests` + `BeautifulSoup4` |
| JS rendering | `Playwright` (headless Chromium) |
| Resume parsing | `python-docx` |
| Job sources | GitHub raw markdown + Greenhouse REST API |
| Output | CSV (sortable in Excel / Google Sheets) |

---

## Setup

### Prerequisites
- Python 3.11+
- A free Google Gemini API key — get one at [aistudio.google.com/app/apikey](https://aistudio.google.com/app/apikey)

### Install

```bash
git clone https://github.com/shwetaptl/job-search-automation.git
cd job-search-automation

pip install -r requirements.txt
playwright install chromium
```

### Configure API key

```bash
export GEMINI_API_KEY="your-key-here"

# To make it permanent:
echo 'export GEMINI_API_KEY="your-key-here"' >> ~/.zshrc
source ~/.zshrc
```

### Build your profile (one-time)

```bash
python job_search.py --init path/to/your_resume.docx
```

This reads your `.docx` resume, uses Gemini to extract a clean scoring-optimized summary, and saves it as `profile_data.md`. This file is gitignored — it stays local only.

---

## Usage

### Run the full job search
```bash
python job_search.py
```

### Only show jobs posted in the last 24 hours
```bash
python job_search.py --today
```

### Update your profile with new information
```bash
python job_search.py --update "Completed AWS Solutions Architect certification in July 2026"
```

Gemini merges the new fact into your stored profile doc. All future runs use the updated version.

### Check what profile the script is using
```bash
python job_search.py --show
```

You can also open and manually edit `profile_data.md` directly in any text editor.

---

## Job Sources

| Source | Type | Focus |
|---|---|---|
| [speedyapply/2026-SWE-College-Jobs](https://github.com/speedyapply/2026-SWE-College-Jobs) | GitHub markdown | New grad USA |
| [SimplifyJobs/New-Grad-Positions](https://github.com/SimplifyJobs/New-Grad-Positions) | GitHub HTML table | New grad, Simplify-tracked |
| [jobright-ai/2026-Software-Engineer-New-Grad](https://github.com/jobright-ai/2026-Software-Engineer-New-Grad) | GitHub markdown | New grad |
| [cvrve/New-Grad](https://github.com/cvrve/New-Grad) | GitHub markdown | New grad |
| Greenhouse API | REST JSON | 22 top tech companies direct |

**Greenhouse companies included:** Stripe, Airbnb, Figma, Discord, Reddit, Cloudflare, Databricks, Robinhood, Coinbase, Brex, Gusto, Lyft, Pinterest, Twitch, Dropbox, HubSpot, Asana, Okta, MongoDB, Amplitude, Lattice, Anthropic

---

## Output

The script writes a CSV file named `jobs_YYYYMMDD_HHMMSS.csv`:

| Column | Description |
|---|---|
| `score` | Match score 0–100 (sorted highest first) |
| `company` | Company name |
| `role` | Job title |
| `location` | Job location |
| `posted` | When the job was posted (age or date) |
| `reason` | One-sentence explanation of the score |
| `source` | Which source the job came from |
| `url` | Direct apply link |

Open in Excel or Google Sheets — sort by `score`, filter by `source`, done.

---

## Configuration

Edit these constants at the top of `job_search.py`:

```python
MIN_SCORE = 60        # Only output jobs scoring >= this (0-100)
MAX_JOBS = 200        # Cap on jobs processed per run (safety limit)
REQUEST_DELAY = 0.5   # Seconds between HTTP requests (be polite)
```

To add or remove Greenhouse companies:
```python
GREENHOUSE_COMPANIES = [
    "stripe", "airbnb", "figma",
    # add any company slug from boards.greenhouse.io/{slug}
]
```

---

## Best Time to Run

Tuesday, Wednesday, and Thursday mornings (9–10 AM Eastern) are peak job posting days for US tech companies. Use `--today` on daily runs to catch freshly posted roles before the application pipeline fills up.

---

## Project Structure

```
job-search-automation/
├── job_search.py       # Main pipeline script
├── requirements.txt    # Python dependencies
├── README.md           # This file
├── .gitignore          # Excludes profile_data.md, CSVs, secrets
└── profile_data.md     # Your local profile (gitignored, never committed)
```

---

## Author

**Shweta Patel**
- GitHub: [github.com/shwetaptl](https://github.com/shwetaptl)
- LinkedIn: [linkedin.com/in/shwetaptl](https://linkedin.com/in/shwetaptl)
