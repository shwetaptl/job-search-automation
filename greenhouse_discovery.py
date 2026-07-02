# Greenhouse Slug Discovery + Integration Agent
# ─────────────────────────────────────────────
# PURPOSE: Test which H1B-sponsor companies use Greenhouse ATS,
# find their working API slugs, and update job_search.py automatically.
#
# HOW TO USE:
#   cd /Users/shweta/Downloads/Projects/Automation_Job_Search
#   python greenhouse_discovery.py
#
# REQUIREMENTS: pip install requests

import requests
import time
import re
import os

BASE_URL = "https://boards-api.greenhouse.io/v1/boards/{}/jobs"
JOB_SEARCH_PATH = "job_search.py"
REQUEST_DELAY = 0.4  # seconds between requests — be respectful

# ─────────────────────────────────────────────────────────────────────────────
# ALL H1B SPONSOR COMPANIES + SLUG CANDIDATES
# Multiple slugs per company — the script tests each until one works
# ─────────────────────────────────────────────────────────────────────────────
SLUG_CANDIDATES = {

    # ── TIER 1: Healthcare (Meditab background directly relevant) ──────────
    "Hims & Hers":              ["hims", "himsandhers", "himshers"],
    "Availity":                 ["availity"],
    "Waystar":                  ["waystar", "waystarhealth"],
    "Inovalon":                 ["inovalon"],
    "Arcadia":                  ["arcadia", "arcadiasolutions"],
    "Netsmart Technologies":    ["netsmart", "netsmartechnologies"],
    "athenahealth":             ["athenahealth", "athena"],
    "Modernizing Medicine":     ["modernizingmedicine", "modmed"],
    "PointClickCare":           ["pointclickcare", "pcc"],
    "Veeva Systems":            ["veeva", "veevasystems"],
    "ResMed":                   ["resmed"],
    "Change Healthcare":        ["changehealthcare", "changeinc"],
    "Guardian Health":          ["guardianhealth", "guardian"],
    "GE HealthCare":            ["gehealthcare", "gehealthcaretechnologies", "ge"],
    "Philips":                  ["philips", "philipshealthtech"],
    "Medtronic":                ["medtronic"],
    "Allscripts":               ["allscripts"],
    "Cerner (Oracle Health)":   ["cerner", "oraclehealth", "oracle"],
    "Epic Systems":             ["epic", "epicsystems"],

    # ── TIER 1: Fintech ────────────────────────────────────────────────────
    "Brex":                     ["brex"],
    "Robinhood":                ["robinhood"],
    "Coinbase":                 ["coinbase"],
    "SoFi":                     ["sofi"],
    "Affirm":                   ["affirm"],
    "Marqeta":                  ["marqeta"],
    "Chime":                    ["chime"],
    "Plaid":                    ["plaid"],
    "Square / Block":           ["block", "square", "squareinc", "blockorg"],
    "FIS Global":               ["fis", "fisglobal", "fisv"],
    "Fiserv":                   ["fiserv"],
    "Jack Henry":               ["jackhenry", "jackhenryassociates"],
    "SS&C Technologies":        ["ssc", "ssctechnologies", "ssctech"],
    "Broadridge":               ["broadridge", "broadridgefinancial"],
    "Morningstar":              ["morningstar"],
    "Envestnet":                ["envestnet"],
    "Tradeweb":                 ["tradeweb"],
    "Solera":                   ["solera"],

    # ── TIER 1: Enterprise / ERP ───────────────────────────────────────────
    "Tyler Technologies":       ["tylertech", "tylertechnologies"],
    "Yardi Systems":            ["yardi", "yardisystems"],
    "RealPage":                 ["realpage"],
    "MRI Software":             ["mrisoftware", "mri"],
    "Roper Technologies":       ["ropertech", "ropertechnologies"],
    "Trimble":                  ["trimble"],
    "Bentley Systems":          ["bentley", "bentleysystems"],
    "PTC Inc":                  ["ptc"],
    "Verint":                   ["verint"],
    "NICE Systems":             ["nice", "nicesystems"],
    "Pegasystems":              ["pega", "pegasystems"],
    "OpenText":                 ["opentext"],
    "Precisely":                ["precisely"],
    "Epicor":                   ["epicor"],
    "Infor":                    ["infor"],
    "Domo":                     ["domo"],
    "Qualtrics":                ["qualtrics"],

    # ── TIER 1: Insurance Tech ─────────────────────────────────────────────
    "Applied Systems (EZLynx)": ["appliedsystems", "ezlynx", "appliedsystemsinc"],
    "Guidewire":                ["guidewire", "guidewiresoftware"],
    "Vertafore":                ["vertafore"],
    "Duck Creek":               ["duckcreek", "duckcreektechnologies"],
    "Majesco":                  ["majesco"],
    "Sapiens":                  ["sapiens", "sapiensintl"],

    # ── TIER 2: SaaS / Cloud ───────────────────────────────────────────────
    "ServiceNow":               ["servicenow"],
    "Workday":                  ["workday"],
    "Atlassian":                ["atlassian"],
    "HubSpot":                  ["hubspot"],
    "Okta":                     ["okta"],
    "MongoDB":                  ["mongodb"],
    "Elastic":                  ["elastic", "elasticnv"],
    "Splunk":                   ["splunk"],
    "Datadog":                  ["datadog"],
    "Twilio":                   ["twilio"],
    "Cloudflare":               ["cloudflare"],
    "CrowdStrike":              ["crowdstrike"],
    "Zscaler":                  ["zscaler"],
    "Databricks":               ["databricks"],
    "Snowflake":                ["snowflake"],
    "Dynatrace":                ["dynatrace"],
    "New Relic":                ["newrelic"],
    "PagerDuty":                ["pagerduty"],
    "Sumo Logic":               ["sumologic"],
    "Grafana Labs":             ["grafanalabs", "grafana"],
    "Harness":                  ["harness"],
    "Sisense":                  ["sisense"],
    "Amplitude":                ["amplitude"],
    "Lattice":                  ["lattice"],
    "Discord":                  ["discord"],
    "Reddit":                   ["reddit"],
    "Asana":                    ["asana"],
    "Flexport":                 ["flexport"],
    "Procore":                  ["procore"],
    "Autodesk":                 ["autodesk"],
    "Ansys":                    ["ansys"],
    "Instructure (Canvas)":     ["instructure"],
    "PowerSchool":              ["powerschool"],

    # ── TIER 2: AI / Data / Analytics ─────────────────────────────────────
    "Fractal Analytics":        ["fractal", "fractalanalytics"],
    "Quantiphi":                ["quantiphi"],
    "ZS Associates":            ["zs", "zsassociates"],
    "Nagarro":                  ["nagarro"],
    "Informatica":              ["informatica"],

    # ── TIER 3: Cybersecurity ──────────────────────────────────────────────
    "Palo Alto Networks":       ["paloaltonetworks", "paloalto"],
    "Fortinet":                 ["fortinet"],
    "SentinelOne":              ["sentinelone"],
    "Rapid7":                   ["rapid7"],
    "Qualys":                   ["qualys"],
    "Ping Identity":            ["pingidentity", "ping"],
    "SailPoint":                ["sailpoint"],
    "OneSpan":                  ["onespan"],

    # ── TIER 3: Consulting / Direct Hire ──────────────────────────────────
    "Slalom":                   ["slalom"],
    "Thoughtworks":             ["thoughtworks"],
    "West Monroe":              ["westmonroe", "westmonroepartners"],
    "Publicis Sapient":         ["publicissapient", "sapient"],
    "EPAM Systems":             ["epam", "epamsystems"],
    "Capgemini":                ["capgemini"],
    "LTIMindtree":              ["ltimindtree", "lti", "mindtree"],
    "Infosys":                  ["infosys"],
    "Wipro":                    ["wipro"],
    "HCL Technologies":         ["hcltech", "hcltechnologies", "hcl"],
    "Cognizant":                ["cognizant"],
    "TCS":                      ["tcs", "tataconsultancyservices"],

    # ── Already in job_search.py (re-verify these still work) ─────────────
    "Stripe":                   ["stripe"],
    "Airbnb":                   ["airbnb"],
    "Figma":                    ["figma"],
    "Lyft":                     ["lyft"],
    "Pinterest":                ["pinterest"],
    "Twitch":                   ["twitch"],
    "Dropbox":                  ["dropbox"],
    "Gusto":                    ["gusto"],
    "Anthropic":                ["anthropic"],
}

# ─────────────────────────────────────────────────────────────────────────────
# STEP 1: TEST ALL SLUGS
# ─────────────────────────────────────────────────────────────────────────────
def test_all_slugs():
    confirmed = {}
    not_found = []
    errors = []

    print("=" * 65)
    print("STEP 1: Testing Greenhouse API slugs for all companies")
    print("=" * 65)
    print()

    for company, slugs in SLUG_CANDIDATES.items():
        found = False
        for slug in slugs:
            try:
                url = BASE_URL.format(slug)
                r = requests.get(url, timeout=10, headers={"User-Agent": "Mozilla/5.0"})
                if r.status_code == 200:
                    data = r.json()
                    job_count = len(data.get("jobs", []))
                    confirmed[company] = {"slug": slug, "jobs": job_count, "url": url}
                    print(f"  OK  {company:<35} slug: '{slug}' -> {job_count} jobs")
                    found = True
                    break
                elif r.status_code == 404:
                    pass  # try next slug
            except requests.exceptions.Timeout:
                errors.append(f"{company} ({slug}): timeout")
            except Exception as e:
                errors.append(f"{company} ({slug}): {str(e)[:50]}")
            time.sleep(REQUEST_DELAY)

        if not found:
            not_found.append(company)
            print(f"  --  {company:<35} -> not on Greenhouse")

    return confirmed, not_found, errors


# ─────────────────────────────────────────────────────────────────────────────
# STEP 2: PRINT SUMMARY + SLUG LIST
# ─────────────────────────────────────────────────────────────────────────────
def print_summary(confirmed, not_found, errors):
    print()
    print("=" * 65)
    print(f"RESULTS SUMMARY")
    print("=" * 65)
    print(f"  Confirmed on Greenhouse: {len(confirmed)} companies")
    print(f"  Not on Greenhouse:       {len(not_found)} companies")
    if errors:
        print(f"  Errors (retry manually): {len(errors)}")
    print()

    print("=" * 65)
    print("COPY-PASTE THIS INTO job_search.py")
    print("Replace your existing GREENHOUSE_COMPANIES list with this:")
    print("=" * 65)
    print()
    print("GREENHOUSE_COMPANIES = [")
    for company, data in confirmed.items():
        print(f'    "{data["slug"]}",  # {company} ({data["jobs"]} jobs)')
    print("]")
    print()

    workday_companies = [
        "Epic Systems", "Medtronic", "GE HealthCare", "Philips", "ServiceNow",
        "Snowflake", "Splunk", "Datadog", "Autodesk", "Fiserv", "FIS Global",
        "Tyler Technologies", "Trimble", "Broadridge", "CrowdStrike", "Zscaler",
        "Palo Alto Networks", "Workday", "Infor", "Epicor",
    ]
    lever_companies = [
        "Harness", "Grafana Labs", "SentinelOne", "Sumo Logic",
        "Applied Systems (EZLynx)", "Waystar", "New Relic", "Sisense",
    ]
    taleo_companies = [
        "Cerner (Oracle Health)", "Allscripts", "Infosys", "TCS", "Wipro",
        "HCL Technologies", "Cognizant", "Capgemini",
    ]
    smart_recruiters = [
        "Nagarro", "LTIMindtree", "EPAM Systems", "Thoughtworks",
    ]
    icims_companies = [
        "Yardi Systems", "RealPage", "MRI Software", "Jack Henry",
        "Vertafore", "Duck Creek", "SS&C Technologies",
    ]

    print("=" * 65)
    print("COMPANIES NOT ON GREENHOUSE — ALTERNATIVE ATS")
    print("=" * 65)

    print("\n  WORKDAY ATS (Playwright fallback can handle these):")
    for c in not_found:
        if c in workday_companies:
            print(f"     - {c}")

    print("\n  LEVER ATS (Playwright fallback can handle these):")
    for c in not_found:
        if c in lever_companies:
            print(f"     - {c}")

    print("\n  TALEO/ORACLE ATS:")
    for c in not_found:
        if c in taleo_companies:
            print(f"     - {c}")

    print("\n  SMARTRECRUITERS:")
    for c in not_found:
        if c in smart_recruiters:
            print(f"     - {c}")

    print("\n  iCIMS:")
    for c in not_found:
        if c in icims_companies:
            print(f"     - {c}")

    remaining = [c for c in not_found if c not in
                 workday_companies + lever_companies + taleo_companies +
                 smart_recruiters + icims_companies]
    if remaining:
        print("\n  UNKNOWN ATS (check their career page directly):")
        for c in remaining:
            print(f"     - {c}")

    if errors:
        print("\n  ERRORS — retry these manually:")
        for e in errors:
            print(f"     - {e}")


# ─────────────────────────────────────────────────────────────────────────────
# STEP 3: AUTO-UPDATE job_search.py
# ─────────────────────────────────────────────────────────────────────────────
def update_job_search_py(confirmed, job_search_path=JOB_SEARCH_PATH):
    if not os.path.exists(job_search_path):
        print(f"\nCould not find {job_search_path} — skipping auto-update.")
        print("Copy the GREENHOUSE_COMPANIES list above manually.")
        return

    with open(job_search_path, "r") as f:
        content = f.read()

    new_list_lines = ["GREENHOUSE_COMPANIES = ["]
    for company, data in confirmed.items():
        new_list_lines.append(f'    "{data["slug"]}",  # {company} ({data["jobs"]} jobs)')
    new_list_lines.append("]")
    new_list = "\n".join(new_list_lines)

    pattern = r'GREENHOUSE_COMPANIES\s*=\s*\[.*?\]'
    if re.search(pattern, content, re.DOTALL):
        updated = re.sub(pattern, new_list, content, flags=re.DOTALL)
        with open(job_search_path, "w") as f:
            f.write(updated)
        print(f"\nAUTO-UPDATED: {job_search_path}")
        print(f"GREENHOUSE_COMPANIES now has {len(confirmed)} confirmed companies.")
    else:
        print(f"\nCould not find GREENHOUSE_COMPANIES in {job_search_path}.")
        print("Copy the list above and paste it manually.")


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print()
    print("Greenhouse Slug Discovery — Job Search Automation")
    print(f"Testing {len(SLUG_CANDIDATES)} companies against Greenhouse API")
    print()

    confirmed, not_found, errors = test_all_slugs()
    print_summary(confirmed, not_found, errors)

    print()
    update_choice = input("Auto-update job_search.py with confirmed slugs? (y/n): ").strip().lower()
    if update_choice == "y":
        update_job_search_py(confirmed)
    else:
        print("\nSkipped auto-update. Copy the GREENHOUSE_COMPANIES list above manually.")

    print("\nDone!")
