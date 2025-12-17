import random
from datetime import datetime, timedelta
from pathlib import Path
import csv

random.seed(7)

ISSUES = [
    ("VPN Authentication", ["P2", "P3"]),
    ("Remote Access", ["P2", "P3"]),
    ("O365 Login", ["P2", "P3"]),
    ("Teams Audio", ["P3"]),
    ("Teams Video", ["P3"]),
    ("Conference Room AV", ["P2", "P3"]),
    ("Wi-Fi Drop", ["P2", "P3"]),
    ("Endpoint Patch/Update", ["P3", "P4"]),
    ("Printer", ["P4"]),
    ("Disk Full", ["P3", "P4"]),
    ("EDR Alert Investigation", ["P1", "P2"]),        # security-ish
    ("Phishing Report Triage", ["P2", "P3"]),         # security-ish
    ("MFA Device Reset", ["P2", "P3"]),               # security-ish
    ("Password Lockout", ["P3"]),
]

ROLES = ["Executive", "Trader", "Analyst", "Staff", "Engineer"]
DEVICES = ["Laptop", "Desktop"]
SITES = ["NYC-HQ", "NYC-Branch", "Remote"]
NETWORK = ["Wired", "WiFi", "VPN"]
VENDORS = ["Microsoft", "Cisco", "Zoom", "Okta", "CrowdStrike", "Dell", "HP", "Unknown"]

def gen_resolution_minutes(priority: str) -> int:
    # realistic-ish response distributions
    if priority == "P1":
        return max(10, int(random.gauss(55, 20)))
    if priority == "P2":
        return max(15, int(random.gauss(180, 80)))
    if priority == "P3":
        return max(10, int(random.gauss(600, 260)))
    return max(20, int(random.gauss(1500, 500)))  # P4

def main():
    out = Path("data/sample_incidents.csv")
    out.parent.mkdir(parents=True, exist_ok=True)

    now = datetime.now()
    start = now - timedelta(days=30)

    rows = []
    for i in range(1, 241):  # 240 incidents
        opened = start + timedelta(minutes=random.randint(0, 30 * 24 * 60))
        issue, pri_choices = random.choice(ISSUES)
        priority = random.choice(pri_choices)

        role = random.choices(ROLES, weights=[10, 12, 20, 45, 13], k=1)[0]
        device = random.choices(DEVICES, weights=[70, 30], k=1)[0]
        site = random.choices(SITES, weights=[45, 20, 35], k=1)[0]
        net = random.choices(NETWORK, weights=[30, 25, 45], k=1)[0]
        vendor = random.choice(VENDORS)

        # unresolved tail (about 6%)
        unresolved = random.random() < 0.06

        res_minutes = gen_resolution_minutes(priority)
        resolved_at = "" if unresolved else (opened + timedelta(minutes=res_minutes)).strftime("%Y-%m-%d %H:%M")

        rows.append({
            "incident_id": f"INC{i:04d}",
            "opened_at": opened.strftime("%Y-%m-%d %H:%M"),
            "resolved_at": resolved_at,
            "user_role": role,
            "device_type": device,
            "site": site,
            "network_path": net,
            "vendor": vendor,
            "issue_category": issue,
            "priority": priority,
            "resolution_minutes": "" if unresolved else res_minutes,
            "resolved": "No" if unresolved else "Yes",
        })

    with out.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)

    print(f"Wrote {len(rows)} rows to {out}")

if __name__ == "__main__":
    main()
