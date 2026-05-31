"""
send_final_reminders.py

Final reminder variant: finds "Awaiting Response" items overdue by
MIN_DAYS_OVERDUE days (default 2), and sends a "final opportunity" notice
per contractor via the locally installed Outlook app.
"""

import pandas as pd
import json
import os
import win32com.client
from datetime import date, datetime
from difflib import SequenceMatcher

# ─── Configuration ────────────────────────────────────────────────────────────
SENDER_EMAIL      = "nachiketm@vconstruct.in"

CC_EMAILS         = []

PROJECT_NAME      = "Freebird PH1"

RFQ_LOG_FILE      = "FreebirdPH1_RFQ_LOG.xlsx"
RFQ_SHEET         = "Clearstory Export CN Log"
RFQ_HEADER_ROW    = 3                            # 0-indexed (row 4 in Excel)

CONTACT_FILE      = "Freebird  Contact Directory.xlsx"
NAME_MAPPING_FILE = "name_mapping.json"
LOG_FILE          = "final_reminder_log.txt"

MIN_DAYS_OVERDUE  = 2                            # Items overdue by 2+ days
DRY_RUN           = False

TEST_EMAIL        = "nachiketm@vconstruct.in"
SAVE_AS_DRAFT     = True
# ──────────────────────────────────────────────────────────────────────────────


def load_name_mapping():
    if os.path.exists(NAME_MAPPING_FILE):
        with open(NAME_MAPPING_FILE) as f:
            return json.load(f)
    return {}


def fuzzy_match(name, candidates, threshold=0.55):
    best, best_score = None, 0
    name_lower = name.lower()
    for candidate in candidates:
        score = SequenceMatcher(None, name_lower, str(candidate).lower()).ratio()
        if score > best_score:
            best_score = score
            best = candidate
    return best if best_score >= threshold else None


def fmt_date(d):
    if pd.isna(d) or d is None:
        return ""
    if hasattr(d, "strftime"):
        return d.strftime("%-m/%-d/%Y") if os.name != "nt" else d.strftime("%#m/%#d/%Y")
    return str(d)


def build_email(sub_abbreviation, items, report_date):
    date_str = report_date.strftime("%m/%d/%Y")
    subject = f"{PROJECT_NAME} - Final Reminder Overdue RFQ - {date_str}"

    items_sorted = sorted(items, key=lambda x: x["days_overdue"], reverse=True)

    rows_html = ""
    for item in items_sorted:
        rows_html += f"""
      <tr style="line-height:1.4;">
        <td style="padding:5px 10px;border:1px solid #c0c0c0;font-size:11px;white-space:nowrap;">{item["reference"]}</td>
        <td style="padding:5px 10px;border:1px solid #c0c0c0;font-size:11px;">{item["cn_title"]}</td>
        <td style="padding:5px 10px;border:1px solid #c0c0c0;font-size:11px;background:#ebebeb;white-space:nowrap;">{sub_abbreviation}</td>
        <td style="padding:5px 10px;border:1px solid #c0c0c0;font-size:11px;background:#fff2cc;text-align:center;white-space:nowrap;">{item["date_sent"]}</td>
        <td style="padding:5px 10px;border:1px solid #c0c0c0;font-size:11px;background:#fff2cc;text-align:center;white-space:nowrap;">{item["due_date"]}</td>
        <td style="padding:5px 10px;border:1px solid #c0c0c0;font-size:11px;background:#ff0000;color:#ffffff;font-weight:bold;text-align:center;white-space:nowrap;">{item["days_overdue"]}</td>
      </tr>"""

    html = f"""<html>
<body style="font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#222;margin:0;padding:0;">

<p style="margin:6px 0;">All,</p>

<p style="margin:6px 0;">Below PCIs has been open for over one week, and we have not received a detailed change impact submission from your company. Unless your detailed cost impact submission is received within the next <strong>24 Hours</strong>, DPR Sundt JV will assume there are <strong>No cost or schedule impacts</strong> to you and will close this change impact with the Owner at <strong>$0</strong>.</p>

<p style="margin:6px 0;">This is your <strong>final reminder</strong> and final opportunity to pursue a cost and/or time adjustment to your subcontract related to this change. Your pricing must include a detailed breakdown of labor, material, and equipment. If you have a separate contract, you must also include a corresponding <strong>E-25 attachment</strong>.</p>

<table style="border-collapse:collapse;margin:10px 0;font-family:Calibri,Arial,sans-serif;">
  <thead>
    <tr style="background:#1f3864;color:white;line-height:1.4;">
      <th style="padding:5px 10px;border:1px solid #1f3864;font-size:11px;text-align:left;white-space:nowrap;">PCI No.</th>
      <th style="padding:5px 10px;border:1px solid #1f3864;font-size:11px;text-align:left;white-space:nowrap;">PCI Description</th>
      <th style="padding:5px 10px;border:1px solid #1f3864;font-size:11px;text-align:left;white-space:nowrap;">Sub Abbreviation</th>
      <th style="padding:5px 10px;border:1px solid #1f3864;font-size:11px;text-align:center;white-space:nowrap;">RFQ Sent</th>
      <th style="padding:5px 10px;border:1px solid #1f3864;font-size:11px;text-align:center;white-space:nowrap;">RFQ Due</th>
      <th style="padding:5px 10px;border:1px solid #1f3864;font-size:11px;text-align:center;white-space:nowrap;">Days Delay</th>
    </tr>
  </thead>
  <tbody>{rows_html}
  </tbody>
</table>

<p style="margin:6px 0;">Please note you are providing pricing through Clearstory only, separate e-mails as pricings won't be considered.</p>

<p style="margin:6px 0;">Thank you,</p>

<p style="margin:6px 0;">
  <strong>Nachiket Mahajan</strong> | vConstruct | <a href="http://www.vconstruct.com" style="color:#222;">www.vconstruct.com</a><br>
  India: +91 9028085157<br>
  <span style="color:#555;font-size:10pt;">4th &amp; 5th floor, Tower B1, SEZ, Magarpatta City, Pune - 411013, India</span>
</p>

</body>
</html>"""

    return subject, html


def log(message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{timestamp}] {message}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def main():
    today = date.today()
    log(f"=== Final reminder run started - today is {today} | DRY_RUN={DRY_RUN} ===")

    df = pd.read_excel(RFQ_LOG_FILE, sheet_name=RFQ_SHEET, header=RFQ_HEADER_ROW)
    df["Due Date"]  = pd.to_datetime(df["Due Date"],  errors="coerce").dt.date
    df["Date Sent"] = pd.to_datetime(df["Date Sent"], errors="coerce").dt.date
    df["days_overdue"] = df["Due Date"].apply(
        lambda d: (today - d).days if pd.notna(d) else None
    )

    overdue = df[
        (df["CN Response Status"] == "Awaiting Response") &
        (df["days_overdue"] == MIN_DAYS_OVERDUE)
    ].copy()

    log(f"Found {len(overdue)} overdue items across {overdue['Contractor Name'].nunique()} contractors")

    if overdue.empty:
        log("Nothing to send. Exiting.")
        return

    contacts_df = pd.read_excel(CONTACT_FILE)
    contacts_df["Trade Partner"] = contacts_df["Trade Partner"].ffill()
    accepted      = contacts_df[contacts_df["Status"] == "Accepted"]
    trade_partners = contacts_df["Trade Partner"].dropna().unique().tolist()
    name_mapping  = load_name_mapping()

    sent_count = 0
    skip_count = 0

    for contractor, group in overdue.groupby("Contractor Name"):

        mapped_name = name_mapping.get(contractor)
        if mapped_name is None and contractor in trade_partners:
            mapped_name = contractor
        if not mapped_name:
            mapped_name = fuzzy_match(contractor, trade_partners)

        if not mapped_name:
            log(f"  SKIP -- No contact match for '{contractor}' -> add mapping to {NAME_MAPPING_FILE}")
            skip_count += 1
            continue

        recipients = accepted[accepted["Trade Partner"] == mapped_name]
        if recipients.empty:
            recipients = contacts_df[contacts_df["Trade Partner"] == mapped_name]

        if recipients.empty:
            log(f"  SKIP -- Matched '{mapped_name}' but no email addresses found")
            skip_count += 1
            continue

        items = [
            {
                "reference":    row["Reference #"],
                "cn_title":     row["CN Title"],
                "date_sent":    fmt_date(row["Date Sent"]),
                "due_date":     fmt_date(row["Due Date"]),
                "days_overdue": int(row["days_overdue"]),
            }
            for _, row in group.iterrows()
        ]

        to_emails  = [TEST_EMAIL] if TEST_EMAIL else recipients["Email"].dropna().tolist()
        subject, html_body = build_email(mapped_name, items, today)

        log(f"  Contractor : {contractor}")
        log(f"  Matched to : {mapped_name}")
        log(f"  Items      : {len(items)}  |  Days overdue: {[i['days_overdue'] for i in items]}")
        log(f"  TO         : {to_emails}{' [TEST MODE]' if TEST_EMAIL else ''}")
        if CC_EMAILS:
            log(f"  CC         : {CC_EMAILS}")
        log(f"  Subject    : {subject}")

        if DRY_RUN:
            log(f"  [DRY RUN] Email NOT sent")
            log("")
            continue

        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)
            mail.SentOnBehalfOfName = SENDER_EMAIL
            mail.To      = "; ".join(to_emails)
            mail.CC      = "; ".join(CC_EMAILS) if CC_EMAILS else ""
            mail.Subject = subject
            mail.HTMLBody = html_body

            if SAVE_AS_DRAFT:
                mail.Save()
                log(f"  SAVED TO DRAFTS")
            else:
                mail.Send()
                log(f"  SENT OK")
            sent_count += 1
        except Exception as exc:
            log(f"  ERROR -- {exc}")

        log("")

    log(f"=== Run complete - sent: {sent_count}, skipped: {skip_count} ===\n")


if __name__ == "__main__":
    main()
