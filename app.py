# ============================================================
# 1. IMPORTS & GLOBAL CONFIGURATION
# ============================================================
import os
import re
import io
import json
import math
import tempfile
import warnings
import smtplib
from datetime import datetime
from collections import Counter
from zoneinfo import ZoneInfo
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

import pandas as pd
import gradio as gr
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side

warnings.filterwarnings("ignore")

# Environment Variables for Direct Gmail SMTP Dispatch
os.environ.setdefault("SENDER_EMAIL", "nasirnawaz918@gmail.com")
os.environ.setdefault("SENDER_PASSWORD", "ymqj motw fpxq bqqv")

# Timezone Configuration (Pakistan Standard Time - PKT)
PKT = ZoneInfo("Asia/Karachi")

# Runtime Memory Incident Tracking Database
incident_records = {}


# ============================================================
# 2. SHARED HELPERS & TIME FUNCTIONS
# ============================================================
def clean(value):
    """Clean and strip input string safely."""
    return (value or "").strip()


def get_current_time():
    """Returns current datetime in Asia/Karachi timezone."""
    return datetime.now(PKT)


def format_dt(value):
    """Formats datetime object to standard corporate representation."""
    if not value:
        return "NA"
    return value.strftime("%d-%m-%Y %I:%M %p")


def detect_service_type(text):
    """Detects standard telecom service type from label or complaint string."""
    value = clean(text).upper()

    checks = [
        ("BGPDIA", "BGP DIA"),
        ("DIA BGP", "BGP DIA"),
        ("MPLS", "MPLS"),
        ("DPLC", "DPLC"),
        ("TURBONET", "Turbonet"),
        ("TURBO", "Turbonet"),
        ("SIP PRI", "SIP PRI"),
        ("SIP_PRI", "SIP PRI"),
        ("VPBX", "VPBX"),
        ("IPLC", "IPLC"),
        ("DARKCORE", "Darkcore Fiber"),
        ("DARK CORE", "Darkcore Fiber"),
        ("M2M", "M2M"),
        ("PRI", "PRI"),
        ("SIP", "SIP"),
        ("DIA", "DIA"),
        ("P2P", "P2P"),
    ]

    for token, service in checks:
        if token in value:
            return service

    # Detect VPBX / Master phone numbers
    digits = re.sub(r"\D", "", clean(text))
    if len(digits) in (10, 11) and (digits.startswith("3") or digits.startswith("03")):
        return "VPBX"

    return "Corporate"


def extract_service_id(label):
    """Extracts short unique service link identifier for concise subjects."""
    value = clean(label)
    if not value:
        return "Service"

    patterns = [
        r"\bDPLC\d+SL\d+\b",
        r"\bDIA\d+SL\d+\b",
        r"\bTurbo\d+SL\d+\b",
        r"\bLNK[A-Z0-9]+\b",
        r"\bMPLS[A-Z0-9-]*\b",
        r"\bPRI[A-Z0-9-]*\b",
        r"\bSIP[A-Z0-9-]*\b",
    ]

    for pattern in patterns:
        match = re.search(pattern, value, flags=re.IGNORECASE)
        if match:
            return match.group(0)

    if "_" in value:
        candidate = value.split("_")[-1].strip()
        if candidate:
            return candidate[:70]

    return value[:70]


def priority_prefix(priority):
    """Returns email subject prefix tag according to escalation priority."""
    mapping = {
        "Normal": "",
        "Follow-up": "FOLLOW-UP | ",
        "Urgent": "URGENT | ",
        "Critical": "CRITICAL | ",
    }
    return mapping.get(priority, "")


def make_subject(label, topic, priority="Normal", ticket=""):
    """Standardized email subject generator."""
    service_id = extract_service_id(label)
    ticket = clean(ticket)
    ticket_part = f" | {ticket}" if ticket else ""
    return f"{priority_prefix(priority)}{topic} | {service_id}{ticket_part}"


def salutation(audience):
    """Returns appropriate professional salutation based on target audience."""
    if audience == "Customer":
        return "Dear Customer,"
    return "Dear Team,"


def calculate_duration(start_time, end_time):
    """Calculates formatted total outage duration between two datetime objects."""
    if not start_time or not end_time:
        return "NA"

    total_minutes = max(0, int((end_time - start_time).total_seconds() // 60))
    days, remainder = divmod(total_minutes, 1440)
    hours, minutes = divmod(remainder, 60)

    if days:
        return f"{days} Day(s) {hours} Hour(s) {minutes} Minute(s)"
    if hours:
        return f"{hours} Hour(s) {minutes} Minute(s)"
    return f"{minutes} Minute(s)"


def elapsed_text(start_time, end_time=None):
    """Calculates concise running age string for active incidents."""
    if not start_time:
        return "Not started"

    end_time = end_time or get_current_time()
    minutes = max(0, int((end_time - start_time).total_seconds() // 60))

    if minutes >= 1440:
        days, rem = divmod(minutes, 1440)
        hours, mins = divmod(rem, 60)
        return f"{days}d {hours}h {mins}m"
    if minutes >= 60:
        hours, mins = divmod(minutes, 60)
        return f"{hours}h {mins}m"
    return f"{minutes}m"


def incident_age_text(reported_time, added_time=None, status="OPEN"):
    """Generates visual status indicators (🟢 🟡 🟠 🔴) based on active response age."""
    base_time = reported_time or added_time
    if not base_time:
        return "—"

    if status == "CLOSED":
        return "Closed"

    minutes = max(0, int((get_current_time() - base_time).total_seconds() // 60))

    if minutes < 5:
        marker = "🟢"
    elif minutes < 15:
        marker = "🟡"
    elif minutes < 30:
        marker = "🟠"
    else:
        marker = "🔴"

    prefix = "Queued " if not reported_time else ""
    return f"{marker} {prefix}{elapsed_text(base_time)}"


TIME_MODES = [
    "Automatic - Current Pakistan Time",
    "Manual Date/Time",
]


def parse_datetime_input(value):
    """Parses various date-time format strings into a PKT-aware datetime object."""
    value = clean(value)
    if not value:
        raise ValueError("Manual date/time is empty.")

    formats = [
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d %I:%M:%S %p",
        "%Y-%m-%d %I:%M %p",
        "%Y-%m-%dT%H:%M:%S",
        "%Y-%m-%dT%H:%M",
        "%d-%m-%Y %H:%M:%S",
        "%d-%m-%Y %H:%M",
        "%d-%m-%Y %I:%M:%S %p",
        "%d-%m-%Y %I:%M %p",
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%d/%m/%Y %I:%M:%S %p",
        "%d/%m/%Y %I:%M %p",
    ]

    for fmt in formats:
        try:
            parsed = datetime.strptime(value, fmt)
            return parsed.replace(tzinfo=PKT)
        except ValueError:
            continue

    raise ValueError(
        "Invalid date/time. Use YYYY-MM-DD HH:MM, DD-MM-YYYY HH:MM, or include AM/PM."
    )


def resolve_event_time(mode, manual_value):
    """Resolves event timestamp based on auto vs manual selection."""
    if mode == "Manual Date/Time":
        return parse_datetime_input(manual_value)
    return get_current_time()


def ensure_incident(label, issue_type="Service issue", ticket="", stage="Initial"):
    """Ensures incident record exists in runtime memory and updates its state."""
    label = clean(label)
    if not label:
        return

    now = get_current_time()

    if label not in incident_records:
        incident_records[label] = {
            "added_time": now,
            "reported_time": now,
            "restoration_time": None,
            "service_type": detect_service_type(label),
            "issue_type": clean(issue_type) or "Service issue",
            "ticket": clean(ticket),
            "status": "OPEN",
            "last_stage": stage,
        }
        return

    record = incident_records[label]

    if not record.get("reported_time") and record.get("status") == "QUEUED":
        record["reported_time"] = now
        record["status"] = "OPEN"

    record["last_stage"] = stage
    if issue_type:
        record["issue_type"] = clean(issue_type)
    if ticket:
        record["ticket"] = clean(ticket)


def update_incident_stage(label, stage):
    """Updates the last stage marker of an existing active incident."""
    label = clean(label)
    if label in incident_records:
        incident_records[label]["last_stage"] = stage


def format_ettr(ettr_choice, custom_ettr):
    """Formats standardized Expected Time to Restore (ETTR) statements."""
    if ettr_choice == "Awaited":
        return "ETTR will be shared once received from the concerned team."
    if ettr_choice == "Not Available":
        return "Currently, no confirmed ETTR is available. Further updates will be shared accordingly."
    if ettr_choice == "Not Applicable":
        return ""
    if ettr_choice == "Custom":
        value = clean(custom_ettr)
        return (
            f"The tentative ETTR shared by the concerned team is {value}."
            if value
            else "ETTR will be shared once received from the concerned team."
        )
    return f"The tentative ETTR shared by the concerned team is {ettr_choice}."


def multiline_links(value):
    """Splits multiline input text into clean lists of individual link labels."""
    entries = []
    for raw in clean(value).splitlines():
        item = raw.strip(" \t-•")
        if item:
            entries.append(item)
    return entries


# ============================================================
# 3. COMPLAINT MANAGER ENGINE
# ============================================================
COMPLAINT_DEFAULT_ISSUES = [
    "Link is down.",
    "Service degradation.",
    "Intermittent connectivity.",
    "High latency observed.",
    "Packet loss observed.",
    "Slow browsing issue.",
    "Bandwidth issue.",
    "Website / application issue.",
    "Voice / VPBX issue.",
    "Other / To be classified",
]


def get_complaint_labels(include_closed=True):
    """Retrieves list of runtime complaint labels sorted by addition time."""
    labels = []
    for label, record in incident_records.items():
        if not include_closed and record.get("status") == "CLOSED":
            continue
        labels.append(label)

    labels.sort(
        key=lambda x: incident_records[x].get("added_time") or get_current_time(),
        reverse=True,
    )
    return labels


def register_complaints(raw_labels, default_issue, ticket):
    """Registers multiple complaints into runtime memory queue simultaneously."""
    entries = multiline_links(raw_labels)
    if not entries:
        return (
            gr.Dropdown(choices=get_complaint_labels(), value=None),
            "⚠️ Paste one or more complaints/service labels, one per line.",
            refresh_dashboard("Queued + Open"),
        )

    now = get_current_time()
    added = []
    skipped = []
    restarted = []

    for label in entries:
        existing = incident_records.get(label)

        if existing and existing.get("status") in ("QUEUED", "OPEN"):
            skipped.append(label)
            continue

        if existing and existing.get("status") == "CLOSED":
            restarted.append(label)

        incident_records[label] = {
            "added_time": now,
            "reported_time": None,
            "restoration_time": None,
            "service_type": detect_service_type(label),
            "issue_type": clean(default_issue) or "Other / To be classified",
            "ticket": clean(ticket),
            "status": "QUEUED",
            "last_stage": "Complaint added - awaiting opening",
        }
        added.append(label)

    choices = get_complaint_labels()
    selected = added[0] if added else (choices[0] if choices else None)

    parts = []
    if added:
        parts.append(f"✅ Added {len(added)} complaint(s) to runtime queue.")
    if restarted:
        parts.append(
            f"♻️ {len(restarted)} previously closed label(s) started as new queued complaint(s)."
        )
    if skipped:
        parts.append(
            f"ℹ️ Skipped {len(skipped)} duplicate complaint(s) already QUEUED/OPEN."
        )

    return (
        gr.Dropdown(choices=choices, value=selected),
        "\n".join(parts),
        refresh_dashboard("Queued + Open"),
    )


def load_selected_complaint(selected_label):
    """Loads complaint details into active editing state."""
    label = clean(selected_label)
    if not label or label not in incident_records:
        return "", "", "⚠️ Please select a runtime complaint."

    record = incident_records[label]
    status = record.get("status", "")
    opened = (
        format_dt(record["reported_time"])
        if record.get("reported_time")
        else "Not opened yet"
    )
    closed = (
        format_dt(record["restoration_time"])
        if record.get("restoration_time")
        else "Not closed"
    )

    summary = f"""✅ Complaint loaded.
Status: {status}
Service Type: {record.get("service_type", "")}
Issue: {record.get("issue_type", "")}
Opening Time: {opened}
Closing Time: {closed}
Ticket: {record.get("ticket", "")}"""

    return label, record.get("ticket", ""), summary


def remove_selected_complaint(selected_label):
    """Deletes complaint record from active memory."""
    label = clean(selected_label)
    if not label or label not in incident_records:
        return (
            gr.Dropdown(choices=get_complaint_labels(), value=None),
            "⚠️ Please select a valid runtime complaint.",
            refresh_dashboard("Queued + Open"),
        )

    del incident_records[label]
    choices = get_complaint_labels()
    selected = choices[0] if choices else None

    return (
        gr.Dropdown(choices=choices, value=selected),
        "✅ Selected complaint removed from runtime memory.",
        refresh_dashboard("Queued + Open"),
    )


def refresh_complaint_selector():
    """Refreshes dropdown list choices of complaints."""
    choices = get_complaint_labels()
    return gr.Dropdown(
        choices=choices,
        value=choices[0] if choices else None,
    )


def get_complaint_labels_by_status(statuses=None):
    """Filters complaint labels by status (QUEUED, OPEN, CLOSED)."""
    labels = []
    allowed = set(statuses) if statuses else None

    for complaint_label, record in incident_records.items():
        status = record.get("status", "QUEUED")
        if allowed is not None and status not in allowed:
            continue
        labels.append(complaint_label)

    labels.sort(
        key=lambda x: incident_records[x].get("added_time") or get_current_time(),
        reverse=True,
    )
    return labels


def refresh_opening_complaint_selector():
    """Refreshes list choices for Opening Email tab."""
    choices = get_complaint_labels_by_status({"QUEUED", "OPEN"})
    return gr.Dropdown(
        choices=choices,
        value=choices[0] if choices else None,
    )


def refresh_closure_complaint_selector():
    """Refreshes list choices for Closure Email tab."""
    choices = get_complaint_labels_by_status({"OPEN"})
    return gr.Dropdown(
        choices=choices,
        value=choices[0] if choices else None,
    )


def resolve_complaint_label(selected_label, fallback_label):
    """Resolves label between selected complaint dropdown or textbox fallback."""
    return clean(selected_label) or clean(fallback_label)


def resolve_picker_time(value):
    """Converts picker datetime input into PKT timezone aware object."""
    if value in (None, ""):
        return get_current_time()

    if isinstance(value, datetime):
        if value.tzinfo is None:
            return value.replace(tzinfo=PKT)
        return value.astimezone(PKT)

    if isinstance(value, (int, float)):
        return datetime.fromtimestamp(value, tz=PKT)

    return parse_datetime_input(value)


# ============================================================
# 4. CLICK-ONLY DATE / CLOCK PICKER HELPERS
# ============================================================
AUTO_TIME_MODE = "Automatic - Current Pakistan Time"
MANUAL_TIME_MODE = "Manual - Mouse Clock Picker"

HOUR_CHOICES = [f"{i:02d}" for i in range(1, 13)]
MINUTE_CHOICES = [f"{i:02d}" for i in range(60)]
AMPM_CHOICES = ["AM", "PM"]


def _picker_date_to_date(value):
    """Helper to convert string/datetime date pickers into python date objects."""
    if value in (None, ""):
        return None

    if isinstance(value, datetime):
        return value.astimezone(PKT).date() if value.tzinfo else value.date()

    if isinstance(value, (int, float)):
        return datetime.fromtimestamp(value, tz=PKT).date()

    raw = clean(value)
    if not raw:
        return None

    for fmt in (
        "%Y-%m-%d",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%dT%H:%M:%S",
        "%Y-%m-%dT%H:%M",
        "%d-%m-%Y",
        "%d/%m/%Y",
    ):
        try:
            return datetime.strptime(raw, fmt).date()
        except ValueError:
            continue

    raise ValueError("Please select a valid date from the calendar.")


def resolve_click_clock_time(mode, date_value, hour_value, minute_value, ampm_value):
    """Calculates precise datetime object from mouse clock controls."""
    if mode != MANUAL_TIME_MODE:
        return get_current_time()

    selected_date = _picker_date_to_date(date_value)
    if selected_date is None:
        raise ValueError("Please select the date from the calendar.")

    hour_text = clean(hour_value)
    minute_text = clean(minute_value)
    ampm = clean(ampm_value).upper()

    if not hour_text or not minute_text or ampm not in AMPM_CHOICES:
        raise ValueError("Please select hour, minute and AM/PM using the clock controls.")

    hour12 = int(hour_text)
    minute = int(minute_text)
    if not (1 <= hour12 <= 12):
        raise ValueError("Hour must be between 1 and 12.")
    if not (0 <= minute <= 59):
        raise ValueError("Minute must be between 00 and 59.")

    hour24 = hour12 % 12
    if ampm == "PM":
        hour24 += 12

    return datetime(
        selected_date.year,
        selected_date.month,
        selected_date.day,
        hour24,
        minute,
        0,
        tzinfo=PKT,
    )


def clock_card_html(hour_value=None, minute_value=None, ampm_value=None, date_value=None):
    """Renders dynamic analog clock SVG preview card with digital time display."""
    now = get_current_time()

    try:
        hour12 = int(clean(hour_value)) if clean(hour_value) else ((now.hour - 1) % 12) + 1
    except ValueError:
        hour12 = ((now.hour - 1) % 12) + 1

    try:
        minute = int(clean(minute_value)) if clean(minute_value) else now.minute
    except ValueError:
        minute = now.minute

    ampm = clean(ampm_value).upper() if clean(ampm_value) else now.strftime("%p")
    if ampm not in AMPM_CHOICES:
        ampm = now.strftime("%p")

    minute_angle = minute * 6
    hour_angle = (hour12 % 12) * 30 + minute * 0.5

    def endpoint(cx, cy, length, angle_deg):
        angle = math.radians(angle_deg - 90)
        return cx + length * math.cos(angle), cy + length * math.sin(angle)

    cx, cy = 92, 72
    hx, hy = endpoint(cx, cy, 30, hour_angle)
    mx, my = endpoint(cx, cy, 43, minute_angle)

    nums = []
    for n in range(1, 13):
        x, y = endpoint(cx, cy, 55, n * 30)
        nums.append(
            f'<text x="{x:.1f}" y="{y + 4:.1f}" text-anchor="middle" '
            'font-size="11" font-weight="700" fill="#fff">'
            f'{n}</text>'
        )

    date_text = ""
    try:
        d = _picker_date_to_date(date_value)
        if d:
            date_text = d.strftime("%d-%m-%Y")
    except Exception:
        date_text = ""

    digital = f"{hour12:02d}:{minute:02d} {ampm}"
    subtitle = f"{date_text} · Pakistan (PKT)" if date_text else "Pakistan (PKT)"

    return f'''<div style="display:flex;align-items:center;gap:24px;background:#191919;color:white;border-radius:22px;padding:16px 24px;max-width:520px;min-height:150px;box-sizing:border-box;">
      <div style="min-width:190px;">
        <div style="font-size:30px;font-weight:750;line-height:1.1;">{digital}</div>
        <div style="font-size:14px;color:#c9c9c9;margin-top:10px;">{subtitle}</div>
        <div style="font-size:12px;color:#9fdaff;margin-top:8px;">Mouse clock preview</div>
      </div>
      <svg width="184" height="144" viewBox="0 0 184 144" role="img" aria-label="Selected analogue time">
        <circle cx="92" cy="72" r="67" fill="#202020" stroke="#3a3a3a" stroke-width="2"/>
        {''.join(nums)}
        <line x1="92" y1="72" x2="{hx:.1f}" y2="{hy:.1f}" stroke="#e8e8e8" stroke-width="4" stroke-linecap="round"/>
        <line x1="92" y1="72" x2="{mx:.1f}" y2="{my:.1f}" stroke="#e8e8e8" stroke-width="3" stroke-linecap="round"/>
        <circle cx="92" cy="72" r="4" fill="#ff6b35"/>
      </svg>
    </div>'''


def current_clock_picker_values():
    """Resets clock picker state variables to current PKT time."""
    now = get_current_time().replace(second=0, microsecond=0)
    return (
        MANUAL_TIME_MODE,
        now.strftime("%Y-%m-%d"),
        now.strftime("%I"),
        now.strftime("%M"),
        now.strftime("%p"),
        clock_card_html(
            now.strftime("%I"),
            now.strftime("%M"),
            now.strftime("%p"),
            now.strftime("%Y-%m-%d"),
        ),
    )


def update_clock_preview(date_value, hour_value, minute_value, ampm_value):
    """Wrapper to update clock preview on dropdown value changes."""
    return clock_card_html(hour_value, minute_value, ampm_value, date_value)


# ============================================================
# 5. SMART PASTE / COMPLAINT PARSER
# ============================================================
QUICK_SCENARIOS = [
    "Acknowledgement",
    "Link Down",
    "Link Up / Monitoring",
    "Fiber Break",
    "Node Offline",
    "Power Issue",
    "Customer End Issue",
    "Last Mile Issue",
    "Port Down",
    "Vendor Issue",
    "Configuration Issue",
    "High Utilization / Link Choking",
    "Packet Loss / Latency",
    "Slow Browsing",
    "Website / Application Issue",
    "BRAS / User Offline",
    "Optical Power Issue",
    "Routing / BGP Issue",
    "Rerouting",
    "Submarine Cable Cut",
    "Hardware Fault",
    "No Issue Observed",
    "Confirmation Sent",
    "Missed Response Apology",
]


def detect_scenario(text):
    """Detects primary outage scenario from raw text keywords."""
    value = clean(text).lower()

    rules = [
        (["submarine", "smw", "cable cut"], "Submarine Cable Cut"),
        (["fiber break", "fibre break", "fiber cut"], "Fiber Break"),
        (["node offline", "node down"], "Node Offline"),
        (["power issue", "power down"], "Power Issue"),
        (["packet loss", "latency"], "Packet Loss / Latency"),
        (["slow speed", "slow browsing", "slow internet"], "Slow Browsing"),
        (["website", "url", "banking app", "tiktok"], "Website / Application Issue"),
        (["user offline", "bras"], "BRAS / User Offline"),
        (["optical power", "rx power", "tx power"], "Optical Power Issue"),
        (["bgp", "route advertisement", "routing issue", "advertise to ptcl"], "Routing / BGP Issue"),
        (["utilization", "choking", "congestion"], "High Utilization / Link Choking"),
        (["intermittent", "fluctuation"], "Last Mile Issue"),
        (["incoming call", "outgoing call", "call connectivity"], "Vendor Issue"),
        (["link down", "showing down", "service outage", "down"], "Link Down"),
        (["link up", "working fine", "restored"], "Link Up / Monitoring"),
    ]

    for keywords, scenario in rules:
        if any(keyword in value for keyword in keywords):
            return scenario
    return "Acknowledgement"


def parse_complaint(raw_text):
    """Parses raw customer email/text dump to extract label, VLAN, ticket, and scenario."""
    text = clean(raw_text)
    if not text:
        return "", "", "Acknowledgement", "", "", "Paste a complaint/email first."

    label = ""

    match = re.search(r"(ESSClient[^\r\n]+)", text, flags=re.IGNORECASE)
    if match:
        label = match.group(1).strip()
    else:
        link_match = re.search(
            r"(?:Link\s*Name|Service\s*Details|Label)\s*[:=-]\s*([^\r\n]+)",
            text,
            flags=re.IGNORECASE,
        )
        if link_match:
            label = link_match.group(1).strip()
        else:
            site_match = re.search(
                r"\bSITE[-:\s]+([A-Z0-9 _-]+)", text, flags=re.IGNORECASE
            )
            if site_match:
                label = f"SITE-{site_match.group(1).strip()}"

    vlan_match = re.search(r"\bVLAN\s*[-:]?\s*(\d{1,4})\b", text, flags=re.IGNORECASE)
    vlan = vlan_match.group(1) if vlan_match else ""

    ticket_match = re.search(
        r"\b(?:INC|TT|TKT|TICKET)[-:#\s]*([A-Z0-9-]{4,})\b", text, flags=re.IGNORECASE
    )
    ticket = ticket_match.group(0).strip() if ticket_match else ""

    scenario = detect_scenario(text)
    service = detect_service_type(label or text)

    summary = (
        f"Detected Service: {service}\n"
        f"Detected Scenario: {scenario}\n"
        f"Detected VLAN: {vlan or 'Not found'}\n"
        f"Detected Ticket: {ticket or 'Not found'}"
    )

    return label, vlan, scenario, ticket, ticket, summary


# ============================================================
# 6. OPENING EMAIL ENGINE
# ============================================================
OPENING_ISSUES = [
    "Link is down.",
    "Service degradation.",
    "Intermittent connectivity.",
    "High latency observed.",
    "Packet loss observed.",
    "Slow browsing issue.",
    "Bandwidth issue.",
    "Website / application issue.",
    "BRAS / user connectivity issue.",
    "Call connectivity issue.",
    "Incoming call issue.",
    "Outgoing call issue.",
    "Other / Manual",
]


def generate_opening_email(
    selected_complaint,
    label,
    issue_type,
    custom_issue,
    ticket,
    priority,
    opening_time_mode,
    opening_date,
    opening_hour,
    opening_minute,
    opening_ampm,
):
    """Generates standardized Opening Acknowledgement email template."""
    label = resolve_complaint_label(selected_complaint, label)
    if not label:
        return "", "", "", "⚠️ Please select a complaint or enter the service label."

    issue = (
        clean(custom_issue) or "Service issue reported."
        if issue_type == "Other / Manual"
        else issue_type
    )

    try:
        requested_time = resolve_click_clock_time(
            opening_time_mode,
            opening_date,
            opening_hour,
            opening_minute,
            opening_ampm,
        )
    except (ValueError, TypeError) as exc:
        return label, "", "", f"⚠️ Invalid opening date/time: {exc}"

    service_type = detect_service_type(label)
    now = get_current_time()
    record = incident_records.get(label)
    reused_existing_time = False

    if record and record.get("status") == "OPEN" and record.get("reported_time"):
        reported_time = record["reported_time"]
        reused_existing_time = True
    else:
        reported_time = requested_time
        incident_records[label] = {
            "added_time": record.get("added_time") if record else now,
            "reported_time": reported_time,
            "restoration_time": None,
            "service_type": (
                record.get("service_type")
                if record and record.get("service_type")
                else service_type
            ),
            "issue_type": issue,
            "ticket": clean(ticket) or (record.get("ticket", "") if record else ""),
            "status": "OPEN",
            "last_stage": "Initial response",
        }

    record = incident_records[label]
    record["issue_type"] = issue
    record["ticket"] = clean(ticket) or record.get("ticket", "")
    record["last_stage"] = "Initial response"
    record["status"] = "OPEN"
    record["restoration_time"] = None

    subject = make_subject(label, issue.rstrip("."), priority, record.get("ticket", ""))

    email = f"""Dear Customer,

We acknowledge receipt of your complaint regarding service degradation/impact on your {record.get("service_type", service_type)} service. Our Corporate NOC has initiated an investigation.

Details:
Service Details: {label}
Issue Type: {issue}
Incident Start Time: {format_dt(reported_time)}
Current Status: Under investigation
Reference Ticket: {record.get("ticket", "")}

Our technical teams are actively working to identify the root cause and restore service at the earliest possible time. Regular updates will be shared until the issue is fully resolved.

We appreciate your patience and cooperation."""

    if reused_existing_time:
        runtime = (
            f"ℹ️ Complaint was already OPEN. Existing opening time retained: "
            f"{format_dt(reported_time)}."
        )
    else:
        source_note = (
            "manual mouse clock"
            if opening_time_mode == MANUAL_TIME_MODE
            else "current Pakistan time"
        )
        runtime = (
            f"✅ Complaint OPENED at {format_dt(reported_time)} "
            f"for {extract_service_id(label)} ({source_note})."
        )

    return label, subject, email, runtime


# ============================================================
# 7. QUICK RESPONSE ENGINE
# ============================================================
QUICK_STAGES = [
    "Initial",
    "Team Engaged",
    "Troubleshooting",
    "Team Dispatched",
    "Fault Located",
    "Restoration In Progress",
    "Rerouting In Progress",
    "Monitoring",
    "Service Restored",
]

QUICK_TEAMS = [
    "Concerned team",
    "Transmission Optical",
    "Transmission Microwave",
    "Field Operations (FOPs)",
    "IP Core",
    "Uplink ISP",
    "Vendor / Last-mile team",
    "Customer",
]


def scenario_statement(scenario):
    """Returns standard scenario initial explanation statement."""
    statements = {
        "Acknowledgement":
            "We acknowledge receipt of the reported issue and are currently assessing the matter.",
        "Link Down":
            "Please be informed that the reported link is currently down.",
        "Link Up / Monitoring":
            "The reported link is currently up and is being monitored for stability.",
        "Fiber Break":
            "The reported service is currently impacted due to a fiber break in the transmission segment.",
        "Node Offline":
            "The reported service is currently impacted as the serving node is offline.",
        "Power Issue":
            "The reported service is currently impacted due to a power issue at the serving node.",
        "Customer End Issue":
            "Our initial checks indicate that the reported issue is related to the customer-end connectivity.",
        "Last Mile Issue":
            "Our initial analysis indicates a possible last-mile connectivity issue.",
        "Port Down":
            "The reported service is impacted due to a port-down condition.",
        "Vendor Issue":
            "The issue has been escalated to the concerned vendor/upstream team for further investigation.",
        "Configuration Issue":
            "The reported service is impacted due to a configuration-related issue.",
        "High Utilization / Link Choking":
            "The link is experiencing high utilization, which may result in congestion and service degradation.",
        "Packet Loss / Latency":
            "The reported packet-loss/latency issue is under detailed investigation.",
        "Slow Browsing":
            "The reported slow-browsing/speed issue is under investigation.",
        "Website / Application Issue":
            "The reported website/application accessibility issue is under investigation.",
        "BRAS / User Offline":
            "The affected user/session is currently not being observed online at our BRAS end.",
        "Optical Power Issue":
            "Abnormal optical power has been observed and requires last-mile/optical verification.",
        "Routing / BGP Issue":
            "The reported routing/BGP path issue is under investigation and route advertisement/path selection is being verified.",
        "Rerouting":
            "Rerouting feasibility is being evaluated to minimize the service impact.",
        "Submarine Cable Cut":
            "The current service degradation is associated with an upstream/submarine cable fault.",
        "Hardware Fault":
            "The reported service is impacted due to a hardware-related fault.",
        "No Issue Observed":
            "We have checked the reported service and no abnormality is currently observed at our end.",
        "Confirmation Sent":
            "A confirmation email has been sent to the customer. We are awaiting their acknowledgement.",
        "Missed Response Apology":
            "We sincerely regret the delay in our response. Due to an internal technical issue, the email was inadvertently overlooked.",
    }
    return statements.get(scenario, "The reported issue is currently under investigation.")


def stage_statement(stage, team):
    """Returns stage progress description sentence."""
    team_name = clean(team) or "concerned team"
    statements = {
        "Initial":
            f"The issue has been escalated and the {team_name} is being engaged.",
        "Team Engaged":
            f"The {team_name} is actively engaged and working on the reported issue.",
        "Troubleshooting":
            f"Detailed troubleshooting is currently in progress with the {team_name}.",
        "Team Dispatched":
            f"The {team_name} has been dispatched for onsite verification/restoration.",
        "Fault Located":
            "The fault has been localized and restoration activity is being carried out on priority.",
        "Restoration In Progress":
            "Restoration activity is currently in progress and is being followed up on priority.",
        "Rerouting In Progress":
            "Rerouting activity is in progress to restore the affected service through an alternate path.",
        "Monitoring":
            "The service is currently under observation to verify stability and performance.",
        "Service Restored":
            "The service has been restored and is currently operating normally.",
    }
    return statements.get(stage, "")


def generate_quick_response(
    label,
    scenario,
    stage,
    team,
    ettr_choice,
    custom_ettr,
    priority,
    audience,
    ticket,
    custom_note,
):
    """Generates rapid standard response email for ongoing incidents."""
    label = clean(label)
    if not label:
        return "", "⚠️ Please enter the service label.", ""

    if stage == "Initial":
        ensure_incident(label, scenario, ticket, "Quick initial")
    else:
        update_incident_stage(label, stage)

    subject = make_subject(label, scenario, priority, ticket)
    hello = salutation(audience)

    if scenario == "Missed Response Apology":
        body = f"""{hello}

We sincerely regret the delay in our response.

Due to an internal technical issue on our end, this email was inadvertently overlooked. We are currently assessing the reported issue and will share a detailed update shortly.

Thank you for your understanding and patience."""
    elif scenario == "Confirmation Sent":
        body = f"""{hello}

A confirmation email has been sent to the customer. Once acknowledgement is received, further updates will be shared accordingly."""
    else:
        parts = [
            hello,
            "",
            scenario_statement(scenario),
            "",
            stage_statement(stage, team),
        ]

        if clean(custom_note):
            parts.extend(["", clean(custom_note)])

        ettr_line = format_ettr(ettr_choice, custom_ettr)
        if ettr_line:
            parts.extend(["", ettr_line])

        parts.extend([
            "",
            "Further updates will be shared accordingly.",
            "",
            "Your patience and cooperation are highly appreciated."
        ])
        body = "\n".join(parts)

    policy = ""
    if detect_service_type(label) == "Turbonet":
        policy = (
            "⚠️ Internal Turbo policy: Do not share the Turbonet utilization "
            "graph directly with the customer. Route graph-related requests "
            "through KAM/GCSSQ as per departmental instruction."
        )

    return subject, body, policy


# ============================================================
# 8. TRANSMISSION & MASS OUTAGE ENGINE
# ============================================================
TRANSMISSION_FAULTS = [
    "Node offline",
    "Power issue at Node",
    "Single Fiber break",
    "Dual Fiber break",
    "Triple Fiber break",
    "Spur Fiber break",
    "Multiple Fiber break",
    "DWDM outage",
    "Submarine cable cut",
    "Frequency interference",
]

TRANSMISSION_ACTIONS = [
    "Concerned team engaged",
    "Fault localization in progress",
    "Team moving to fault location",
    "Fiber splicing in progress",
    "Restoration activity in progress",
    "Rerouting feasibility under evaluation",
    "Rerouting initiated",
    "Priority links being rerouted",
    "Node power restoration in progress",
    "Service under monitoring",
]

REROUTING_OPTIONS = [
    "Not Applicable",
    "Feasibility under evaluation",
    "Rerouting initiated",
    "Rerouting completed",
    "Not possible – only MW path available",
    "No alternate path available",
]


def transmission_fault_sentence(fault_type):
    """Helper to convert transmission fault type into clean text sentence."""
    mapping = {
        "Node offline": "the serving node is currently offline",
        "Power issue at Node": "a power issue at the serving node",
        "Single Fiber break": "a single fiber break in the transmission segment",
        "Dual Fiber break": "a dual fiber break in the transmission segment",
        "Triple Fiber break": "a triple fiber break in the transmission segment",
        "Spur Fiber break": "a spur fiber break in the transmission segment",
        "Multiple Fiber break": "multiple fiber breaks in the transmission segment",
        "DWDM outage": "an outage in the DWDM transmission segment",
        "Submarine cable cut": "an upstream/submarine cable cut",
        "Frequency interference": "frequency interference on the wireless transmission path",
    }
    return mapping.get(fault_type, "a transmission-related fault")


def transmission_action_sentence(action):
    """Helper to convert transmission action into clean sentence."""
    mapping = {
        "Concerned team engaged": "Our concerned team is actively engaged and working on the restoration.",
        "Fault localization in progress": "Fault localization is currently in progress.",
        "Team moving to fault location": "The field/transmission team is moving towards the fault location.",
        "Fiber splicing in progress": "Fiber splicing/restoration activity is currently in progress.",
        "Restoration activity in progress": "Restoration activity is currently in progress on priority.",
        "Rerouting feasibility under evaluation": "Rerouting feasibility is being evaluated to minimize the service impact.",
        "Rerouting initiated": "Rerouting has been initiated to restore the affected services through an alternate path.",
        "Priority links being rerouted": "Priority corporate links are being rerouted gradually.",
        "Node power restoration in progress": "The concerned team is working to restore power at the affected node.",
        "Service under monitoring": "The service is currently under monitoring for stability.",
    }
    return mapping.get(action, "")


def rerouting_sentence(option):
    """Helper to generate rerouting status sentence."""
    mapping = {
        "Not Applicable": "",
        "Feasibility under evaluation": "Rerouting feasibility is currently under evaluation.",
        "Rerouting initiated": "Rerouting has been initiated for the affected services.",
        "Rerouting completed": "Rerouting has been completed for the affected service(s).",
        "Not possible – only MW path available": "Rerouting is currently not possible as only the microwave path is available.",
        "No alternate path available": "No alternate transmission path is currently available for rerouting.",
    }
    return mapping.get(option, "")


def generate_transmission_email(
    label,
    affected_links,
    fault_type,
    action_status,
    rerouting_status,
    fault_location,
    ettr_choice,
    custom_ettr,
    priority,
    audience,
):
    """Generates transmission mass outage email updates for multiple links."""
    label = clean(label)
    links = multiline_links(affected_links)

    if not label and not links:
        return "", "⚠️ Enter a service label or paste affected service labels.", ""

    primary = label or links[0]
    ensure_incident(primary, fault_type, "", action_status)
    update_incident_stage(primary, action_status)

    subject_topic = fault_type
    if len(links) > 1:
        subject_topic = f"Mass Outage - {fault_type}"

    subject = make_subject(primary, subject_topic, priority)
    hello = salutation(audience)

    fault_sentence = transmission_fault_sentence(fault_type)
    action_sentence = transmission_action_sentence(action_status)
    rr_sentence = rerouting_sentence(rerouting_status)

    lines = [
        hello,
        "",
        f"Please be informed that the reported service is currently impacted due to {fault_sentence}.",
    ]

    if clean(fault_location):
        lines.extend(["", f"Fault Location: {clean(fault_location)}"])

    if links:
        lines.extend(["", "Affected Services:"])
        for number, item in enumerate(links, start=1):
            lines.append(f"{number}. {item}")

    if action_sentence:
        lines.extend(["", action_sentence])

    if rr_sentence:
        lines.extend(["", rr_sentence])

    lines.extend([
        "",
        "We are closely following up with the concerned team and will share further restoration updates accordingly."
    ])

    ettr_line = format_ettr(ettr_choice, custom_ettr)
    if ettr_line:
        lines.extend(["", ettr_line])

    lines.extend([
        "",
        "Your patience and cooperation are highly appreciated."
    ])

    internal_note = ""
    if fault_type == "Submarine cable cut":
        internal_note = (
            "Suggested handling: mention the upstream carrier/cable name only "
            "when it has been formally confirmed by the concerned team."
        )

    return subject, "\n".join(lines), internal_note


# ============================================================
# 9. CUSTOMER-END & FINDINGS ENGINE
# ============================================================
FINDING_OPTIONS = [
    "No problematic alarm observed at Node end",
    "Port is up",
    "Tunnel is intact",
    "ARP is not observed",
    "MAC is not observed",
    "User is offline at BRAS",
    "Optical power is optimal",
    "Traffic is observed at our end",
    "Latency is optimal",
    "No abnormality observed in transmission path",
]

CUSTOMER_ACTIONS = [
    "Verify last-mile media",
    "Check CPE/router and power",
    "Check SFP/optical module and equipment port",
    "Reconnect the affected PPPoE user",
    "Test through standalone laptop/device",
    "Bypass LAN and perform direct testing",
    "Share active POC for joint troubleshooting",
    "Share problematic stats for further investigation",
]


def finding_bullet(finding, service_type):
    if finding == "Traffic is observed at our end" and service_type == "Turbonet":
        return "Traffic is being observed at our end."
    return finding + "."


def customer_action_sentence(action):
    mapping = {
        "Verify last-mile media": "Kindly verify the service from the last-mile media and share your feedback.",
        "Check CPE/router and power": "Kindly verify the CPE/router status and power availability at your end.",
        "Check SFP/optical module and equipment port": "Kindly check the SFP/optical module and equipment port at your end and share updated status.",
        "Reconnect the affected PPPoE user": "Kindly reconnect the affected PPPoE user and confirm whether the session comes online.",
        "Test through standalone laptop/device": "Kindly perform testing through a standalone laptop/device to isolate any internal LAN impact.",
        "Bypass LAN and perform direct testing": "Kindly bypass the internal LAN and perform direct testing from the edge device.",
        "Share active POC for joint troubleshooting": "Kindly share an active POC so joint troubleshooting can be arranged.",
        "Share problematic stats for further investigation": "If the issue persists, kindly share the problematic statistics for further investigation.",
    }
    return mapping.get(action, "")


def generate_customer_end_email(
    label,
    issue_summary,
    findings,
    requested_action,
    priority,
    ticket,
):
    """Generates response email when issue is identified at customer side."""
    label = clean(label)
    if not label:
        return "", "⚠️ Please enter the service label.", ""

    service_type = detect_service_type(label)
    ensure_incident(label, issue_summary, ticket, "Customer-end verification")
    update_incident_stage(label, "Customer-end verification")

    subject = make_subject(
        label,
        issue_summary or "Customer-End Verification",
        priority,
        ticket
    )

    lines = [
        "Dear Customer,",
        "",
        "We have thoroughly checked the reported service.",
        "",
        "Findings are shared below:",
    ]

    if findings:
        for item in findings:
            lines.append(f"• {finding_bullet(item, service_type)}")
    else:
        lines.append("• No abnormality is currently observed at our end.")

    action_line = customer_action_sentence(requested_action)
    if action_line:
        lines.extend(["", action_line])

    lines.extend([
        "",
        "Please share your feedback so we can proceed accordingly.",
        "",
        "Your cooperation in this matter will be highly appreciated."
    ])

    policy = ""
    if service_type == "Turbonet":
        policy = (
            "⚠️ Internal Turbo policy: Do not share the Turbonet graph with "
            "the customer. For graph-related requests, route the matter through "
            "KAM/GCSSQ."
        )

    return subject, "\n".join(lines), policy


# ============================================================
# 10. STATS & TROUBLESHOOTING ENGINE
# ============================================================
STATS_SCENARIOS = [
    "Packet Loss / High Latency",
    "Slow Speed / Throughput",
    "Turbo Slow Speed",
    "Website Issue",
    "Particular Website / URL Issue",
    "Banking Application Issue",
    "TikTok Issue",
    "WhatsApp / Meta Issue",
    "PPPoE / BRAS Issue",
    "Voice / PRI / SIP Issue",
    "Optical Power Issue",
    "Routing / BGP Issue",
]


def stats_items(scenario):
    mapping = {
        "Packet Loss / High Latency": [
            "Ping statistics from the affected user",
            "WinMTR / traceroute towards the problematic destination",
            "Source IP / testing IP pool",
            "Destination IP / URL",
            "Affected user details",
            "Exact issue occurrence date and time",
        ],
        "Slow Speed / Throughput": [
            "Speedtest result",
            "Speedtest server details",
            "Affected user/source IP",
            "Direct/standalone testing result after bypassing the LAN",
            "Current issue occurrence time",
        ],
        "Turbo Slow Speed": [
            "Direct speed and browsing test through a standalone laptop/device",
            "Affected PPPoE user / source IP",
            "Speedtest server details and result",
            "Issue occurrence time",
        ],
        "Website Issue": [
            "Source IP pool",
            "Browser error screenshot",
            "Traceroute / WinMTR towards the problematic destination",
            "NSLOOKUP result of the problematic destination",
            "NSLOOKUP result of whoami.akamai.com",
            "Tracetcp towards the affected URL on port 443",
            "Primary and secondary DNS configuration",
        ],
        "Particular Website / URL Issue": [
            "Snapshot of whatismyipaddress.com",
            "Source IP pool",
            "Browser error screenshot",
            "Traceroute towards the problematic destination",
            "WinMTR / tracetcp towards the affected URL on port 443",
            "NSLOOKUP result of the problematic destination",
            "Primary and secondary DNS configuration",
        ],
        "Banking Application Issue": [
            "Snapshot of whatismyipaddress.com",
            "Source IP pool",
            "Browser/application error screenshot",
            "NSLOOKUP / traceroute / WinMTR details where applicable",
            "PCAP trace from the affected Android phone using PCAPdroid",
        ],
        "TikTok Issue": [
            "DNS server IP configured at the edge device",
            "NSLOOKUP whoami.akamai.com screenshot",
            "Standalone test using Ethernet cable and a test PPPoE user",
            "Windows Task Manager > Performance > Ethernet screenshot while playing a TikTok video",
            "Affected user/source IP and issue occurrence time",
        ],
        "WhatsApp / Meta Issue": [
            "Affected user/source IP",
            "Issue occurrence time",
            "Ping / WinMTR towards a reachable problematic destination if available",
            "DNS configuration",
            "Standalone testing result after bypassing the LAN",
        ],
        "PPPoE / BRAS Issue": [
            "PPPoE username",
            "Affected user MAC address",
            "CPE/router status",
            "Exact disconnect/login time",
            "Standalone dial/test result",
            "Screenshot/error message if authentication fails",
        ],
        "Voice / PRI / SIP Issue": [
            "Party-A number",
            "Party-B number",
            "Exact call date and time",
            "Incoming / outgoing call scenario",
            "Observed error / announcement",
            "CLI displayed at Party-B where applicable",
        ],
        "Optical Power Issue": [
            "Current TX/RX optical power readings",
            "SFP/optical module details",
            "Equipment port status",
            "Patch-cord/fiber inspection result",
            "Relevant alarm/screenshot for reference",
        ],
        "Routing / BGP Issue": [
            "Source IP pool / affected prefix",
            "Problematic destination IP / URL",
            "Traceroute / WinMTR",
            "Current route advertisement details",
            "Expected upstream/path if known",
        ],
    }
    return mapping.get(scenario, ["Relevant problematic statistics and timestamps"])


def generate_stats_request(label, scenario, audience, custom_context):
    """Generates required technical stats email request."""
    label = clean(label)
    subject = make_subject(
        label or scenario,
        f"Required Stats - {scenario}",
        "Normal"
    )

    lines = [
        salutation(audience),
        "",
        "To proceed with detailed investigation of the reported issue, kindly share the following information/statistics:",
        "",
    ]

    for item in stats_items(scenario):
        lines.append(f"• {item}")

    if clean(custom_context):
        lines.extend(["", clean(custom_context)])

    if scenario == "Banking Application Issue":
        lines.extend([
            "",
            "PCAPdroid capture method:",
            "Select PCAP file → select the target application → START capture → reproduce the banking-app issue/login error → STOP capture → share the generated PCAP file."
        ])

    lines.extend([
        "",
        "Once the required information is received, we will proceed with further investigation accordingly.",
        "",
        "Your cooperation is highly appreciated."
    ])

    policy = ""
    if scenario == "Turbo Slow Speed":
        policy = (
            "⚠️ Internal Turbo policy: Never share the Turbonet graph directly "
            "with the customer. If the customer requests the graph, route the "
            "request to KAM; graph-policy escalation should be handled with GCSSQ."
        )

    return subject, "\n".join(lines), policy


# ============================================================
# 11. CUSTOMER FOLLOW-UP ENGINE
# ============================================================
CUSTOMER_FOLLOWUPS = [
    "Verify service after restoration",
    "Awaiting requested stats",
    "Awaiting active POC / access",
    "Customer-end verification pending",
    "Awaiting acknowledgement / confirmation",
    "Second follow-up - no response",
    "Retest through standalone device",
    "Request complete link details",
]


def generate_customer_followup(label, followup_type, previous_request, priority):
    """Generates standard customer follow-up email."""
    label = clean(label)
    subject = make_subject(label or "Customer", followup_type, priority)

    mapping = {
        "Verify service after restoration":
            "The reported service is currently up and operating normally at our end. Kindly verify the service status and share your feedback/confirmation.",
        "Awaiting requested stats":
            "We are still awaiting the requested problematic statistics/details from your end. Kindly share the required information so we can proceed with further investigation.",
        "Awaiting active POC / access":
            "Kindly share an active POC and arrange the required access so joint/onsite troubleshooting can be carried out without further delay.",
        "Customer-end verification pending":
            "Our checks are currently normal at our end. Kindly complete the requested customer-end/last-mile verification and share updated status.",
        "Awaiting acknowledgement / confirmation":
            "We are awaiting your acknowledgement/confirmation on the current service status. Kindly verify and update us accordingly.",
        "Second follow-up - no response":
            "This is a follow-up regarding the reported issue. We are still awaiting the requested information/confirmation from your end. Kindly update us so the case can be progressed accordingly.",
        "Retest through standalone device":
            "Kindly retest the service through a standalone device after bypassing the internal LAN and share the observed results.",
        "Request complete link details":
            "At the time of raising a complaint, kindly provide the complete link/service details along with an active POC. This will help avoid ambiguity and minimize restoration/troubleshooting delays.",
    }

    body = f"""Dear Customer,

{mapping.get(followup_type, "Kindly share the requested update.")}"""

    if clean(previous_request):
        body += f"\n\nPending / Required Information:\n{clean(previous_request)}"

    body += "\n\nYour cooperation in this regard will be highly appreciated."
    return subject, body


# ============================================================
# 12. PROGRESS UPDATE ENGINE
# ============================================================
PROGRESS_OPTIONS = [
    "Concerned team engaged",
    "Fault localization in progress",
    "Team dispatched",
    "Team reached site",
    "Fiber splicing in progress",
    "Rerouting in progress",
    "Configuration activity in progress",
    "Vendor engaged",
    "Customer coordination required",
    "Testing in progress",
    "Monitoring link stability",
]

PROGRESS_TEXT = {
    "Concerned team engaged": "Our concerned team is actively engaged and working on the reported issue.",
    "Fault localization in progress": "Fault localization is currently in progress and is being followed up on priority.",
    "Team dispatched": "The concerned field team has been dispatched for onsite troubleshooting/restoration.",
    "Team reached site": "The field team has reached the site and troubleshooting/restoration activity is in progress.",
    "Fiber splicing in progress": "Fiber splicing activity is currently in progress at the affected section.",
    "Rerouting in progress": "Rerouting activity is in progress to restore the affected service through an alternate path.",
    "Configuration activity in progress": "The concerned technical team is carrying out the required configuration activity.",
    "Vendor engaged": "The concerned vendor/upstream team has been engaged and is investigating the issue.",
    "Customer coordination required": "Further troubleshooting requires customer-side coordination. Kindly ensure an active POC is available.",
    "Testing in progress": "Detailed testing is currently in progress with the concerned teams.",
    "Monitoring link stability": "The link is currently under monitoring to verify stability and service performance.",
}


def generate_progress_email(
    label,
    progress_status,
    ettr_choice,
    custom_ettr,
    custom_note,
    priority,
    audience,
):
    """Generates progress status update email."""
    label = clean(label)
    if not label:
        return "", "⚠️ Please enter the service label."

    update_incident_stage(label, progress_status)
    subject = make_subject(label, f"Progress Update - {progress_status}", priority)

    body = f"""{salutation(audience)}

Please find the latest update regarding the reported service:

{PROGRESS_TEXT.get(progress_status, progress_status)}"""

    if clean(custom_note):
        body += f"\n\n{clean(custom_note)}"

    ettr_line = format_ettr(ettr_choice, custom_ettr)
    if ettr_line:
        body += f"\n\n{ettr_line}"

    body += (
        "\n\nFurther updates will be shared accordingly."
        "\n\nYour patience and cooperation are highly appreciated."
    )

    return subject, body


# ============================================================
# 13. VENDOR & INTERNAL ESCALATION ENGINE
# ============================================================
ESCALATION_TARGETS = [
    "Netsat",
    "Comstar",
    "VT",
    "GCS",
    "WISS",
    "Nexlinx",
    "Connect",
    "Other",
]

ESCALATION_LEVELS = [
    "Initial engagement",
    "Follow-up",
    "Urgent",
    "Critical",
    "Multiple follow-ups / no response",
    "Request exact delay reason / RFO",
]


def generate_escalation_email(
    label,
    target,
    escalation_level,
    findings,
    requested_action,
    priority,
    custom_target="",
):
    """Generates vendor / internal team escalation email and HTML matrix preview."""
    label_clean = clean(label)
    if not label_clean:
        return "", "⚠️ Please enter the service label.", ""

    target_name = custom_target.strip() if target == "Other" and custom_target else target
    greeting = f"Dear {target_name} Team,"

    template = """{greeting}

Please check the below-mentioned service, as the customer is currently facing a connectivity issue.

Service Details: {label}

Kindly investigate the issue, restore the service on priority, and share the ETTR at the earliest.

{extra_content}

Your prompt support and cooperation will be highly appreciated."""

    topic = f"{escalation_level} - {target_name}"
    subject = make_subject(label_clean, topic, priority)

    intro_map = {
        "Initial engagement": "",
        "Follow-up": "",
        "Urgent": f"Kindly prioritize the reported issue on urgent basis. The corporate customer is currently experiencing service impact/outage.",
        "Critical": f"The reported issue is critically impacting our corporate customer. Immediate engagement and restoration on top priority are required.",
        "Multiple follow-ups / no response": f"Despite repeated follow-ups, we have not received a progressive update from the {target_name} team. The continued delay is impacting customer communication and is not acceptable.",
        "Request exact delay reason / RFO": f"We have repeatedly requested a specific reason for the delay / clear RFO; however, only generic updates have been received. Kindly share the exact reason, defined action plan, and expected restoration timeline.",
    }
    extra_parts = [intro_map.get(escalation_level, "Kindly check and update.")]

    if findings:
        extra_parts.append("\nCurrent Findings:")
        if isinstance(findings, str):
            lines = [line.strip() for line in findings.splitlines() if line.strip()]
            for line in lines:
                extra_parts.append(f"• {line}")
        elif isinstance(findings, (list, tuple)):
            for item in findings:
                if item:
                    extra_parts.append(f"• {item}.")

    req_act_clean = clean(requested_action)
    if req_act_clean:
        extra_parts.append(f"\nRequired Action: {req_act_clean}")

    if escalation_level in {
        "Urgent",
        "Critical",
        "Multiple follow-ups / no response",
        "Request exact delay reason / RFO",
    }:
        extra_parts.append("\nKindly share a clear progressive update/action plan at the earliest.")

    extra_content = "\n".join(extra_parts)

    body = template.format(
        greeting=greeting,
        label=label_clean,
        extra_content=extra_content,
    )

    # Render HTML matrix live preview from unified JSON database
    matrix_html = render_escalation_matrix_html(target_name)

    return subject, body, matrix_html


# ============================================================
# 14. FIELD VISIT & ACCESS ENGINE
# ============================================================
FIELD_STATUSES = [
    "Team dispatched",
    "Team moving to site",
    "Team reached site",
    "Access required",
    "Customer POC required",
    "Visit scheduled",
    "FE details awaited",
    "Customer unavailable",
]


def generate_field_visit_email(
    label,
    field_status,
    eta,
    poc,
    access_details,
    audience,
    priority,
):
    """Generates field visit & site access request email."""
    label = clean(label)
    if not label:
        return "", "⚠️ Please enter the service label."

    subject = make_subject(label, field_status, priority)

    status_map = {
        "Team dispatched": "Our field team has been dispatched for onsite troubleshooting.",
        "Team moving to site": "Our field team is currently moving towards the site.",
        "Team reached site": "Our field team has reached the site and troubleshooting is in progress.",
        "Access required": "Site/access support is required to proceed with onsite troubleshooting.",
        "Customer POC required": "An active customer POC is required to proceed with joint troubleshooting.",
        "Visit scheduled": "An onsite troubleshooting visit has been scheduled.",
        "FE details awaited": "Field engineer details are currently awaited from the concerned team.",
        "Customer unavailable": "The onsite activity could not be completed as the customer/POC was unavailable.",
    }

    lines = [
        salutation(audience),
        "",
        status_map.get(field_status, field_status),
        "",
        f"Service Details: {label}",
    ]

    if clean(eta):
        lines.append(f"ETA / Visit Time: {clean(eta)}")
    if clean(poc):
        lines.append(f"POC: {clean(poc)}")
    if clean(access_details):
        lines.extend(["", f"Access / Coordination Details: {clean(access_details)}"])

    lines.extend([
        "",
        "Kindly extend the required coordination/support so the issue can be concluded at the earliest."
    ])

    return subject, "\n".join(lines)


# ============================================================
# 15. VPBX / VOICE / PRI / SIP ENGINE
# ============================================================
VOICE_ISSUES = [
    "Incoming (Call Landing)",
    "Outgoing (Call Landing)",
    "Incoming (Ext. Call Landing)",
    "Outgoing (Ext. Call Landing)",
    "Incoming Distortion",
    "Outgoing Distortion",
    "Incoming Ext. Distortion",
    "Outgoing Ext. Distortion",
    "OMO Issues",
    "Master No. Issues",
    "IVR Sound Modification",
    "IVR Routing Setting",
    "IVR Remove",
    "Extension Addition",
    "Extension Removal",
    "Account Defination",
    "Account Related (Info)",
    "In Portal Modifications",
    "CAP Portal Issue",
    "Portal Login / Credentials",
    "Call routing",
    "Ring All",
    "SIP trunk / PRI down",
    "Other VPBX / Voice Issue",
]

VPBX_TARGETS = [
    "Customer",
    "CONVEX",
    "NSS Team",
    "Product Team",
    "Seller / Partner",
]


def vpbx_required_details(issue):
    if issue in {
        "Incoming Distortion",
        "Outgoing Distortion",
        "Incoming Ext. Distortion",
        "Outgoing Ext. Distortion",
    }:
        return (
            "Kindly share Party-A, Party-B, exact call date/time, call direction, "
            "affected extension/master number, and a brief description of the distortion "
            "for detailed investigation."
        )

    if issue in {
        "Incoming (Call Landing)",
        "Outgoing (Call Landing)",
        "Incoming (Ext. Call Landing)",
        "Outgoing (Ext. Call Landing)",
        "OMO Issues",
        "Master No. Issues",
        "Call routing",
        "SIP trunk / PRI down",
    }:
        return (
            "Kindly share Party-A, Party-B, exact call date/time, call direction, "
            "affected extension/master number, and the observed error/announcement "
            "for detailed investigation."
        )

    if issue == "IVR Sound Modification":
        return (
            "Kindly share the required IVR audio/recording, master number, and the exact "
            "placement/flow where the recording is required."
        )

    if issue in {"IVR Routing Setting", "IVR Remove", "Ring All"}:
        return (
            "Kindly share the required IVR/call-flow instructions, affected master number, "
            "extensions, sequence/priority, and working/off-time requirement where applicable."
        )

    if issue in {"Extension Addition", "Extension Removal"}:
        return (
            "Kindly share the master number, extension number(s), mobile/FLL number(s), "
            "and the required routing/sequence details."
        )

    if issue in {"CAP Portal Issue", "Portal Login / Credentials", "In Portal Modifications"}:
        return (
            "Kindly share the master number/account, portal username where applicable, "
            "error screenshot, and the exact portal function/modification required."
        )

    if issue in {"Account Defination", "Account Related (Info)"}:
        return (
            "Kindly share the master number/account details and the exact account information "
            "or configuration required."
        )

    return (
        "Kindly share the affected master number/extension, exact issue details, "
        "testing time, and relevant evidence for detailed investigation."
    )


def generate_voice_email(
    label,
    voice_issue,
    party_a,
    party_b,
    call_time,
    call_direction,
    affected_extension,
    observed_behavior,
    response_to,
    priority,
):
    """Generates VPBX / Voice / SIP email updates."""
    label = clean(label)
    service_reference = label or "VPBX / Voice Service"
    subject = make_subject(service_reference, voice_issue, priority)

    party_a = clean(party_a)
    party_b = clean(party_b)
    call_time = clean(call_time)
    call_direction = clean(call_direction)
    affected_extension = clean(affected_extension)
    observed_behavior = clean(observed_behavior)

    details = []
    if party_a:
        details.append(f"Party-A: {party_a}")
    if party_b:
        details.append(f"Party-B: {party_b}")
    if call_time:
        details.append(f"Call Date/Time: {call_time}")
    if call_direction and call_direction != "Not Applicable":
        details.append(f"Call Direction: {call_direction}")
    if affected_extension:
        details.append(f"Affected Extension/Master: {affected_extension}")
    if observed_behavior:
        details.append(f"Observed Behavior / Request: {observed_behavior}")

    if response_to == "Customer":
        lines = [
            "Dear Customer,",
            "",
            f"We acknowledge your reported {voice_issue.lower()}.",
        ]

        if label:
            lines.extend(["", f"Service Details: {label}"])

        if details:
            lines.extend(["", "Available Details:"] + details)

        lines.extend([
            "",
            vpbx_required_details(voice_issue),
            "",
            "Once the required details are received, we will proceed with further investigation and update you accordingly.",
        ])

        if label in incident_records:
            update_incident_stage(
                label,
                f"VPBX customer information requested - {voice_issue}",
            )

        return subject, "\n".join(lines)

    target_name = response_to
    lines = [
        "Dear Team,",
        "",
        f"Customer is reporting {voice_issue.lower()} on the below VPBX/voice service.",
    ]

    if label:
        lines.extend(["", f"Service Details: {label}"])

    if details:
        lines.extend(["", "Complaint / Call Details:"] + details)
    else:
        lines.extend(["", vpbx_required_details(voice_issue)])

    if response_to == "CONVEX":
        action = (
            "Kindly investigate the issue at CONVEX end and share your findings/update "
            "at the earliest for further customer communication."
        )
    elif response_to == "NSS Team":
        action = (
            "Kindly verify the NSS/routing side and share your findings/update at the earliest."
        )
    elif response_to == "Product Team":
        action = (
            "Kindly review the requested VPBX/account configuration and share the required "
            "approval/findings at the earliest."
        )
    else:
        action = (
            "Kindly review the reported VPBX/voice requirement and share your findings/update "
            "at the earliest."
        )

    lines.extend(["", action])

    if label in incident_records:
        update_incident_stage(label, f"VPBX escalated to {target_name}")

    return subject, "\n".join(lines)


# ============================================================
# 16. RCA & CLOSURE ENGINE
# ============================================================
RCA_OPTIONS = [
    "Fiber Break (Dual/Triple)",
    "Fiber Break (Single)",
    "Fiber Break (Spur)",
    "Power Issue",
    "Power Issue at PTN / Node",
    "Port Down",
    "Port Problematic",
    "Customer related issue",
    "Customer own last mile",
    "Customer end power issue",
    "Hardware Faulty",
    "Configurations Issue",
    "Link Choking / Over-utilization",
    "Frequency Interference",
    "Optical Power Issue",
    "Vendor / Upstream Issue",
    "Routing Issue",
    "Submarine Cable Cut",
    "VPBX - No Issue Found after testing",
    "VPBX - IVR Updated",
    "VPBX - Extension Added",
    "VPBX - Extension Removed",
    "VPBX - Call Routing Updated",
    "VPBX - Account Defined",
    "VPBX - Portal Configuration Updated",
    "VPBX - OMO Issue",
    "VPBX - NSS Routing Issue",
    "VPBX - Customer Extension/Device Issue",
    "No Issue Observed",
    "Other / Manual",
]

ISSUE_FOUND_AT = [
    "Customer",
    "CONVEX",
    "NSS TEAM",
    "PRODUCT TEAM",
    "SELLER",
    "NOMC - Transmission Optical",
    "Vendor",
    "Region - Field Operations (FOPs)",
    "NOMC - IP Core",
    "Uplink ISP",
    "NOMC - Transmission Microwave",
    "TXN PM",
    "TXN CM",
    "RCA Awaited",
    "Other",
]

AUTO_ACTIONS = {
    "Fiber Break (Dual/Triple)": "Affected fiber section was restored/spliced and service connectivity was normalized.",
    "Fiber Break (Single)": "Affected fiber section was restored/spliced and service connectivity was normalized.",
    "Fiber Break (Spur)": "Spur fiber fault was restored and service connectivity was normalized.",
    "Power Issue": "Power was restored at the affected site/node and service became operational.",
    "Power Issue at PTN / Node": "Power was restored at the affected PTN/node and service became operational.",
    "Port Down": "Port connectivity was restored and the service became operational.",
    "Port Problematic": "The problematic port/connectivity was rectified and service was restored.",
    "Customer related issue": "Service became operational after customer-end verification/restoration.",
    "Customer own last mile": "The customer/last-mile issue was rectified and service became operational.",
    "Customer end power issue": "Customer-end power was restored and service became operational.",
    "Hardware Faulty": "The faulty hardware/equipment was rectified/replaced and service was restored.",
    "Configurations Issue": "Required configuration was rectified and service was restored.",
    "Link Choking / Over-utilization": "Bandwidth/utilization was normalized and the service was verified.",
    "Frequency Interference": "Wireless parameters/path were optimized and service stability was restored.",
    "Optical Power Issue": "Optical parameters were normalized and service connectivity was restored.",
    "Vendor / Upstream Issue": "The upstream/vendor issue was resolved and service was restored.",
    "Routing Issue": "Routing/path configuration was rectified and service was restored.",
    "Submarine Cable Cut": "Upstream capacity/path was restored or stabilized after the submarine cable incident.",
    "VPBX - No Issue Found after testing": "Live testing was performed and no abnormality was observed during the testing window.",
    "VPBX - IVR Updated": "The requested IVR configuration/recording was updated successfully.",
    "VPBX - Extension Added": "The requested extension(s) were added and the configuration was updated.",
    "VPBX - Extension Removed": "The requested extension(s) were removed and the configuration was updated.",
    "VPBX - Call Routing Updated": "The required call routing/IVR flow was updated and verified.",
    "VPBX - Account Defined": "The requested VPBX account/configuration was defined successfully.",
    "VPBX - Portal Configuration Updated": "The requested portal configuration/settings were updated successfully.",
    "VPBX - OMO Issue": "The OMO-related issue was rectified/cleared and call connectivity was restored.",
    "VPBX - NSS Routing Issue": "The NSS/routing issue was rectified and call connectivity was restored.",
    "VPBX - Customer Extension/Device Issue": "The issue was found at the customer extension/device side and service was verified after customer-end correction.",
    "No Issue Observed": "No activity was performed at our end; the service was already found operational.",
}


def generate_closure_email(
    selected_complaint,
    label,
    issue_found_at,
    custom_issue_found_at,
    root_cause_option,
    custom_root_cause,
    corrective_action_option,
    custom_corrective_action,
    final_status,
    priority,
    closure_time_mode,
    closure_date,
    closure_hour,
    closure_minute,
    closure_ampm,
):
    """Generates closure RFO email and calculates total impact duration automatically."""
    label = resolve_complaint_label(selected_complaint, label)
    if not label:
        return "", "", "", "⚠️ Please select an OPEN complaint or enter the service label."

    record = incident_records.get(label)
    service_type = (
        record.get("service_type")
        if record and record.get("service_type")
        else detect_service_type(label)
    )

    root_cause = (
        clean(custom_root_cause) or "No issue observed"
        if root_cause_option == "Other / Manual"
        else root_cause_option
    )

    found_at = (
        clean(custom_issue_found_at)
        if issue_found_at == "Other"
        else issue_found_at
    )

    if corrective_action_option == "Auto from Root Cause":
        corrective_action = AUTO_ACTIONS.get(
            root_cause,
            "Required corrective action was completed and service was restored.",
        )
    elif corrective_action_option == "Other / Manual":
        corrective_action = (
            clean(custom_corrective_action)
            or "Required corrective action was completed."
        )
    else:
        corrective_action = corrective_action_option

    try:
        restoration_time = resolve_click_clock_time(
            closure_time_mode,
            closure_date,
            closure_hour,
            closure_minute,
            closure_ampm,
        )
    except (ValueError, TypeError) as exc:
        return label, "", "", f"⚠️ Invalid closing date/time: {exc}"

    no_issue_root_causes = [
        "No Issue Observed",
        "No Issue Found after testing",
        "VPBX - No Issue Found after testing",
    ]

    customer_end_indicators = [
        "Customer related issue",
        "Customer own last mile",
        "Customer end power issue",
        "VPBX - Customer Extension/Device Issue",
        "Customer",
    ]

    is_no_issue_at_noc = False

    if root_cause in no_issue_root_causes:
        is_no_issue_at_noc = True

    if root_cause_option == "Other / Manual" and clean(custom_root_cause):
        custom_rc_lower = clean(custom_root_cause).lower()
        if any(phrase in custom_rc_lower for phrase in ["no issue", "no fault", "no abnormality", "no problem"]):
            is_no_issue_at_noc = True

    if found_at in customer_end_indicators:
        is_no_issue_at_noc = True

    no_action_phrases = [
        "No activity performed at our end",
        "no activity was performed",
        "already found operational",
        "already operating normally",
    ]
    if any(phrase in corrective_action.lower() for phrase in no_action_phrases):
        is_no_issue_at_noc = True

    if record and record.get("reported_time"):
        reported_time = record["reported_time"]

        if restoration_time < reported_time:
            return (
                label,
                "",
                "",
                f"⚠️ Closing time ({format_dt(restoration_time)}) cannot be earlier than "
                f"the saved opening time ({format_dt(reported_time)}).",
            )

        formatted_reported = format_dt(reported_time)

        if is_no_issue_at_noc:
            duration = "NA"
        else:
            duration = calculate_duration(reported_time, restoration_time)

        record["status"] = "CLOSED"
        record["last_stage"] = "Closed"
        record["restoration_time"] = restoration_time
        record["root_cause"] = root_cause
        record["issue_found_at"] = found_at
        record["corrective_action"] = corrective_action
    else:
        formatted_reported = "NA"
        duration = "NA"

        if record:
            record["status"] = "CLOSED"
            record["last_stage"] = "Closed without saved opening time"
            record["restoration_time"] = restoration_time

    subject = make_subject(label, "Service Restored / Closure", priority)

    if is_no_issue_at_noc:
        email = f"""Dear Customer,

We have thoroughly checked the reported service and no abnormality is currently observed at our end.

Incident Summary:
Service Details: {label}
Issue Found At: {found_at or "NA"}
Root Cause: {root_cause}
Corrective Action: {corrective_action}
Reported Time: {formatted_reported}
Verification Time: {format_dt(restoration_time)}
Service Impact Duration: NA

Note: As no issue was identified within our network, the service impact duration is not applicable.

The service is currently operating normally. Kindly verify and confirm if performance is satisfactory at your end.

We regret any inconvenience caused and appreciate your cooperation.

{final_status}"""
    else:
        email = f"""Dear Customer,

We are pleased to inform you that your {service_type} service is fully operational.

Incident Summary:
Service Details: {label}
Issue Found At: {found_at or "NA"}
Root Cause: {root_cause}
Corrective Action: {corrective_action}
Reported Time: {formatted_reported}
Restoration Time: {format_dt(restoration_time)}
Service Impact Duration: {duration}

The service is now operating normally. Kindly verify and confirm if performance is satisfactory at your end.

We regret any inconvenience caused and appreciate your cooperation throughout the restoration process.

{final_status}"""

    if record and record.get("reported_time"):
        if is_no_issue_at_noc:
            runtime = (
                f"✅ Complaint CLOSED at {format_dt(restoration_time)}. "
                f"Service Impact Duration: NA (No issue found at NOC end)."
            )
        else:
            runtime = (
                f"✅ Complaint CLOSED at {format_dt(restoration_time)}. "
                f"Total impact duration: {duration}."
            )
    elif record:
        runtime = (
            f"ℹ️ Complaint CLOSED at {format_dt(restoration_time)}, "
            "but no opening time was saved, so duration is NA."
        )
    else:
        runtime = (
            "ℹ️ Closure generated without a runtime complaint record. "
            "Reported Time and Duration are NA."
        )

    return label, subject, email, runtime


# ============================================================
# 17. OUTAGE ANALYZER ENGINE (WITH SHIFT HISTORY LOG)
# ============================================================
outage_history = []  # Runtime shift history store


def extract_bw_numeric(label):
    match = re.search(r'(\d+(?:\.\d+)?)\s*(?:M|G|K)BPS', str(label).replace('Mpbs', 'Mbps'), re.IGNORECASE)
    if match:
        val = float(match.group(1))
        if 'G' in match.group(0).upper():
            val *= 1000
        return val
    return 0


def extract_bw_text(label):
    match = re.search(r'(\d+(?:\.\d+)?)\s*((?:M|G|K)BPS)', str(label).replace('Mpbs', 'Mbps'), re.IGNORECASE)
    if match:
        return f"{match.group(1)} {match.group(2).upper()}"
    return "NA"


def detect_category(label):
    lbl = str(label).upper()
    if "MPLS" in lbl: return "MPLS"
    if "DPLC" in lbl: return "DPLC"
    if "DIA" in lbl: return "DIA"
    if "TURBO" in lbl or "TURBONET" in lbl: return "Turbonet"
    if "SIP" in lbl or "PRI" in lbl: return "SIP/PRI"
    if "VPBX" in lbl: return "VPBX"
    if "IPLC" in lbl: return "IPLC"
    if "M2M" in lbl: return "M2M"
    if "P2P" in lbl: return "P2P"
    return "Other"


def extract_client_name(label):
    parts = [p.strip() for p in str(label).strip().split("_") if p.strip()]
    if parts and "ESSCLIENT" in parts[0].upper():
        parts = parts[1:]
    bw_idx = next((i for i, p in enumerate(parts) if re.search(r'\d+(?:\.\d+)?\s*(?:M|G|K)BPS', p.replace('Mpbs', 'Mbps'), re.I)), None)
    c_parts = parts[:bw_idx] if bw_idx is not None else (parts[:-1] if len(parts) > 1 else parts)
    control_words = {"DIA", "MPLS", "DPLC", "M2M", "TURBO", "TURBONET", "BGP", "BGPDIA", "DIABGP", "SIP", "PRI", "SIPPRI"}
    clean_parts = [p for p in c_parts if p.upper() not in control_words and not p.isdigit()]
    return " ".join(clean_parts[:-1] if len(clean_parts) >= 2 else clean_parts).strip()


def extract_unique_id(label):
    parts = str(label).split('_')
    return parts[-1].strip() if len(parts) > 1 else ""


def load_and_clean_data(file_path):
    if file_path.lower().endswith(".csv"):
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            lines = f.readlines()
        header_idx = 0
        for i, line in enumerate(lines):
            if "Location Info" in line and "Last Occurred (ST)" in line:
                header_idx = i
                break
        clean_csv_data = "".join(lines[header_idx:])
        return pd.read_csv(io.StringIO(clean_csv_data), sep=",", engine="python")
    else:
        df = pd.read_excel(file_path, header=None)
        header_idx_list = df[df.apply(lambda r: r.astype(str).str.contains("Location Info", case=False).any(), axis=1)].index
        if not header_idx_list.empty:
            header_idx = header_idx_list[0]
            df.columns = df.iloc[header_idx]
            df = df.iloc[header_idx + 1:].reset_index(drop=True)
        return df


def process_outage_file(file_obj):
    """Processes uploaded CSV/Excel alarm dump files."""
    if file_obj is None:
        return "⚠️ Please upload an Excel or CSV file.", None, None, get_history_table(), generate_handover_summary()

    file_path = file_obj.name
    try:
        df = load_and_clean_data(file_path)
        df.columns = [str(c).strip() for c in df.columns]

        if "Location Info" not in df.columns or "Last Occurred (ST)" not in df.columns:
            return "❌ Error: 'Location Info' or 'Last Occurred (ST)' column not found.", None, None, get_history_table(), generate_handover_summary()

        processed = []
        for _, row in df.iterrows():
            raw_label = str(row["Location Info"]).strip()
            if pd.isna(raw_label) or raw_label.upper() == "NAN" or "VISIBILITY" in raw_label.upper() or not raw_label.upper().startswith("ESS"):
                continue

            processed.append({
                "Service": raw_label,
                "Bandwidth_Text": extract_bw_text(raw_label),
                "Numeric_BW": extract_bw_numeric(raw_label),
                "Client": extract_client_name(raw_label),
                "Category": detect_category(raw_label),
                "UniqueID": extract_unique_id(raw_label),
                "Last Occurred (ST)": row["Last Occurred (ST)"]
            })

        if not processed:
            return "❌ No valid ESS service records found in file.", None, None, get_history_table(), generate_handover_summary()

        df_unsorted = pd.DataFrame(processed).drop_duplicates(subset=["Service"], keep="first").reset_index(drop=True)
        df_sorted = df_unsorted.sort_values(by="Numeric_BW", ascending=False).reset_index(drop=True)

        df_sorted["Time_Parsed"] = pd.to_datetime(df_sorted["Last Occurred (ST)"], errors="coerce")
        valid_times = df_sorted["Time_Parsed"].dropna()
        outlook_time = valid_times.iloc[0].strftime("%Y-%m-%d %I:%M:%S %p") if not valid_times.empty else "Not Found"

        counts = Counter(df_sorted["Category"])
        summary = ", ".join([f"{c}x {s}" for s, c in sorted(counts.items(), key=lambda x: x[1], reverse=True)])

        priority_df = df_sorted[df_sorted["Numeric_BW"] > 250]
        normal_df = df_sorted[df_sorted["Numeric_BW"] <= 250]

        # Generate Console Text Output
        text_lines = []
        text_lines.append("=" * 80)
        text_lines.append(f"{'OUTLOOK READY TEXT':^80}")
        text_lines.append("=" * 80)
        text_lines.append(f"PTN Layer: {len(df_sorted)}x ({summary}) are down")
        text_lines.append("last occured time:")
        text_lines.append(outlook_time)
        text_lines.append("=" * 80 + "\n")

        text_lines.append("Transmission Section\n")
        text_lines.append("Priority Clients")

        for _, row in priority_df.iterrows():
            text_lines.append(row['Service'])

        text_lines.append("\n")

        for _, row in normal_df.iterrows():
            text_lines.append(row['Service'])

        text_lines.append("\nOutage Thread")
        text_lines.append("-" * 50)
        text_lines.append("\n")

        for _, row in df_unsorted.iterrows():
            text_lines.append(row['Service'])

        console_output = "\n".join(text_lines)

        # Append to Shift History Log
        now_str = datetime.now(PKT).strftime("%I:%M %p")
        outage_history.append({
            "id": len(outage_history) + 1,
            "processed_time": now_str,
            "file_name": os.path.basename(file_path),
            "total_links": len(df_sorted),
            "summary": summary,
            "priority_count": len(priority_df),
            "occurred_time": outlook_time,
        })

        # Setup Timestamps for Export Files
        file_time = valid_times.iloc[0].strftime("%Y-%m-%d_%I-%M_%p") if not valid_times.empty else datetime.now().strftime("%Y-%m-%d_%I-%M_%p")
        temp_dir = tempfile.gettempdir()

        # =========================================================
        # 1. FULL CATEGORIZED EXCEL REPORT
        # =========================================================
        out_filepath = os.path.join(temp_dir, f"Outage_Links_{file_time}.xlsx")
        wb = Workbook()
        ws = wb.active
        ws.title = "Categorized Links"

        fill = PatternFill(fill_type="solid", fgColor="D9D9D9")
        font = Font(bold=True)
        border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

        ws.append(["Part 1: Raw Unsorted Data"])
        ws.cell(row=1, column=1).font = Font(bold=True)
        headers = ["Service", "Bandwidth", "Client", "Category", "UniqueID", "Last Occurred (ST)"]
        ws.append(headers)

        for c in ws[2]:
            c.fill = fill
            c.font = font
            c.border = border

        for _, r in df_unsorted.iterrows():
            ws.append([r["Service"], r["Bandwidth_Text"], r["Client"], r["Category"], r["UniqueID"], r["Last Occurred (ST)"]])

        gap_row = ws.max_row + 3
        ws.cell(row=gap_row, column=1, value="Part 2: Sorted Data for Transmission Rerouting (Descending Bandwidth)")
        ws.cell(row=gap_row, column=1).font = Font(bold=True, italic=True)

        header_row = gap_row + 1
        ws.append(headers)
        for c in ws[header_row]:
            c.fill = fill
            c.font = font
            c.border = border

        for _, r in df_sorted.iterrows():
            bw_val = float(r["Numeric_BW"])
            bw_num = int(bw_val) if bw_val.is_integer() else bw_val
            if bw_num == 0: bw_num = "NA"
            time_fmt = r['Time_Parsed'].strftime("%Y-%m-%d %I:%M:%S %p") if pd.notna(r['Time_Parsed']) else r["Last Occurred (ST)"]
            ws.append([r["Service"], bw_num, r["Client"], r["Category"], r["UniqueID"], time_fmt])

        for col_idx, width in enumerate([75, 15, 35, 15, 25, 25], start=1):
            ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = width

        wb.save(out_filepath)

        # =========================================================
        # 2. SIMPLE OUTAGE LINKS EXCEL FILE (Single "Links" Column)
        # =========================================================
        simple_filepath = os.path.join(temp_dir, f"Outage_Raw_Links_{file_time}.xlsx")
        wb_simple = Workbook()
        ws_simple = wb_simple.active
        ws_simple.title = "Outage Links"

        # Header
        ws_simple.append(["Links"])
        ws_simple.cell(row=1, column=1).font = Font(bold=True)

        # Raw links list
        for _, r in df_unsorted.iterrows():
            ws_simple.append([r["Service"]])

        ws_simple.column_dimensions['A'].width = 80
        wb_simple.save(simple_filepath)

        # Returns 5 outputs now instead of 4
        return console_output, out_filepath, simple_filepath, get_history_table(), generate_handover_summary()

    except Exception as e:
        return f"❌ Processing Error: {str(e)}", None, None, get_history_table(), generate_handover_summary()



def get_history_table():
    rows = []
    for item in reversed(outage_history):
        rows.append([
            f"Outage #{item['id']}",
            item['processed_time'],
            item['total_links'],
            item['priority_count'],
            item['summary'],
            item['occurred_time'],
        ])
    return rows


def generate_handover_summary():
    if not outage_history:
        return "No outages processed in the current shift history."

    total_outages = len(outage_history)
    total_links = sum(item['total_links'] for item in outage_history)
    total_priority = sum(item['priority_count'] for item in outage_history)

    lines = [
        "==================================================",
        "📋 SHIFT HANDOVER SUMMARY / DAILY OUTAGE REPORT",
        "==================================================",
        f"• Total Outages Handled Today: {total_outages}",
        f"• Total Down Links Processed: {total_links}",
        f"• High Priority Links (>250M): {total_priority}",
        "--------------------------------------------------",
        "DETAILS BREAKDOWN:",
    ]

    for item in outage_history:
        lines.append(
            f"• Outage #{item['id']} ({item['processed_time']}): "
            f"{item['total_links']} Links Down ({item['summary']}) | "
            f"Occurred: {item['occurred_time']}"
        )

    lines.append("==================================================")
    return "\n".join(lines)


def reset_outage_history():
    global outage_history
    outage_history = []
    return [], "No outages processed in the current shift history."


# ============================================================
# 18. OUTAGE EMAIL ALERT DISPATCH ENGINE
# ============================================================
def compute_start_time_string(mode, date_str, hour_str, min_str, ampm_str):
    if mode == AUTO_TIME_MODE:
        now = datetime.now(PKT)
        return now.strftime("%Y-%m-%d %I:%M %p")
    else:
        d_str = date_str.split(" ")[0] if date_str else datetime.now(PKT).strftime("%Y-%m-%d")
        h = hour_str if hour_str else "12"
        m = min_str if min_str else "00"
        ap = ampm_str if ampm_str else "AM"
        return f"{d_str} {h}:{m} {ap}"


def generate_outage_clock_card_html(mode, date_str, hour_str, min_str, ampm_str):
    if mode == AUTO_TIME_MODE:
        now = datetime.now(PKT)
        d_display = now.strftime("%d-%m-%Y")
        h = int(now.strftime("%I"))
        m = int(now.strftime("%M"))
        formatted_time = now.strftime("%I:%M %p")
    else:
        try:
            d_parsed = datetime.strptime(date_str.split(" ")[0], "%Y-%m-%d")
            d_display = d_parsed.strftime("%d-%m-%Y")
        except Exception:
            d_display = datetime.now(PKT).strftime("%d-%m-%Y")

        try:
            h = int(hour_str) if hour_str else 12
            m = int(min_str) if min_str else 0
        except ValueError:
            h, m = 12, 0

        formatted_time = f"{h:02d}:{m:02d} {ampm_str if ampm_str else 'AM'}"

    h_angle = ((h % 12) + m / 60.0) * 30.0
    m_angle = m * 6.0
    cx, cy = 60, 60

    h_rad = math.radians(h_angle - 90)
    hx = cx + 20 * math.cos(h_rad)
    hy = cy + 20 * math.sin(h_rad)

    m_rad = math.radians(m_angle - 90)
    mx = cx + 30 * math.cos(m_rad)
    my = cy + 30 * math.sin(m_rad)

    numbers_svg = ""
    for i in range(1, 13):
        ang = math.radians(i * 30 - 90)
        nx = cx + 38 * math.cos(ang)
        ny = cy + 38 * math.sin(ang) + 3
        numbers_svg += f'<text x="{nx:.1f}" y="{ny:.1f}" font-size="8" fill="#ffffff" font-weight="bold" text-anchor="middle" font-family="sans-serif">{i}</text>'

    return f"""
    <div style="background-color: #1a1a1a; border-radius: 16px; padding: 20px 25px; color: white; display: flex; align-items: center; justify-content: space-between; max-width: 440px; margin-top: 10px; font-family: Arial, sans-serif; box-shadow: 0 4px 12px rgba(0,0,0,0.4);">
        <div>
            <div style="font-size: 32px; font-weight: bold; color: #ffffff; letter-spacing: 1px;">{formatted_time}</div>
            <div style="font-size: 13px; color: #cccccc; margin-top: 6px;">{d_display} · Pakistan (PKT)</div>
            <div style="font-size: 11px; color: #4a90e2; margin-top: 10px;">Mouse clock preview</div>
        </div>
        <div>
            <svg width="120" height="120" viewBox="0 0 120 120">
                <circle cx="60" cy="60" r="52" fill="#242424" stroke="#3a3a3c" stroke-width="2"/>
                {numbers_svg}
                <line x1="60" y1="60" x2="{hx:.1f}" y2="{hy:.1f}" stroke="#ffffff" stroke-width="3.5" stroke-linecap="round"/>
                <line x1="60" y1="60" x2="{mx:.1f}" y2="{my:.1f}" stroke="#ff5722" stroke-width="2.5" stroke-linecap="round"/>
                <circle cx="60" cy="60" r="4" fill="#ff5722"/>
            </svg>
        </div>
    </div>
    """


def toggle_outage_start_mode(mode):
    is_manual = (mode == MANUAL_TIME_MODE)
    return (
        gr.update(interactive=is_manual),
        gr.update(interactive=is_manual),
        gr.update(interactive=is_manual),
        gr.update(interactive=is_manual),
        gr.update(interactive=is_manual)
    )


def toggle_outage_custom_box(choice):
    if choice == "Other":
        return gr.update(visible=True)
    return gr.update(visible=False, value="")


def generate_dynamic_outage_table(region, status, event_detail, custom_detail, network_elements, ptn_impacted, otn_impacted, start_time_mode, start_date, start_hour, start_minute, start_ampm):
    final_event_detail = custom_detail if (event_detail == "Other" and custom_detail.strip()) else event_detail
    start_time = compute_start_time_string(start_time_mode, start_date, start_hour, start_minute, start_ampm)

    duration = ""
    reason = "Under investigation"
    reroute_status = "Initiated"

    ptn_text = f"PTN Layer: {ptn_impacted}" if ptn_impacted and not ptn_impacted.startswith("PTN Layer:") else (ptn_impacted if ptn_impacted else "PTN Layer: NA")
    otn_text = f"OTN Layer: {otn_impacted}" if otn_impacted and not otn_impacted.startswith("OTN Layer:") else (otn_impacted if otn_impacted else "NA")

    header_message = f"Please note that we are observing an outage in the <b>{region}</b> region. Details are given below for your reference:"

    return f"""
    <div style="font-family: Calibri, Arial, sans-serif; font-size: 15px; color: #000; background:#ffffff; padding:20px; width: 100%; max-width: 900px; margin: 10px 0;">
        <p style="font-size: 16px; margin-bottom: 20px;">
            <b>Dear All</b><br><br>
            {header_message}
        </p>

        <table style="width: 100%; background:#ffffff; border-collapse: collapse; border: 1.5px solid #000;">
            <tr style="border: 1px solid #000;">
                <td style="background-color: #e2efda; font-weight: bold; width: 30%; padding: 8px 12px; border: 1px solid #000;">Event Details</td>
                <td style="background:#ffffff; padding: 8px 12px; border: 1px solid #000;">{final_event_detail}</td>
            </tr>
            <tr style="border: 1px solid #000;">
                <td style="background-color: #e2efda; font-weight: bold; padding: 8px 12px; border: 1px solid #000;">Network Elements</td>
                <td style="background:#ffffff; padding: 8px 12px; border: 1px solid #000; font-weight: bold;">{network_elements if network_elements else 'Long Haul Ring 6A 1st 10G'}</td>
            </tr>
            <tr style="border: 1px solid #000;">
                <td style="background-color: #e2efda; font-weight: bold; padding: 8px 12px; border: 1px solid #000;">Event Start Time</td>
                <td style="background:#ffffff; padding: 8px 12px; border: 1px solid #000;">{start_time}</td>
            </tr>
            <tr style="border: 1px solid #000;">
                <td style="background-color: #e2efda; font-weight: bold; padding: 8px 12px; border: 1px solid #000;">Event End Time</td>
                <td style="background:#ffffff; padding: 8px 12px; border: 1px solid #000;"></td>
            </tr>
            <tr style="border: 1px solid #000;">
                <td style="background-color: #e2efda; font-weight: bold; padding: 8px 12px; border: 1px solid #000;">Event Duration</td>
                <td style="background:#ffffff; padding: 8px 12px; border: 1px solid #000;">{duration}</td>
            </tr>
            <tr style="border: 1px solid #000;">
                <td rowspan="2" style="background-color: #e2efda; font-weight: bold; padding: 8px 12px; border: 1px solid #000; vertical-align: middle;">Impacted Clients</td>
                <td style="background:#ffffff; padding: 6px 12px; border: 1px solid #000;">{ptn_text}</td>
            </tr>
            <tr style="border: 1px solid #000;">
                <td style="background:#ffffff; padding: 6px 12px; border: 1px solid #000;">{otn_text}</td>
            </tr>
            <tr style="border: 1px solid #000;">
                <td style="background-color: #e2efda; font-weight: bold; padding: 8px 12px; border: 1px solid #000;">Reason For <span style="background-color: #ffff00; padding: 0 2px;">Outage</span></td>
                <td style="background:#ffffff; padding: 8px 12px; border: 1px solid #000;">{reason}</td>
            </tr>
            <tr style="border: 1px solid #000;">
                <td style="background-color: #e2efda; font-weight: bold; padding: 8px 12px; border: 1px solid #000;">Rerouting Status</td>
                <td style="background:#ffffff; padding: 8px 12px; border: 1px solid #000;">{reroute_status}</td>
            </tr>
            <tr style="border: 1px solid #000;">
                <td rowspan="2" style="background-color: #e2efda; font-weight: bold; padding: 8px 12px; border: 1px solid #000; vertical-align: middle;">Re-routed Clients</td>
                <td style="background:#ffffff; padding: 6px 12px; border: 1px solid #000;">PTN Layer: NA</td>
            </tr>
            <tr style="border: 1px solid #000;">
                <td style="background:#ffffff; padding: 6px 12px; border: 1px solid #000;">OTN Layer: NA</td>
            </tr>
        </table>
    </div>
    """


def send_outage_email(to_recipients, cc_recipients, subject, html_content):
    """Dispatches outage alert email via Gmail SMTP server."""
    if not to_recipients or not to_recipients.strip():
        return "❌ Error: 'To Recipients' field cannot be empty."

    sender_email = os.environ.get("SENDER_EMAIL", "").strip()
    sender_password = os.environ.get("SENDER_PASSWORD", "").strip()
    smtp_server = os.environ.get("SMTP_SERVER", "smtp.gmail.com").strip()
    smtp_port = int(os.environ.get("SMTP_PORT", 587))

    if not sender_email or not sender_password:
        return (
            "❌ Missing Email Credentials!\n"
            "Please configure SENDER_EMAIL & SENDER_PASSWORD in Render Environment Variables or local environment."
        )

    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = subject.strip() if subject and subject.strip() else "[OUTAGE ALERT] NOC Notification"
        msg["From"] = sender_email
        msg["To"] = to_recipients

        if cc_recipients and cc_recipients.strip():
            msg["Cc"] = cc_recipients

        msg.attach(MIMEText(html_content, "html"))

        to_list = [x.strip() for x in to_recipients.split(",") if x.strip()]
        cc_list = [x.strip() for x in cc_recipients.split(",") if x.strip()]
        all_recipients = to_list + cc_list

        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, all_recipients, msg.as_string())

        return f"✅ Email sent successfully from {sender_email}"

    except Exception as err:
        return f"❌ Transmission Failed:\n{err}"


# ============================================================
# 19. VENDOR ESCALATION MATRIX MANAGER (JSON PERSISTENCE)
# ============================================================
VENDOR_JSON_FILE = "vendors.json"

DEFAULT_VENDOR_MATRIX = {
    "Netsat": [
        {"Level": "Level 1", "Name": "Netsat Support Desk", "Designation": "Help desk / CNOC Netsat", "Time": "1 Hour", "Phone": "021-111638728", "Email": "support.cmpak@netsat.net.pk"},
        {"Level": "Level 1", "Name": "Ali Nadeem", "Designation": "Help desk / CNOC Netsat", "Time": "1 Hour", "Phone": "0309 1112422; 0312-2431968", "Email": "ali.nadeem@netsat.net.pk"},
        {"Level": "Level 2 (Central Region)", "Name": "Zafar Iqbal (CTR)", "Designation": "Regional head (Lahore)", "Time": "3 hours", "Phone": "0301-8114185", "Email": "zafar.iqbal@netsat.net.pk"},
        {"Level": "Level 2 (Central Region)", "Name": "Shoukat Khan (MTR)", "Designation": "Regional head (Multan)", "Time": "3 hours", "Phone": "0302-8271762", "Email": "shoukat.khan@netsat.net.pk"},
        {"Level": "Level 2 (South Region)", "Name": "Farhan (Sindh)", "Designation": "Regional head (Sindh)", "Time": "3 hours", "Phone": "0301-8114182", "Email": "m.farhan@netsat.net.pk"},
        {"Level": "Level 2 (South Region)", "Name": "Syed Yousuf Hussain (Balochistan)", "Designation": "Regional head (Balochistan)", "Time": "3 hours", "Phone": "0300-8259632", "Email": "syed.yousuf@netsat.net.pk"},
        {"Level": "Level 2 (North Region)", "Name": "Asad Imran (North/KPK)", "Designation": "Regional head (North)", "Time": "3 hours", "Phone": "0311-8814294", "Email": "asad.imran@netsat.net.pk"},
        {"Level": "Level 3", "Name": "M. Tahir Qureshi", "Designation": "Operational Support", "Time": "6 hours", "Phone": "0300 8232647", "Email": "m.tahir@netsat.net.pk"},
        {"Level": "Level 3", "Name": "Muhammad Sumair", "Designation": "Operation Head", "Time": "8 hours", "Phone": "0301-8114181", "Email": "sumair@netsat.net.pk"},
        {"Level": "Level 4", "Name": "Syed Mubashir Imam", "Designation": "Senior Operational Head", "Time": "12 hours", "Phone": "0301-8292632", "Email": "mubashir.imam@netsat.net.pk"},
        {"Level": "Level 5", "Name": "Kamran Qaiser", "Designation": "HOD Wireless", "Time": "Urgent/Emergency Maintenance", "Phone": "0301 8114177", "Email": "kamran.qaiser@netsat.net.pk"}
    ],
    "Comstar": [
        {"Level": "Level 1", "Name": "Support Desk", "Designation": "CS", "Time": "10-15min", "Phone": "0333-1312343", "Email": "cs@comstar.com.pk"},
        {"Level": "Level 2", "Name": "Abdul Wajid", "Designation": "TL (South)", "Time": "15-30min", "Phone": "0334-2594529", "Email": "awajid@comstar.com.pk"},
        {"Level": "Level 3", "Name": "Osman Javaid", "Time": "30-45min", "Designation": "Manager Engineering", "Phone": "0336-5479293", "Email": "ojavaid@comstar.com.pk"}
    ],
    "Vision Telecom": [
        {"Level": "Level 1", "Name": "Corporate Helpdesk", "Designation": "Support", "Time": "Immediate", "Phone": "0308-8881418", "Email": "support@visiontelecom.com.pk"},
        {"Level": "Level 2", "Name": "Rizwan Younis", "Designation": "Team Lead", "Time": "30min", "Phone": "0300-0807140", "Email": "rizwan.younis@visiontelecom.com.pk"}
    ]
}


def load_vendors_matrix():
    """Loads vendor matrix dictionary directly from JSON file, initializing defaults if missing."""
    if not os.path.exists(VENDOR_JSON_FILE):
        save_vendors_matrix(DEFAULT_VENDOR_MATRIX)
        return DEFAULT_VENDOR_MATRIX
    try:
        with open(VENDOR_JSON_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            return data if data else DEFAULT_VENDOR_MATRIX
    except Exception:
        return DEFAULT_VENDOR_MATRIX


def save_vendors_matrix(data):
    """Saves updated vendor matrix dictionary directly to vendors.json."""
    try:
        with open(VENDOR_JSON_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
    except Exception as e:
        print(f"Error saving vendors.json: {e}")


# Initialize dynamic vendor matrix database
vendors_matrix_db = load_vendors_matrix()


def get_vendor_list():
    """Returns dynamic list of all registered vendor names."""
    return list(vendors_matrix_db.keys())


def get_vendor_emails_string(vendor_name):
    """Extracts all emails for a vendor into semicolon-separated string for Outlook."""
    if not vendor_name or vendor_name not in vendors_matrix_db:
        return ""
    emails = []
    for row in vendors_matrix_db[vendor_name]:
        raw_email = row.get("Email", "")
        for e in re.split(r"[;, ]+", raw_email):
            clean_e = e.strip()
            if clean_e and clean_e not in emails:
                emails.append(clean_e)
    return "; ".join(emails)


def render_escalation_matrix_html(vendor_name):
    """Renders HTML matrix table with strict red header corporate styling and pinpoint email CSS selection."""
    if not vendor_name or vendor_name not in vendors_matrix_db:
        return "<p style='color: gray;'>No vendor selected or matrix empty.</p>"

    matrix = vendors_matrix_db[vendor_name]

    html = f"""
    <style>
        .matrix-table, 
        .matrix-table th, 
        .matrix-table td {{
            -webkit-user-select: none !important;
            -moz-user-select: none !important;
            -ms-user-select: none !important;
            user-select: none !important;
        }}

        .matrix-table td.email-cell, 
        .matrix-table td.email-cell a {{
            -webkit-user-select: text !important;
            -moz-user-select: text !important;
            -ms-user-select: text !important;
            user-select: text !important;
            cursor: text !important;
            color: #0000ee !important;
            font-weight: bold;
        }}
    </style>

    <div style="font-family: Arial, sans-serif; max-width: 100%; overflow-x: auto; margin-top: 10px;">
        <table class="matrix-table" style="width: 100%; border-collapse: collapse; border: 1.5px solid #000; font-size: 13px; text-align: left; background:#fff;">
            <thead>
                <tr style="background-color: #7b241c; color: white;">
                    <th colspan="6" style="padding: 10px; font-size: 15px; font-weight: bold; border-bottom: 2px solid #000; text-align:center;">
                        Escalation Matrix - {vendor_name}
                    </th>
                </tr>
                <tr style="background-color: #922b21; color: white; border-bottom: 1.5px solid #000;">
                    <th style="padding: 8px; border: 1px solid #7b241c; width: 18%;">Escalation levels</th>
                    <th style="padding: 8px; border: 1px solid #7b241c; width: 20%;">Name</th>
                    <th style="padding: 8px; border: 1px solid #7b241c; width: 20%;">Designation</th>
                    <th style="padding: 8px; border: 1px solid #7b241c; width: 14%;">Escalation Time</th>
                    <th style="padding: 8px; border: 1px solid #7b241c; width: 14%;">Contact Number</th>
                    <th style="padding: 8px; border: 1px solid #7b241c; width: 14%;">Email address</th>
                </tr>
            </thead>
            <tbody>
    """

    for row in matrix:
        email_val = row.get('Email', '')
        email_html = f'<a href="mailto:{email_val}">{email_val}</a>' if email_val else ''
        html += f"""
        <tr style="border: 1px solid #ccc; background-color: #ffffff; color:#000;">
            <td style="padding: 7px; border: 1px solid #333; font-weight: bold;">{row.get('Level', '')}</td>
            <td style="padding: 7px; border: 1px solid #333;">{row.get('Name', '')}</td>
            <td style="padding: 7px; border: 1px solid #333;">{row.get('Designation', '')}</td>
            <td style="padding: 7px; border: 1px solid #333;">{row.get('Time', '')}</td>
            <td style="padding: 7px; border: 1px solid #333;">{row.get('Phone', '')}</td>
            <td class="email-cell" style="padding: 7px; border: 1px solid #333;">{email_html}</td>
        </tr>
        """

    html += """
            </tbody>
        </table>
    </div>
    """
    return html


def get_matrix_dataframe(vendor_name):
    if not vendor_name or vendor_name not in vendors_matrix_db:
        return []
    return [[r.get("Level", ""), r.get("Name", ""), r.get("Designation", ""), r.get("Time", ""), r.get("Phone", ""), r.get("Email", "")] for r in vendors_matrix_db[vendor_name]]


def add_vendor_name(new_vendor_name):
    """Adds a new vendor entry and syncs to vendors.json."""
    if not new_vendor_name or not new_vendor_name.strip():
        return gr.update(choices=get_vendor_list()), "", "❌ Vendor name is required!"
    v_name = new_vendor_name.strip()
    if v_name not in vendors_matrix_db:
        vendors_matrix_db[v_name] = []
        save_vendors_matrix(vendors_matrix_db)
    return gr.update(choices=get_vendor_list(), value=v_name), get_vendor_emails_string(v_name), f"✅ Vendor '{v_name}' added & saved to vendors.json!"


def add_contact_to_matrix(vendor_name, level, name, designation, time, phone, email):
    """Adds contact person to specified escalation level and auto-saves to vendors.json."""
    if not vendor_name:
        return render_escalation_matrix_html(vendor_name), get_matrix_dataframe(vendor_name), get_vendor_emails_string(vendor_name), "⚠️ Please select a Vendor first!"

    if not name or not email:
        return render_escalation_matrix_html(vendor_name), get_matrix_dataframe(vendor_name), get_vendor_emails_string(vendor_name), "❌ Name and Email are required!"

    vendors_matrix_db[vendor_name].append({
        "Level": level,
        "Name": name.strip(),
        "Designation": designation.strip(),
        "Time": time.strip(),
        "Phone": phone.strip(),
        "Email": email.strip()
    })

    save_vendors_matrix(vendors_matrix_db)

    return (
        render_escalation_matrix_html(vendor_name),
        get_matrix_dataframe(vendor_name),
        get_vendor_emails_string(vendor_name),
        f"✅ Contact added & saved to vendors.json for {vendor_name}!"
    )


def delete_contact_from_matrix(vendor_name, name_to_delete):
    """Deletes contact person from matrix and auto-saves to vendors.json."""
    if not vendor_name or vendor_name not in vendors_matrix_db:
        return render_escalation_matrix_html(vendor_name), get_matrix_dataframe(vendor_name), get_vendor_emails_string(vendor_name), "⚠️ Vendor not found!"

    vendors_matrix_db[vendor_name] = [r for r in vendors_matrix_db[vendor_name] if r.get("Name") != name_to_delete.strip()]
    save_vendors_matrix(vendors_matrix_db)

    return (
        render_escalation_matrix_html(vendor_name),
        get_matrix_dataframe(vendor_name),
        get_vendor_emails_string(vendor_name),
        f"🗑️ Deleted '{name_to_delete}' & saved changes to vendors.json!"
    )


# ============================================================
# 20. ACTIVE DASHBOARD & MONITORING
# ============================================================
def refresh_dashboard(view_mode):
    records = []

    for label, record in incident_records.items():
        status = record.get("status", "QUEUED")

        if view_mode == "Open Only" and status != "OPEN":
            continue
        if view_mode == "Queued Only" and status != "QUEUED":
            continue
        if view_mode == "Closed Only" and status != "CLOSED":
            continue
        if view_mode == "Queued + Open" and status not in ("QUEUED", "OPEN"):
            continue

        added_time = record.get("added_time")
        reported_time = record.get("reported_time")
        restoration_time = record.get("restoration_time")

        if reported_time:
            age = incident_age_text(
                reported_time,
                added_time=added_time,
                status=status,
            )
        else:
            age = incident_age_text(
                None,
                added_time=added_time,
                status=status,
            )

        duration = (
            calculate_duration(reported_time, restoration_time)
            if reported_time and restoration_time
            else ""
        )

        records.append([
            extract_service_id(label),
            record.get("service_type", ""),
            record.get("issue_type", ""),
            format_dt(added_time) if added_time else "",
            format_dt(reported_time) if reported_time else "Not opened",
            age,
            format_dt(restoration_time) if restoration_time else "",
            duration,
            record.get("ticket", ""),
            record.get("last_stage", ""),
            status,
        ])

    records.sort(
        key=lambda row: row[3] or "",
        reverse=True,
    )
    return records


def check_incident(label):
    label = clean(label)
    if not label:
        return "⚠️ Please enter/select a service label."

    if label not in incident_records:
        return "❌ No runtime complaint record found for this label."

    record = incident_records[label]
    added = record.get("added_time")
    reported = record.get("reported_time")
    restored = record.get("restoration_time")

    duration = (
        calculate_duration(reported, restored)
        if reported and restored
        else "Running" if reported else "Not started"
    )

    return f"""Status: {record.get("status", "QUEUED")}
Service Type: {record.get("service_type", "")}
Issue Type: {record.get("issue_type", "")}
Added to Queue: {format_dt(added) if added else "NA"}
Opening Time: {format_dt(reported) if reported else "Not opened"}
Current Age: {incident_age_text(reported, added, record.get("status", "QUEUED"))}
Closing Time: {format_dt(restored) if restored else "Not closed"}
Total Duration: {duration}
Last Stage: {record.get("last_stage", "")}
Reference Ticket: {record.get("ticket", "")}"""


def reset_incident(label):
    label = clean(label)
    if not label:
        return "⚠️ Please enter/select a service label."

    if label in incident_records:
        del incident_records[label]
        return "✅ Runtime complaint record cleared."

    return "❌ No runtime complaint found for this label."


# ============================================================
# 21. GRADIO UI BUILDER
# ============================================================
def build_app():
    with gr.Blocks(title="Corporate NOC Outage Reporting Console") as app:
        gr.Markdown(
            """
# 🌐 Corporate NOC Outage Reporting Console
Runtime-only console for parallel complaint handling, outage alerts, vendor escalations, and daily handover reports.
            """
        )

        with gr.Row():
            label = gr.Textbox(
                label="Service Label / Link Name",
                placeholder="ESSClient_DIA_... / SITE-... / link name",
                lines=1,
            )
            common_ticket = gr.Textbox(
                label="Reference Ticket",
                placeholder="Optional",
                lines=1,
            )

        with gr.Tabs():
            # TAB 1: COMPLAINT MANAGER
            with gr.Tab("🗂️ Complaint Manager"):
                gr.Markdown("### Parallel Complaint Queue")

                bulk_complaints = gr.Textbox(
                    label="Add Multiple Complaints",
                    placeholder=(
                        "ESSClient_DPLC_A.A Network_Tandlianwala-FSD_900Mbps_DPLC14722SL1\n"
                        "ESSClient_Central_Turbo_AAA_Brouadband_Faisalabad_2_981Mbps_Turbo14648SL694\n"
                        "ESSClient_SIP_PRI_2_ACE Money Transfer_Kharian_2Mbps_SIP00042"
                    ),
                    lines=8,
                )

                with gr.Row():
                    complaint_default_issue = gr.Dropdown(
                        choices=COMPLAINT_DEFAULT_ISSUES,
                        value="Link is down.",
                        label="Default Complaint Type",
                    )
                    complaint_bulk_ticket = gr.Textbox(
                        label="Reference Ticket (Optional)",
                        placeholder="Applied to newly added complaints",
                    )

                add_complaints_button = gr.Button(
                    "Add Complaints to Runtime Queue",
                    variant="primary",
                )

                with gr.Row():
                    complaint_selector = gr.Dropdown(
                        choices=[],
                        label="Select Runtime Complaint",
                        info="Select a complaint, then load it into the main Service Label field.",
                    )
                    refresh_complaints_button = gr.Button("Refresh Complaint List")

                with gr.Row():
                    load_complaint_button = gr.Button("Load Selected Complaint", variant="primary")
                    remove_complaint_button = gr.Button("Remove Selected Complaint", variant="stop")

                complaint_manager_status = gr.Textbox(
                    label="Complaint Manager Status",
                    lines=5,
                    interactive=False,
                )

                complaint_manager_dashboard = gr.Dataframe(
                    headers=[
                        "Service", "Type", "Issue", "Added", "Opened",
                        "Age", "Closed", "Total Time", "Ticket", "Last Stage", "Status"
                    ],
                    datatype=["str"] * 11,
                    value=[],
                    interactive=False,
                    label="Queued + Open Complaints",
                )

# TAB 2: OUTAGE ANALYZER
            with gr.Tab("📈 Outage Analyzer"):
                gr.Markdown("### 📊 NOC Alarm & Outage Link Analyzer")
                with gr.Row():
                    with gr.Column():
                        file_input = gr.File(
                            label="Upload Alarm Dump File (CSV / Excel)",
                            file_types=[".csv", ".xlsx", ".xls"],
                        )
                        analyze_button = gr.Button("⚡ Analyze Outage File", variant="primary")
                    with gr.Column():
                        output_file = gr.File(label="📥 Download Categorized Excel Report")
                        simple_links_file = gr.File(label="📄 Download Simple Outage Links Only") # <-- NEW DOWNLOAD BUTTON

                console_output = gr.Textbox(
                    label="📋 Outlook Ready Summary & Thread Output",
                    lines=14,
                    interactive=False,
                )

                gr.Markdown("---")
                gr.Markdown("### 📜 Shift Outage History & Handover Summary")

                with gr.Row():
                    clear_history_btn = gr.Button("🗑️ Clear Shift History", variant="stop")

                history_table = gr.Dataframe(
                    headers=[
                        "Outage ID", "Processed Time", "Total Links Down",
                        "Priority (>250M)", "Categorized Breakdown", "Last Occurred Time"
                    ],
                    datatype=["str", "str", "number", "number", "str", "str"],
                    value=[],
                    interactive=False,
                    label="Shift Outage Log",
                )

                handover_text = gr.Textbox(
                    label="🗣️ Shift Handover Communication Text (Copy for Next Technical Shift)",
                    lines=8,
                    interactive=False,
                )

                # Connect the button to 5 output targets
                analyze_button.click(
                    fn=process_outage_file,
                    inputs=file_input,
                    outputs=[console_output, output_file, simple_links_file, history_table, handover_text],
                )

                clear_history_btn.click(
                    fn=reset_outage_history,
                    inputs=[],
                    outputs=[history_table, handover_text],
                )

            # TAB 3: OUTAGE EMAIL ALERT
            with gr.Tab("📧 Outage Email Alert"):
                gr.Markdown("### 📢 Outage Email Dispatch Console")
                with gr.Group():
                    gr.Markdown("#### ⚙️ Outage Control Form")

                    with gr.Row():
                        outage_region_dd = gr.Dropdown(
                            choices=["North", "South", "Central"],
                            label="Region",
                            value="North"
                        )
                        outage_status_dd = gr.Dropdown(
                            choices=["Occurred", "Restored"],
                            label="Event Status",
                            value="Occurred"
                        )
                        outage_event_detail_dd = gr.Dropdown(
                            choices=[
                                "Spur Fiber break", "Single Fiber Break", "Node offline",
                                "Dual Fiber Break", "Triple Fiber break", "Other"
                            ],
                            label="Event Details",
                            value="Single Fiber Break"
                        )

                    outage_custom_detail_tb = gr.Textbox(
                        label="Custom Event Detail",
                        placeholder="Enter custom outage reason...",
                        visible=False
                    )

                    outage_network_elements_tb = gr.Textbox(
                        label="Network Elements (Between two PTNs)",
                        value="Long Haul Ring 6A 1st 10G"
                    )

                    with gr.Row():
                        outage_ptn_imp_tb = gr.Textbox(
                            label="PTN Layer Impacted",
                            value="14x (4x DIA, 4x DPLC, 3x Other, 3x MPLS) are down"
                        )
                        outage_otn_imp_tb = gr.Textbox(
                            label="OTN Layer Impacted",
                            value="NA"
                        )

                    gr.Markdown("#### ⏱️ Event Start Time Selection")
                    outage_start_time_mode = gr.Radio(
                        choices=[AUTO_TIME_MODE, MANUAL_TIME_MODE],
                        value=MANUAL_TIME_MODE,
                        label="Start Time Mode",
                    )

                    with gr.Row():
                        outage_start_date = gr.DateTime(
                            value=datetime.now(PKT).strftime("%Y-%m-%d"),
                            include_time=False,
                            type="string",
                            label="Start Date",
                        )
                        outage_start_hour = gr.Dropdown(
                            choices=HOUR_CHOICES,
                            value=datetime.now(PKT).strftime("%I"),
                            label="Hour",
                        )
                        outage_start_minute = gr.Dropdown(
                            choices=MINUTE_CHOICES,
                            value=datetime.now(PKT).strftime("%M"),
                            label="Minute",
                        )
                        outage_start_ampm = gr.Radio(
                            choices=AMPM_CHOICES,
                            value=datetime.now(PKT).strftime("%p"),
                            label="AM / PM",
                        )

                    outage_start_set_current = gr.Button("🕐 Set Manual Picker to Current Pakistan Time")

                    outage_start_clock_preview = gr.HTML(
                        value=generate_outage_clock_card_html(
                            MANUAL_TIME_MODE,
                            datetime.now(PKT).strftime("%Y-%m-%d"),
                            datetime.now(PKT).strftime("%I"),
                            datetime.now(PKT).strftime("%M"),
                            datetime.now(PKT).strftime("%p")
                        )
                    )

                gr.Markdown("---")
                outage_table_output = gr.HTML()

                gr.Markdown("---")
                with gr.Group():
                    gr.Markdown("#### 📧 Email Dispatch Control")
                    outage_email_subject_tb = gr.Textbox(
                        label="Email Subject",
                        value="[OUTAGE ALERT] Single Fiber Break observed in North Region"
                    )

                    with gr.Row():
                        outage_to_email_tb = gr.Textbox(
                            label="To Recipients (comma separated)",
                            placeholder="nasirnawaz918@gmail.com, team@company.com"
                        )
                        outage_cc_email_tb = gr.Textbox(
                            label="CC Recipients (comma separated)",
                            placeholder="manager@company.com"
                        )

                    outage_send_email_btn = gr.Button("🚀 Send Outage Email Report", variant="primary")
                    outage_email_status_tb = gr.Textbox(label="Dispatch Status", interactive=False)

                outage_inputs_list = [
                    outage_region_dd, outage_status_dd, outage_event_detail_dd, outage_custom_detail_tb,
                    outage_network_elements_tb, outage_ptn_imp_tb, outage_otn_imp_tb,
                    outage_start_time_mode, outage_start_date, outage_start_hour, outage_start_minute, outage_start_ampm
                ]

                outage_clock_inputs_list = [
                    outage_start_time_mode, outage_start_date, outage_start_hour, outage_start_minute, outage_start_ampm
                ]

                outage_event_detail_dd.change(
                    fn=toggle_outage_custom_box, 
                    inputs=outage_event_detail_dd, 
                    outputs=outage_custom_detail_tb
                )
                outage_start_time_mode.change(
                    fn=toggle_outage_start_mode, 
                    inputs=[outage_start_time_mode], 
                    outputs=[outage_start_date, outage_start_hour, outage_start_minute, outage_start_ampm, outage_start_set_current]
                )
                outage_start_set_current.click(
                    fn=current_clock_picker_values, 
                    inputs=[], 
                    outputs=[
                        outage_start_time_mode, outage_start_date, outage_start_hour, 
                        outage_start_minute, outage_start_ampm, outage_start_clock_preview
                    ]
                )

                for clock_comp in outage_clock_inputs_list:
                    clock_comp.change(
                        fn=generate_outage_clock_card_html, 
                        inputs=outage_clock_inputs_list, 
                        outputs=outage_start_clock_preview
                    )

                for input_comp in outage_inputs_list:
                    input_comp.change(
                        fn=generate_dynamic_outage_table, 
                        inputs=outage_inputs_list, 
                        outputs=outage_table_output
                    )

                outage_send_email_btn.click(
                    fn=send_outage_email,
                    inputs=[outage_to_email_tb, outage_cc_email_tb, outage_email_subject_tb, outage_table_output],
                    outputs=outage_email_status_tb
                )

            # TAB 4: VENDOR MATRIX MANAGER
            with gr.Tab("🏪 Vendor Matrix Manager"):
                gr.Markdown("### 📊 Vendor Escalation Matrix Management (JSON Persistence)")

                with gr.Row():
                    with gr.Column(scale=1):
                        gr.Markdown("#### 1️⃣ Select / Add Vendor")
                        vendor_select_dd = gr.Dropdown(
                            choices=get_vendor_list(),
                            value=get_vendor_list()[0] if get_vendor_list() else None,
                            label="Select Vendor"
                        )
                        new_vendor_tb = gr.Textbox(label="Or Add New Vendor Name", placeholder="e.g. Wateen / Cybernet")
                        add_vendor_btn = gr.Button("➕ Add Vendor Name")

                    with gr.Column(scale=2):
                        gr.Markdown("#### 2️⃣ Add Contact Level to Matrix")
                        with gr.Row():
                            level_dd = gr.Dropdown(
                                choices=[
                                    "Level 1", 
                                    "Level 2 (Central Region)", 
                                    "Level 2 (South Region)", 
                                    "Level 2 (North Region)", 
                                    "Level 3", 
                                    "Level 4", 
                                    "Level 5"
                                ],
                                value="Level 1",
                                label="Escalation Level"
                            )
                            c_name_tb = gr.Textbox(label="Contact Name", placeholder="e.g. Ali Nadeem")
                            c_desig_tb = gr.Textbox(label="Designation", placeholder="e.g. Regional Head")

                        with gr.Row():
                            c_time_tb = gr.Textbox(label="Escalation Time", placeholder="e.g. 1 Hour / 3 hours")
                            c_phone_tb = gr.Textbox(label="Contact Phone", placeholder="0300-XXXXXXX")
                            c_email_tb = gr.Textbox(label="Email Address", placeholder="person@vendor.com")

                        add_contact_btn = gr.Button("➕ Add Contact to Matrix", variant="primary")

                matrix_status_msg = gr.Textbox(label="Action Status", interactive=False)

                gr.Markdown("---")
                gr.Markdown("#### 📋 Copy All Emails for Outlook")
                copy_emails_tb = gr.Textbox(
                    label="All Vendor Emails (Copy-paste directly to Outlook To/CC)",
                    value=get_vendor_emails_string(get_vendor_list()[0] if get_vendor_list() else ""),
                    interactive=True
                )

                gr.Markdown("---")
                gr.Markdown("### 📄 Live Matrix Table Preview (HTML Red Header Format)")
                
                initial_vendor = get_vendor_list()[0] if get_vendor_list() else ""
                matrix_html_preview = gr.HTML(value=render_escalation_matrix_html(initial_vendor))

                gr.Markdown("---")
                with gr.Row():
                    with gr.Column():
                        gr.Markdown("#### 🗑️ Delete Contact Person")
                        del_name_tb = gr.Textbox(label="Enter Exact Name to Delete", placeholder="e.g. Ali Nadeem")
                        del_contact_btn = gr.Button("🗑️ Remove Person", variant="stop")

                    with gr.Column():
                        gr.Markdown("#### 📋 Dataframe View")
                        matrix_df_view = gr.Dataframe(
                            headers=["Level", "Name", "Designation", "Time", "Phone", "Email"],
                            value=get_matrix_dataframe(initial_vendor),
                            interactive=False
                        )

                # Event Bindings
                vendor_select_dd.change(
                    fn=lambda v: (render_escalation_matrix_html(v), get_matrix_dataframe(v), get_vendor_emails_string(v)),
                    inputs=vendor_select_dd,
                    outputs=[matrix_html_preview, matrix_df_view, copy_emails_tb]
                )

                add_vendor_btn.click(
                    fn=add_vendor_name,
                    inputs=new_vendor_tb,
                    outputs=[vendor_select_dd, copy_emails_tb, matrix_status_msg]
                )

                add_contact_btn.click(
                    fn=add_contact_to_matrix,
                    inputs=[vendor_select_dd, level_dd, c_name_tb, c_desig_tb, c_time_tb, c_phone_tb, c_email_tb],
                    outputs=[matrix_html_preview, matrix_df_view, copy_emails_tb, matrix_status_msg]
                )

                del_contact_btn.click(
                    fn=delete_contact_from_matrix,
                    inputs=[vendor_select_dd, del_name_tb],
                    outputs=[matrix_html_preview, matrix_df_view, copy_emails_tb, matrix_status_msg]
                )

            # TAB 5: OPENING
            with gr.Tab("🚨 Opening"):
                gr.Markdown("### Open / Start a Complaint")

                with gr.Row():
                    opening_complaint_selector = gr.Dropdown(
                        choices=[],
                        label="Select Complaint to Open",
                    )
                    opening_refresh_complaints = gr.Button("Refresh Opening Complaint List")

                opening_selected_details = gr.Textbox(
                    label="Selected Complaint Details",
                    lines=3,
                    interactive=False,
                )

                opening_issue = gr.Dropdown(
                    choices=OPENING_ISSUES,
                    value="Link is down.",
                    label="Issue Type",
                )
                opening_custom_issue = gr.Textbox(
                    label="Custom Issue",
                    placeholder="Used only when Other / Manual is selected",
                )
                opening_priority = gr.Dropdown(
                    choices=["Normal", "Urgent", "Critical"],
                    value="Normal",
                    label="Priority",
                )

                gr.Markdown("### 🕐 Opening Clock")
                opening_time_mode = gr.Radio(
                    choices=[AUTO_TIME_MODE, MANUAL_TIME_MODE],
                    value=AUTO_TIME_MODE,
                    label="Opening Time Mode",
                )

                with gr.Row():
                    opening_date = gr.DateTime(
                        value=None,
                        include_time=False,
                        type="string",
                        label="Opening Date",
                        interactive=True,
                    )
                    opening_hour = gr.Dropdown(
                        choices=HOUR_CHOICES,
                        value=None,
                        label="Hour",
                    )
                    opening_minute = gr.Dropdown(
                        choices=MINUTE_CHOICES,
                        value=None,
                        label="Minute",
                    )
                    opening_ampm = gr.Radio(
                        choices=AMPM_CHOICES,
                        value=None,
                        label="AM / PM",
                    )

                opening_set_current = gr.Button("🕐 Set Manual Picker to Current Pakistan Time")
                opening_clock_preview = gr.HTML(
                    value=clock_card_html(),
                    label="Opening Clock Preview",
                )

                opening_button = gr.Button("Generate Opening Email", variant="primary")
                opening_subject = gr.Textbox(label="Subject", interactive=True)
                opening_body = gr.Textbox(label="Opening Email", lines=12, interactive=True)
                opening_runtime = gr.Textbox(label="Runtime Status", interactive=False)

            # TAB 6: QUICK RESPONSE
            with gr.Tab("⚡ Quick Response"):
                with gr.Row():
                    quick_scenario = gr.Dropdown(
                        choices=QUICK_SCENARIOS,
                        value="Acknowledgement",
                        label="Scenario",
                    )
                    quick_stage = gr.Dropdown(
                        choices=QUICK_STAGES,
                        value="Initial",
                        label="Update Stage",
                    )
                    quick_team = gr.Dropdown(
                        choices=QUICK_TEAMS,
                        value="Concerned team",
                        label="Team",
                    )

                with gr.Row():
                    quick_ettr = gr.Dropdown(
                        choices=[
                            "Awaited", "30 Minutes", "1 Hour", "2 Hours",
                            "4 Hours", "Not Available", "Not Applicable", "Custom"
                        ],
                        value="Awaited",
                        label="ETTR",
                    )
                    quick_custom_ettr = gr.Textbox(
                        label="Custom ETTR",
                        placeholder="e.g. 90 Minutes",
                    )
                    quick_priority = gr.Dropdown(
                        choices=["Normal", "Follow-up", "Urgent", "Critical"],
                        value="Normal",
                        label="Priority",
                    )
                    quick_audience = gr.Dropdown(
                        choices=["Customer", "Team / Vendor"],
                        value="Customer",
                        label="Audience",
                    )

                quick_note = gr.Textbox(
                    label="Optional Custom Update",
                    placeholder="Any confirmed ground update to include...",
                    lines=2,
                )

                quick_button = gr.Button("Generate Quick Response", variant="primary")
                quick_subject = gr.Textbox(label="Subject", interactive=True)
                quick_body = gr.Textbox(label="Generated Email", lines=12, interactive=True)
                quick_policy = gr.Textbox(
                    label="Internal Policy / Warning",
                    lines=2,
                    interactive=False,
                )

            # TAB 7: CLOSURE / RFO
            with gr.Tab("✅ Closure / RFO"):
                gr.Markdown("### Close an Active Complaint")

                with gr.Row():
                    closure_complaint_selector = gr.Dropdown(
                        choices=[],
                        label="Select OPEN Complaint to Close",
                    )
                    closure_refresh_complaints = gr.Button("Refresh Open Complaint List")

                closure_selected_details = gr.Textbox(
                    label="Selected Complaint Details",
                    lines=3,
                    interactive=False,
                )

                with gr.Row():
                    issue_found_at = gr.Dropdown(
                        choices=ISSUE_FOUND_AT,
                        value="Customer",
                        label="Issue Found At",
                    )
                    custom_issue_found_at = gr.Textbox(
                        label="Custom Issue Found At",
                        placeholder="Used only for Other",
                    )

                root_cause = gr.Dropdown(
                    choices=RCA_OPTIONS,
                    value="No Issue Observed",
                    label="Root Cause",
                )
                custom_root_cause = gr.Textbox(
                    label="Custom Root Cause",
                    placeholder="Used only for Other / Manual",
                )

                corrective_action = gr.Dropdown(
                    choices=[
                        "Auto from Root Cause", "No activity performed at our end",
                        "Fiber restored", "Configuration rectified", "Link rerouted",
                        "Power restored", "Wireless parameters optimized",
                        "Optical parameters normalized", "Faulty equipment replaced",
                        "Port configuration corrected", "Issue resolved by upstream provider",
                        "Customer end issue rectified", "Link stabilized after troubleshooting",
                        "Other / Manual",
                    ],
                    value="Auto from Root Cause",
                    label="Corrective Action",
                )
                custom_corrective_action = gr.Textbox(
                    label="Custom Corrective Action",
                    placeholder="Used only for Other / Manual",
                )

                final_status = gr.Dropdown(
                    choices=[
                        "Link is up and stable.", "Service is up and working normally.",
                        "Wireless segment is up and normal.", "Link parameters are normal.",
                        "Service has been restored successfully.", "No abnormality is currently observed at our end.",
                    ],
                    value="Service is up and working normally.",
                    label="Final Status",
                )
                closure_priority = gr.Dropdown(
                    choices=["Normal", "Urgent"],
                    value="Normal",
                    label="Priority",
                )

                gr.Markdown("### 🕐 Closing / Restoration Clock")
                closure_time_mode = gr.Radio(
                    choices=[AUTO_TIME_MODE, MANUAL_TIME_MODE],
                    value=AUTO_TIME_MODE,
                    label="Closing Time Mode",
                )

                with gr.Row():
                    closure_date = gr.DateTime(
                        value=None,
                        include_time=False,
                        type="string",
                        label="Restoration Date",
                        interactive=True,
                    )
                    closure_hour = gr.Dropdown(
                        choices=HOUR_CHOICES,
                        value=None,
                        label="Hour",
                    )
                    closure_minute = gr.Dropdown(
                        choices=MINUTE_CHOICES,
                        value=None,
                        label="Minute",
                    )
                    closure_ampm = gr.Radio(
                        choices=AMPM_CHOICES,
                        value=None,
                        label="AM / PM",
                    )

                closure_set_current = gr.Button("🕐 Set Manual Picker to Current Pakistan Time")
                closure_clock_preview = gr.HTML(
                    value=clock_card_html(),
                    label="Closing Clock Preview",
                )

                closure_button = gr.Button("Generate Closure Email", variant="primary")
                closure_subject = gr.Textbox(label="Subject", interactive=True)
                closure_body = gr.Textbox(label="Generated Closure Email", lines=14, interactive=True)
                closure_runtime = gr.Textbox(label="Runtime Status", interactive=False)

            # TAB 8: TRANSMISSION / MASS OUTAGE
            with gr.Tab("📡 Transmission / Mass Outage"):
                affected_links = gr.Textbox(
                    label="Affected Services (Optional - one per line)",
                    placeholder="Paste multiple labels here during mass outage...",
                    lines=5,
                )

                with gr.Row():
                    transmission_fault = gr.Dropdown(
                        choices=TRANSMISSION_FAULTS,
                        value="Single Fiber break",
                        label="Fault Type",
                    )
                    transmission_action = gr.Dropdown(
                        choices=TRANSMISSION_ACTIONS,
                        value="Concerned team engaged",
                        label="Current Action",
                    )
                    transmission_rerouting = gr.Dropdown(
                        choices=REROUTING_OPTIONS,
                        value="Not Applicable",
                        label="Rerouting",
                    )

                with gr.Row():
                    transmission_location = gr.Textbox(
                        label="Fault Location",
                        placeholder="Optional",
                    )
                    transmission_ettr = gr.Dropdown(
                        choices=[
                            "Awaited", "30 Minutes", "1 Hour", "2 Hours",
                            "4 Hours", "Not Available", "Not Applicable", "Custom"
                        ],
                        value="Awaited",
                        label="ETTR",
                    )
                    transmission_custom_ettr = gr.Textbox(
                        label="Custom ETTR",
                        placeholder="Optional",
                    )

                with gr.Row():
                    transmission_priority = gr.Dropdown(
                        choices=["Normal", "Urgent", "Critical"],
                        value="Urgent",
                        label="Priority",
                    )
                    transmission_audience = gr.Dropdown(
                        choices=["Customer", "Team / Vendor"],
                        value="Customer",
                        label="Audience",
                    )

                transmission_button = gr.Button("Generate Transmission Update", variant="primary")
                transmission_subject = gr.Textbox(label="Subject", interactive=True)
                transmission_body = gr.Textbox(label="Generated Email", lines=12, interactive=True)
                transmission_note = gr.Textbox(label="Internal Note", lines=2, interactive=False)

            # TAB 9: CUSTOMER END
            with gr.Tab("👤 Customer End / Findings"):
                customer_issue_summary = gr.Textbox(
                    label="Issue Summary",
                    value="Reported service issue",
                )
                customer_findings = gr.CheckboxGroup(
                    choices=FINDING_OPTIONS,
                    label="Select Findings",
                )
                customer_action = gr.Dropdown(
                    choices=CUSTOMER_ACTIONS,
                    value="Verify last-mile media",
                    label="Requested Customer Action",
                )
                customer_priority = gr.Dropdown(
                    choices=["Normal", "Follow-up", "Urgent"],
                    value="Normal",
                    label="Priority",
                )

                customer_end_button = gr.Button("Generate Customer-End Response", variant="primary")
                customer_end_subject = gr.Textbox(label="Subject", interactive=True)
                customer_end_body = gr.Textbox(label="Generated Email", lines=12, interactive=True)
                customer_policy = gr.Textbox(label="Internal Policy / Warning", lines=2, interactive=False)

            # TAB 10: TROUBLESHOOTING / STATS
            with gr.Tab("🧪 Stats / Troubleshooting"):
                stats_scenario = gr.Dropdown(
                    choices=STATS_SCENARIOS,
                    value="Packet Loss / High Latency",
                    label="Reported Scenario",
                )
                stats_audience = gr.Dropdown(
                    choices=["Customer", "Team / Vendor"],
                    value="Customer",
                    label="Audience",
                )
                stats_context = gr.Textbox(
                    label="Optional Additional Context",
                    placeholder="e.g. Please perform testing from the same affected IP pool.",
                    lines=2,
                )
                stats_button = gr.Button("Generate Required-Stats Email", variant="primary")
                stats_subject = gr.Textbox(label="Subject", interactive=True)
                stats_body = gr.Textbox(label="Generated Email", lines=14, interactive=True)
                stats_policy = gr.Textbox(label="Internal Policy / Warning", lines=2, interactive=False)

            # TAB 11: CUSTOMER FOLLOW-UP
            with gr.Tab("🔔 Customer Follow-up"):
                customer_followup_type = gr.Dropdown(
                    choices=CUSTOMER_FOLLOWUPS,
                    value="Awaiting requested stats",
                    label="Follow-up Type",
                )
                customer_pending = gr.Textbox(
                    label="Pending / Requested Information",
                    placeholder="Optional: paste the specific pending request...",
                    lines=3,
                )
                customer_followup_priority = gr.Dropdown(
                    choices=["Normal", "Follow-up", "Urgent"],
                    value="Follow-up",
                    label="Priority",
                )
                customer_followup_button = gr.Button("Generate Customer Follow-up", variant="primary")
                customer_followup_subject = gr.Textbox(label="Subject", interactive=True)
                customer_followup_body = gr.Textbox(label="Generated Email", lines=10, interactive=True)

            # TAB 12: PROGRESS UPDATE
            with gr.Tab("🔄 Progress Update"):
                progress_status = gr.Dropdown(
                    choices=PROGRESS_OPTIONS,
                    value="Concerned team engaged",
                    label="Current Progress",
                )
                with gr.Row():
                    progress_ettr = gr.Dropdown(
                        choices=[
                            "Awaited", "30 Minutes", "1 Hour", "2 Hours",
                            "4 Hours", "Not Available", "Not Applicable", "Custom"
                        ],
                        value="Awaited",
                        label="ETTR",
                    )
                    progress_custom_ettr = gr.Textbox(label="Custom ETTR", placeholder="Optional")
                    progress_priority = gr.Dropdown(
                        choices=["Normal", "Follow-up", "Urgent", "Critical"],
                        value="Normal",
                        label="Priority",
                    )
                    progress_audience = gr.Dropdown(
                        choices=["Customer", "Team / Vendor"],
                        value="Customer",
                        label="Audience",
                    )

                progress_note = gr.Textbox(label="Additional Ground Update", placeholder="Optional", lines=2)
                progress_button = gr.Button("Generate Progress Update", variant="primary")
                progress_subject = gr.Textbox(label="Subject", interactive=True)
                progress_body = gr.Textbox(label="Generated Email", lines=12, interactive=True)

            # TAB 13: VENDOR / INTERNAL ESCALATION
            with gr.Tab("📨 Vendor / Internal Escalation"):
                with gr.Row():
                    escalation_target = gr.Dropdown(
                        choices=ESCALATION_TARGETS,
                        value="Netsat",
                        label="Escalate To",
                    )
                    custom_escalation_target = gr.Textbox(
                        label="Custom Target Name",
                        placeholder="Used only when 'Other' is selected",
                    )

                escalation_level = gr.Dropdown(
                    choices=ESCALATION_LEVELS,
                    value="Initial engagement",
                    label="Escalation Level",
                )

                escalation_findings = gr.CheckboxGroup(
                    choices=FINDING_OPTIONS,
                    label="Current Findings (Optional)",
                )

                escalation_action = gr.Textbox(
                    label="Required Action (Optional)",
                    placeholder="e.g. Kindly dispatch a team to the site.",
                    lines=2,
                )

                escalation_priority = gr.Dropdown(
                    choices=["Normal", "Follow-up", "Urgent", "Critical"],
                    value="Urgent",
                    label="Priority",
                )

                escalation_button = gr.Button("Generate Escalation Email", variant="primary")
                escalation_subject = gr.Textbox(label="Subject", interactive=True)
                escalation_body = gr.Textbox(label="Generated Email", lines=12, interactive=True)
                matrix_table_output = gr.HTML(label="Vendor Escalation Matrix")

            # TAB 14: FIELD VISIT
            with gr.Tab("🚗 Field Visit / Access"):
                field_status = gr.Dropdown(
                    choices=FIELD_STATUSES,
                    value="Team dispatched",
                    label="Visit / Field Status",
                )
                with gr.Row():
                    field_eta = gr.Textbox(label="ETA / Visit Time", placeholder="e.g. 10:30 AM")
                    field_poc = gr.Textbox(label="POC", placeholder="Optional")
                field_access = gr.Textbox(label="Access / Coordination Details", placeholder="Optional", lines=2)
                with gr.Row():
                    field_audience = gr.Dropdown(
                        choices=["Customer", "Team / Vendor"],
                        value="Customer",
                        label="Audience",
                    )
                    field_priority = gr.Dropdown(
                        choices=["Normal", "Urgent", "Critical"],
                        value="Normal",
                        label="Priority",
                    )
                field_button = gr.Button("Generate Field / Access Email", variant="primary")
                field_subject = gr.Textbox(label="Subject", interactive=True)
                field_body = gr.Textbox(label="Generated Email", lines=10, interactive=True)

            # TAB 15: VPBX / VOICE / PRI / SIP
            with gr.Tab("☎️ VPBX / Voice / PRI / SIP"):
                voice_issue = gr.Dropdown(
                    choices=VOICE_ISSUES,
                    value="Incoming (Call Landing)",
                    label="VPBX / Voice Issue",
                )

                with gr.Row():
                    party_a = gr.Textbox(label="Party-A")
                    party_b = gr.Textbox(label="Party-B")
                    voice_call_time = gr.Textbox(label="Call Date / Time", placeholder="Optional")

                with gr.Row():
                    voice_call_direction = gr.Dropdown(
                        choices=["Not Applicable", "Incoming", "Outgoing", "Both"],
                        value="Not Applicable",
                        label="Call Direction",
                    )
                    voice_extension = gr.Textbox(label="Affected Extension / Master Number", placeholder="Optional")

                voice_behavior = gr.Textbox(
                    label="Observed Behavior / Customer Request",
                    placeholder="e.g. call not landing / distortion / random CLI / IVR flow required...",
                    lines=2,
                )

                with gr.Row():
                    voice_response_to = gr.Dropdown(
                        choices=VPBX_TARGETS,
                        value="Customer",
                        label="Generate Response For",
                    )
                    voice_priority = gr.Dropdown(
                        choices=["Normal", "Urgent", "Critical"],
                        value="Normal",
                        label="Priority",
                    )

                voice_button = gr.Button("Generate VPBX / Voice Email", variant="primary")
                voice_subject = gr.Textbox(label="Subject", interactive=True)
                voice_body = gr.Textbox(label="Generated Email", lines=12, interactive=True)

            # TAB 16: SMART PASTE
            with gr.Tab("🧠 Smart Paste"):
                raw_complaint = gr.Textbox(
                    label="Paste Customer / Internal Email",
                    placeholder="Paste the received complaint here...",
                    lines=10,
                )
                parse_button = gr.Button("Detect Details", variant="primary")

                with gr.Row():
                    detected_vlan = gr.Textbox(label="Detected VLAN")
                    detected_ticket = gr.Textbox(label="Detected Ticket")
                detected_summary = gr.Textbox(
                    label="Detection Summary",
                    lines=4,
                    interactive=False,
                )

            # TAB 17: ACTIVE COMPLAINTS DASHBOARD
            with gr.Tab("📋 Active Complaints / Response Age"):
                with gr.Row():
                    dashboard_view = gr.Dropdown(
                        choices=[
                            "Queued + Open", "Open Only", "Queued Only",
                            "Closed Only", "All Runtime Complaints",
                        ],
                        value="Queued + Open",
                        label="View",
                    )
                    dashboard_complaint_selector = gr.Dropdown(
                        choices=[],
                        label="Select Complaint from Runtime",
                    )

                with gr.Row():
                    dashboard_refresh = gr.Button("Refresh Dashboard", variant="primary")
                    dashboard_selector_refresh = gr.Button("Refresh Complaint Selector")
                    dashboard_load_button = gr.Button("Load Selected Complaint")

                dashboard = gr.Dataframe(
                    headers=[
                        "Service", "Type", "Issue", "Added", "Opened",
                        "Age", "Closed", "Total Time", "Ticket", "Last Stage", "Status"
                    ],
                    datatype=["str"] * 11,
                    value=[],
                    interactive=False,
                    label="Runtime Complaints",
                )

                check_button = gr.Button("Check Current Label")
                incident_status = gr.Textbox(
                    label="Current Label Runtime Details",
                    lines=8,
                    interactive=False,
                )
                reset_button = gr.Button("Clear Current Label from Runtime")

        # ====================================================
        # GLOBAL EVENT BINDINGS
        # ====================================================
        add_complaints_event = add_complaints_button.click(
            fn=register_complaints,
            inputs=[bulk_complaints, complaint_default_issue, complaint_bulk_ticket],
            outputs=[complaint_selector, complaint_manager_status, complaint_manager_dashboard],
        )

        add_complaints_event.then(
            fn=refresh_opening_complaint_selector,
            inputs=[],
            outputs=opening_complaint_selector,
        ).then(
            fn=refresh_closure_complaint_selector,
            inputs=[],
            outputs=closure_complaint_selector,
        ).then(
            fn=refresh_complaint_selector,
            inputs=[],
            outputs=dashboard_complaint_selector,
        )

        refresh_complaints_button.click(
            fn=refresh_complaint_selector,
            inputs=[],
            outputs=complaint_selector,
        )

        load_complaint_button.click(
            fn=load_selected_complaint,
            inputs=complaint_selector,
            outputs=[label, common_ticket, complaint_manager_status],
        )

        opening_refresh_complaints.click(
            fn=refresh_opening_complaint_selector,
            inputs=[],
            outputs=opening_complaint_selector,
        )

        opening_complaint_selector.change(
            fn=load_selected_complaint,
            inputs=opening_complaint_selector,
            outputs=[label, common_ticket, opening_selected_details],
        )

        closure_refresh_complaints.click(
            fn=refresh_closure_complaint_selector,
            inputs=[],
            outputs=closure_complaint_selector,
        )

        closure_complaint_selector.change(
            fn=load_selected_complaint,
            inputs=closure_complaint_selector,
            outputs=[label, common_ticket, closure_selected_details],
        )

        remove_complaint_button.click(
            fn=remove_selected_complaint,
            inputs=complaint_selector,
            outputs=[complaint_selector, complaint_manager_status, complaint_manager_dashboard],
        )

        parse_button.click(
            fn=parse_complaint,
            inputs=raw_complaint,
            outputs=[
                label, detected_vlan, quick_scenario,
                detected_ticket, common_ticket, detected_summary,
            ],
        )

        opening_set_current.click(
            fn=current_clock_picker_values,
            inputs=[],
            outputs=[
                opening_time_mode, opening_date, opening_hour,
                opening_minute, opening_ampm, opening_clock_preview,
            ],
        )

        for _comp in [opening_date, opening_hour, opening_minute, opening_ampm]:
            _comp.change(
                fn=update_clock_preview,
                inputs=[opening_date, opening_hour, opening_minute, opening_ampm],
                outputs=opening_clock_preview,
            )

        closure_set_current.click(
            fn=current_clock_picker_values,
            inputs=[],
            outputs=[
                closure_time_mode, closure_date, closure_hour,
                closure_minute, closure_ampm, closure_clock_preview,
            ],
        )

        for _comp in [closure_date, closure_hour, closure_minute, closure_ampm]:
            _comp.change(
                fn=update_clock_preview,
                inputs=[closure_date, closure_hour, closure_minute, closure_ampm],
                outputs=closure_clock_preview,
            )

        opening_event = opening_button.click(
            fn=generate_opening_email,
            inputs=[
                opening_complaint_selector, label, opening_issue, opening_custom_issue,
                common_ticket, opening_priority, opening_time_mode, opening_date,
                opening_hour, opening_minute, opening_ampm,
            ],
            outputs=[label, opening_subject, opening_body, opening_runtime],
        )

        opening_event.then(
            fn=refresh_opening_complaint_selector,
            inputs=[],
            outputs=opening_complaint_selector,
        ).then(
            fn=refresh_closure_complaint_selector,
            inputs=[],
            outputs=closure_complaint_selector,
        )

        quick_button.click(
            fn=generate_quick_response,
            inputs=[
                label, quick_scenario, quick_stage, quick_team,
                quick_ettr, quick_custom_ettr, quick_priority,
                quick_audience, common_ticket, quick_note,
            ],
            outputs=[quick_subject, quick_body, quick_policy],
        )

        transmission_button.click(
            fn=generate_transmission_email,
            inputs=[
                label, affected_links, transmission_fault, transmission_action,
                transmission_rerouting, transmission_location, transmission_ettr,
                transmission_custom_ettr, transmission_priority, transmission_audience,
            ],
            outputs=[transmission_subject, transmission_body, transmission_note],
        )

        customer_end_button.click(
            fn=generate_customer_end_email,
            inputs=[
                label, customer_issue_summary, customer_findings,
                customer_action, customer_priority, common_ticket,
            ],
            outputs=[customer_end_subject, customer_end_body, customer_policy],
        )

        stats_button.click(
            fn=generate_stats_request,
            inputs=[label, stats_scenario, stats_audience, stats_context],
            outputs=[stats_subject, stats_body, stats_policy],
        )

        customer_followup_button.click(
            fn=generate_customer_followup,
            inputs=[label, customer_followup_type, customer_pending, customer_followup_priority],
            outputs=[customer_followup_subject, customer_followup_body],
        )

        progress_button.click(
            fn=generate_progress_email,
            inputs=[
                label, progress_status, progress_ettr, progress_custom_ettr,
                progress_note, progress_priority, progress_audience,
            ],
            outputs=[progress_subject, progress_body],
        )

        escalation_button.click(
            fn=generate_escalation_email,
            inputs=[
                label, escalation_target, escalation_level, escalation_findings,
                escalation_action, escalation_priority, custom_escalation_target,
            ],
            outputs=[escalation_subject, escalation_body, matrix_table_output],
        )

        field_button.click(
            fn=generate_field_visit_email,
            inputs=[
                label, field_status, field_eta, field_poc,
                field_access, field_audience, field_priority,
            ],
            outputs=[field_subject, field_body],
        )

        voice_button.click(
            fn=generate_voice_email,
            inputs=[
                label, voice_issue, party_a, party_b, voice_call_time,
                voice_call_direction, voice_extension, voice_behavior,
                voice_response_to, voice_priority,
            ],
            outputs=[voice_subject, voice_body],
        )

        closure_event = closure_button.click(
            fn=generate_closure_email,
            inputs=[
                closure_complaint_selector, label, issue_found_at, custom_issue_found_at,
                root_cause, custom_root_cause, corrective_action, custom_corrective_action,
                final_status, closure_priority, closure_time_mode, closure_date,
                closure_hour, closure_minute, closure_ampm,
            ],
            outputs=[label, closure_subject, closure_body, closure_runtime],
        )

        closure_event.then(
            fn=refresh_closure_complaint_selector,
            inputs=[],
            outputs=closure_complaint_selector,
        ).then(
            fn=refresh_opening_complaint_selector,
            inputs=[],
            outputs=opening_complaint_selector,
        ).then(
            fn=refresh_complaint_selector,
            inputs=[],
            outputs=dashboard_complaint_selector,
        )

        dashboard_refresh.click(
            fn=refresh_dashboard,
            inputs=dashboard_view,
            outputs=dashboard,
        )

        dashboard_selector_refresh.click(
            fn=refresh_complaint_selector,
            inputs=[],
            outputs=dashboard_complaint_selector,
        )

        dashboard_load_button.click(
            fn=load_selected_complaint,
            inputs=dashboard_complaint_selector,
            outputs=[label, common_ticket, incident_status],
        )

        check_button.click(
            fn=check_incident,
            inputs=label,
            outputs=incident_status,
        )

        reset_button.click(
            fn=reset_incident,
            inputs=label,
            outputs=incident_status,
        )

    return app


# ============================================================
# 22. APPLICATION LAUNCH
# ============================================================
app = build_app()
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 7860))
    app.launch(
        server_name="0.0.0.0",
        server_port=port,
        auth=[
            ("nasir", "123"),
            ("ijaz", "123"),
            ("inam", "123"),
            ("mazhar", "123"),
        ],
        auth_message="🔒 Corporate NOC Response Console - Authorized Personnel Only"
    )
