#!/usr/bin/env python3
"""
Irish Rail Spend Tracker — Streamlit UI
"""

import io
import os
import re
import time
import pickle
import threading
from datetime import datetime, timedelta
from collections import defaultdict

import streamlit as st

# ── Page config (must be first Streamlit call) ────────────────────────────────
st.set_page_config(
    page_title="Irish Rail Spend Tracker",
    page_icon="🚆",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── Lazy imports (heavy deps only loaded when needed) ─────────────────────────
def _gmail_imports():
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build
    import base64
    return Request, Credentials, InstalledAppFlow, build, base64

# ── Pull in the core logic from main.py ───────────────────────────────────────
from main import (
    get_gmail_service,
    search_messages,
    get_message_body,
    parse_booking_email,
    extract_booking_ref,
    norm_name,
    write_excel,
    week_start,
    week_label,
    IRISH_RAIL_SENDER,
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
HERO_IMAGE = "https://images.unsplash.com/photo-1679410153289-2c4e3fa60a3a?w=1600&q=80&auto=format&fit=crop"
# Intercity train at a station — free to use under Unsplash licence

st.markdown(f"""
<style>
  /* ── Google fonts ── */
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

  html, body, [class*="css"] {{
    font-family: 'Inter', sans-serif;
  }}

  /* ── Hide default Streamlit chrome ── */
  #MainMenu, footer, header {{visibility: hidden;}}
  .block-container {{padding-top: 0 !important; max-width: 960px;}}

  /* ── Hero ── */
  .hero {{
    position: relative;
    width: 100%;
    height: 340px;
    background: url('{HERO_IMAGE}') center/cover no-repeat;
    border-radius: 0 0 16px 16px;
    overflow: hidden;
    margin-bottom: 2rem;
  }}
  .hero-overlay {{
    position: absolute; inset: 0;
    background: linear-gradient(135deg, rgba(0,60,30,0.82) 0%, rgba(0,30,60,0.70) 100%);
    display: flex; flex-direction: column;
    align-items: flex-start; justify-content: flex-end;
    padding: 2.5rem 3rem;
  }}
  .hero-tag {{
    background: #00A651; color: #fff; font-size: 0.72rem;
    font-weight: 700; letter-spacing: 0.12em; text-transform: uppercase;
    padding: 4px 12px; border-radius: 20px; margin-bottom: 0.8rem;
  }}
  .hero-title {{
    color: #fff; font-size: 2.6rem; font-weight: 700;
    line-height: 1.1; margin: 0 0 0.5rem;
  }}
  .hero-sub {{
    color: rgba(255,255,255,0.80); font-size: 1.05rem; margin: 0;
  }}

  /* ── Card ── */
  .card {{
    background: #fff; border: 1px solid #e5e7eb;
    border-radius: 12px; padding: 1.6rem 2rem;
    margin-bottom: 1.2rem;
  }}
  .card-title {{
    font-size: 1rem; font-weight: 700; color: #111;
    margin: 0 0 1rem; display: flex; align-items: center; gap: 8px;
  }}

  /* ── Step badges ── */
  .step-badge {{
    display: inline-flex; align-items: center; justify-content: center;
    width: 26px; height: 26px; border-radius: 50%;
    background: #00A651; color: #fff; font-size: 0.75rem; font-weight: 700;
  }}

  /* ── Metric row ── */
  .metric-row {{
    display: flex; gap: 1rem; flex-wrap: wrap; margin-bottom: 1.4rem;
  }}
  .metric-box {{
    flex: 1; min-width: 140px;
    background: #f0faf4; border: 1px solid #bbf0d0;
    border-radius: 10px; padding: 1rem 1.2rem;
    text-align: center;
  }}
  .metric-value {{
    font-size: 1.8rem; font-weight: 700; color: #065f46; line-height: 1;
  }}
  .metric-label {{
    font-size: 0.78rem; color: #6b7280; margin-top: 4px; font-weight: 500;
  }}

  /* ── Journey table ── */
  .journey-table {{
    width: 100%; border-collapse: collapse; font-size: 0.88rem;
  }}
  .journey-table th {{
    background: #003d1a; color: #fff; padding: 8px 12px;
    text-align: left; font-weight: 600; font-size: 0.78rem;
    letter-spacing: 0.04em; text-transform: uppercase;
  }}
  .journey-table th:first-child {{ border-radius: 8px 0 0 0; }}
  .journey-table th:last-child  {{ border-radius: 0 8px 0 0; }}
  .journey-table td {{
    padding: 9px 12px; border-bottom: 1px solid #f0f0f0; vertical-align: middle;
  }}
  .journey-table tr:last-child td {{ border-bottom: none; }}
  .journey-table tr:hover td {{ background: #f6fff9; }}
  .route-pill {{
    background: #e8f5ee; color: #065f46; padding: 3px 10px;
    border-radius: 20px; font-size: 0.82rem; font-weight: 500;
    white-space: nowrap;
  }}
  .no-travel {{ color: #9ca3af; font-style: italic; font-size: 0.82rem; }}
  .amount-chip {{
    font-weight: 600; color: #065f46;
  }}
  .ref-chip {{
    font-size: 0.78rem; color: #9ca3af; font-family: monospace;
  }}

  /* ── Log box ── */
  .log-box {{
    background: #0d1117; color: #58d68d; font-family: monospace;
    font-size: 0.8rem; border-radius: 8px; padding: 1rem 1.2rem;
    height: 220px; overflow-y: auto; white-space: pre-wrap;
    border: 1px solid #1e3a2e;
  }}

  /* ── Irish Rail green button override ── */
  div.stButton > button {{
    background: #00A651 !important; color: #fff !important;
    border: none !important; border-radius: 8px !important;
    font-weight: 600 !important; padding: 0.55rem 1.4rem !important;
    font-size: 0.95rem !important;
  }}
  div.stButton > button:hover {{
    background: #008c44 !important;
  }}

  /* ── Download button ── */
  div.stDownloadButton > button {{
    background: #003d1a !important; color: #fff !important;
    border: none !important; border-radius: 8px !important;
    font-weight: 600 !important;
  }}

  /* ── Rail track divider ── */
  .track {{
    display: flex; align-items: center; gap: 0; margin: 1.5rem 0;
  }}
  .track-line {{
    flex: 1; height: 3px; background: #003d1a;
  }}
  .track-tie {{
    width: 18px; height: 10px; background: #003d1a;
    margin: 0 4px; border-radius: 2px;
  }}
  .sleeper-row {{
    display: flex; align-items: center; gap: 6px;
  }}

  /* Badges */
  .badge-yes {{
    background: #d1fae5; color: #065f46; border-radius: 20px;
    padding: 2px 10px; font-size: 0.78rem; font-weight: 600;
  }}
  .badge-no {{
    background: #f3f4f6; color: #9ca3af; border-radius: 20px;
    padding: 2px 10px; font-size: 0.78rem;
  }}

  /* Upload zone */
  .upload-hint {{
    font-size: 0.82rem; color: #6b7280; margin-top: -0.5rem; margin-bottom: 0.5rem;
  }}
</style>
""", unsafe_allow_html=True)


# ── Hero banner ───────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero">
  <div class="hero-overlay">
    <div class="hero-tag">🚆 Spend Tracker</div>
    <div class="hero-title">Irish Rail<br>Travel Report</div>
    <div class="hero-sub">Scan your Gmail, see every journey, know what you spent.</div>
  </div>
</div>
""", unsafe_allow_html=True)


# ── Helpers ───────────────────────────────────────────────────────────────────
def rail_divider():
    ties = "".join(['<div class="track-tie"></div>'] * 18)
    st.markdown(f'<div class="track">{ties}</div>', unsafe_allow_html=True)


def run_scan(passenger_name, months, log_lines, result_holder, creds_bytes):
    """Run the full Gmail scan in a background thread, appending to log_lines."""
    try:
        # Write credentials to a temp file so get_gmail_service can find it
        tmp_creds = "credentials_tmp.json"
        with open(tmp_creds, "wb") as f:
            f.write(creds_bytes)

        # Patch creds path for this call
        import main as _main
        orig_path = _main.__file__

        log_lines.append("🔐 Authenticating with Gmail...")
        service = get_gmail_service()
        log_lines.append("✅ Gmail connected.\n")

        end_date = datetime.now()
        start_date = end_date - timedelta(days=months * 30)
        after_str = start_date.strftime("%Y/%m/%d")
        target_norm = norm_name(passenger_name)

        log_lines.append(f"🔍 Searching from {start_date.strftime('%d %b %Y')} to today")
        log_lines.append(f"👤 Passenger: {passenger_name}\n")

        # Step 1 — cancellations
        log_lines.append("Step 1/2 — Finding cancelled bookings...")
        cancelled_refs = set()
        cancel_query = f"from:{IRISH_RAIL_SENDER} subject:cancelled after:{after_str}"
        cancel_msgs = search_messages(service, cancel_query)
        log_lines.append(f"   Found {len(cancel_msgs)} cancellation email(s).")

        for msg in cancel_msgs:
            try:
                body, _ = get_message_body(service, msg["id"])
                if not re.search(r"your booking has been cancelled", body, re.IGNORECASE):
                    continue
                ref = extract_booking_ref(body)
                if ref:
                    cancelled_refs.add(ref)
                    log_lines.append(f"   🚫 Cancelled ref: {ref}")
            except Exception:
                pass

        log_lines.append(f"   {len(cancelled_refs)} cancelled booking(s) will be excluded.\n")

        # Step 2 — booking confirmations
        log_lines.append("Step 2/2 — Fetching booking confirmation emails...")
        booking_query = (
            f"from:{IRISH_RAIL_SENDER} "
            f"subject:\"Thank you for booking\" "
            f"after:{after_str}"
        )
        booking_msgs = search_messages(service, booking_query)
        log_lines.append(f"   Found {len(booking_msgs)} booking email(s). Checking each...\n")

        bookings = []
        excluded_cancelled = 0
        no_passenger = 0

        for i, msg in enumerate(booking_msgs):
            log_lines.append(f"   [{i+1}/{len(booking_msgs)}] Reading email...")
            try:
                body, email_date = get_message_body(service, msg["id"])
                booking = parse_booking_email(body, email_date, msg["id"], target_norm)
                if not booking:
                    no_passenger += 1
                    continue
                if booking["booking_ref"] and booking["booking_ref"] in cancelled_refs:
                    excluded_cancelled += 1
                    log_lines.append(f"   🚫 [{booking['booking_ref']}] Cancelled — excluded")
                    continue
                bookings.append(booking)
                route = (
                    f"{booking['origin']} → {booking['destination']}"
                    if booking["origin"] else "route unknown"
                )
                log_lines.append(
                    f"   ✅ [{booking['booking_ref']}] "
                    f"{booking['travel_date_str']} | {route} | {booking['amount_display']}"
                )
            except Exception as e:
                log_lines.append(f"   ⚠️  Error: {e}")

        log_lines.append(f"\n✅ Done! {len(bookings)} journey(s) found.")
        log_lines.append(f"🚫 {excluded_cancelled} cancelled booking(s) excluded.")
        log_lines.append(f"⏭  {no_passenger} email(s) without {passenger_name}.\n")

        if not bookings:
            result_holder["error"] = "No bookings found. Check the passenger name and try again."
            return

        # Build week grid
        bookings.sort(key=lambda b: b["travel_date"])
        by_week = defaultdict(list)
        for b in bookings:
            by_week[week_start(b["travel_date"])].append(b)

        weeks = []
        current = week_start(start_date)
        today_week = week_start(datetime.now())
        while current <= today_week:
            weeks.append({"monday": current, "bookings": by_week.get(current, [])})
            current += timedelta(weeks=1)

        max_journeys = max((len(w["bookings"]) for w in weeks), default=1)

        # Write Excel to bytes
        log_lines.append("📊 Generating Excel report...")
        buf = io.BytesIO()
        # write_excel expects a file path — write to a temp file then read back
        tmp_xlsx = "irish_rail_tmp.xlsx"
        write_excel(weeks, max_journeys, passenger_name, tmp_xlsx)
        with open(tmp_xlsx, "rb") as f:
            buf.write(f.read())
        os.remove(tmp_xlsx)
        buf.seek(0)

        result_holder["weeks"] = weeks
        result_holder["bookings"] = bookings
        result_holder["excel"] = buf.getvalue()
        result_holder["passenger_name"] = passenger_name

        total_spend = sum(b["amount_raw"] for b in bookings)
        weeks_travelled = sum(1 for w in weeks if w["bookings"])
        log_lines.append(f"📅 Weeks analysed: {len(weeks)}")
        log_lines.append(f"🚆 Weeks with travel: {weeks_travelled}")
        log_lines.append(f"💶 Total spend: €{total_spend:.2f}")
        log_lines.append("✅ Report ready!")

        # Cleanup temp creds
        if os.path.exists(tmp_creds):
            os.remove(tmp_creds)

    except Exception as e:
        log_lines.append(f"\n❌ Error: {e}")
        result_holder["error"] = str(e)


# ── State init ────────────────────────────────────────────────────────────────
if "scan_done" not in st.session_state:
    st.session_state.scan_done = False
if "result" not in st.session_state:
    st.session_state.result = {}


# ── Step 1 — Credentials ──────────────────────────────────────────────────────
st.markdown("""
<div class="card">
  <div class="card-title">
    <span class="step-badge">1</span> Gmail Credentials
  </div>
</div>
""", unsafe_allow_html=True)

st.markdown('<p class="upload-hint">Upload your <code>credentials.json</code> from Google Cloud Console. It stays on your machine — nothing is sent anywhere.</p>', unsafe_allow_html=True)
creds_file = st.file_uploader("credentials.json", type="json", label_visibility="collapsed")

has_token = os.path.exists("token.pickle")
if has_token:
    st.success("✅ Gmail already authorised (token.pickle found) — no re-login needed.", icon="🔑")
elif creds_file:
    st.info("credentials.json loaded. You'll be redirected to Google to authorise on first run.", icon="🔐")

rail_divider()


# ── Step 2 — Options ──────────────────────────────────────────────────────────
st.markdown("""
<div class="card">
  <div class="card-title">
    <span class="step-badge">2</span> Search Options
  </div>
</div>
""", unsafe_allow_html=True)

col1, col2 = st.columns([3, 1])
with col1:
    passenger_name = st.text_input(
        "Passenger name",
        placeholder="e.g. Séamus O'Brien",
        help="Exactly as it appears on your Irish Rail ticket.",
    )
with col2:
    months = st.number_input("Months to look back", min_value=1, max_value=60, value=13)

rail_divider()


# ── Step 3 — Run ──────────────────────────────────────────────────────────────
st.markdown("""
<div class="card">
  <div class="card-title">
    <span class="step-badge">3</span> Run the Scan
  </div>
</div>
""", unsafe_allow_html=True)

run_col, _ = st.columns([2, 4])
with run_col:
    run_btn = st.button("🚆 Scan Gmail now", use_container_width=True)

if run_btn:
    if not passenger_name.strip():
        st.error("Please enter a passenger name.")
    elif not creds_file and not has_token:
        st.error("Please upload credentials.json first.")
    else:
        creds_bytes = creds_file.read() if creds_file else None

        # If credentials.json already on disk (not uploaded), read it
        if creds_bytes is None and os.path.exists("credentials.json"):
            with open("credentials.json", "rb") as f:
                creds_bytes = f.read()
        elif creds_bytes:
            # Persist to disk for get_gmail_service
            with open("credentials.json", "wb") as f:
                f.write(creds_bytes)

        log_lines = []
        result_holder = {}

        log_area = st.empty()
        progress_bar = st.progress(0, text="Starting scan…")

        # Run synchronously (Streamlit Cloud doesn't support background threads easily)
        def update_ui():
            for pct in range(0, 95, 3):
                time.sleep(0.08)

        with st.spinner("Scanning your Gmail inbox… this may take a minute or two."):
            run_scan(passenger_name.strip(), months, log_lines, result_holder, creds_bytes or b"")
            progress_bar.progress(100, text="Scan complete!")

        # Show log
        log_text = "\n".join(log_lines)
        st.markdown(f'<div class="log-box">{log_text}</div>', unsafe_allow_html=True)

        if "error" in result_holder:
            st.error(result_holder["error"])
        else:
            st.session_state.scan_done = True
            st.session_state.result = result_holder
            st.rerun()


# ── Results ───────────────────────────────────────────────────────────────────
if st.session_state.scan_done and st.session_state.result:
    res = st.session_state.result
    weeks = res["weeks"]
    bookings = res["bookings"]
    passenger = res["passenger_name"]

    rail_divider()

    total_spend = sum(b["amount_raw"] for b in bookings)
    weeks_travelled = sum(1 for w in weeks if w["bookings"])
    total_journeys = len(bookings)
    avg_per_trip = total_spend / total_journeys if total_journeys else 0

    # ── Summary metrics ───────────────────────────────────────────────────────
    st.markdown(f"""
    <div class="card">
      <div class="card-title">📊 Summary — {passenger}</div>
      <div class="metric-row">
        <div class="metric-box">
          <div class="metric-value">€{total_spend:.2f}</div>
          <div class="metric-label">Total spend (your share)</div>
        </div>
        <div class="metric-box">
          <div class="metric-value">{total_journeys}</div>
          <div class="metric-label">Journeys</div>
        </div>
        <div class="metric-box">
          <div class="metric-value">{weeks_travelled}</div>
          <div class="metric-label">Weeks travelled</div>
        </div>
        <div class="metric-box">
          <div class="metric-value">€{avg_per_trip:.2f}</div>
          <div class="metric-label">Avg cost per journey</div>
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Download ──────────────────────────────────────────────────────────────
    dl_col, _ = st.columns([2, 4])
    with dl_col:
        st.download_button(
            label="⬇️  Download Excel Report",
            data=res["excel"],
            file_name=f"irish_rail_{passenger.replace(' ', '_').lower()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    rail_divider()

    # ── Journey table ─────────────────────────────────────────────────────────
    st.markdown('<div class="card"><div class="card-title">🗓️ Week-by-Week Breakdown</div>', unsafe_allow_html=True)

    rows_html = ""
    for week in reversed(weeks):  # most recent first
        monday = week["monday"]
        wk_bookings = week["bookings"]
        travelled = len(wk_bookings) > 0
        label = week_label(monday)

        if travelled:
            week_spend = sum(b["amount_raw"] for b in wk_bookings)
            badge = '<span class="badge-yes">✔ Travelled</span>'
            spend_cell = f'<span class="amount-chip">€{week_spend:.2f}</span>'
            journeys_html = ""
            for b in wk_bookings:
                day = b["travel_date"].strftime("%a %d %b") if b["travel_date"] else b["travel_date_str"]
                route = f"{b['origin']} → {b['destination']}" if b["origin"] and b["destination"] else "—"
                ref = b["booking_ref"] or "—"
                journeys_html += f"""
                <tr>
                  <td style="padding-left:1.8rem; color:#6b7280; font-size:0.82rem;">{day}</td>
                  <td><span class="route-pill">{route}</span></td>
                  <td class="amount-chip">{b['amount_display']}</td>
                  <td class="ref-chip">{ref}</td>
                </tr>"""
            rows_html += f"""
            <tr style="background:#f6fff9;">
              <td style="font-weight:600; color:#111;">{label}</td>
              <td>{badge}</td>
              <td>{spend_cell}</td>
              <td></td>
            </tr>
            {journeys_html}"""
        else:
            rows_html += f"""
            <tr>
              <td style="color:#9ca3af;">{label}</td>
              <td><span class="badge-no">— No travel</span></td>
              <td class="no-travel">—</td>
              <td></td>
            </tr>"""

    st.markdown(f"""
    <table class="journey-table">
      <thead>
        <tr>
          <th>Week</th>
          <th>Status</th>
          <th>Spend</th>
          <th>Booking ref</th>
        </tr>
      </thead>
      <tbody>
        {rows_html}
      </tbody>
    </table>
    """, unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

    # ── Reset ─────────────────────────────────────────────────────────────────
    rail_divider()
    if st.button("🔄 Run a new scan"):
        st.session_state.scan_done = False
        st.session_state.result = {}
        st.rerun()

# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown("""
<div style="text-align:center; margin-top:3rem; padding-top:1.2rem;
            border-top:1px solid #e5e7eb; color:#9ca3af; font-size:0.78rem;">
  🚆 Irish Rail Spend Tracker &nbsp;·&nbsp;
  Runs entirely on your machine &nbsp;·&nbsp;
  <a href="https://github.com/sebmarchand44/irish-rail-spend-tracker"
     style="color:#00A651; text-decoration:none;">GitHub</a>
</div>
""", unsafe_allow_html=True)
