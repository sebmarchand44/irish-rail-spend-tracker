#!/usr/bin/env python3
"""
Irish Rail Spend Tracker — Streamlit UI with web OAuth
"""

import io
import os
import re
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

from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request

from main import (
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

# ── Config ────────────────────────────────────────────────────────────────────
SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]
REDIRECT_URI = "https://irish-rail-spend-tracker.streamlit.app"

# Fall back to env vars for local dev without secrets.toml
CLIENT_ID = st.secrets.get("GOOGLE_CLIENT_ID", os.environ.get("GOOGLE_CLIENT_ID", ""))
CLIENT_SECRET = st.secrets.get("GOOGLE_CLIENT_SECRET", os.environ.get("GOOGLE_CLIENT_SECRET", ""))

# ── Custom CSS ────────────────────────────────────────────────────────────────
HERO_IMAGE = "https://images.unsplash.com/photo-1679410153289-2c4e3fa60a3a?w=1600&q=80&auto=format&fit=crop"

st.markdown(f"""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
  html, body, [class*="css"] {{ font-family: 'Inter', sans-serif; }}
  #MainMenu, footer, header {{visibility: hidden;}}
  .block-container {{padding-top: 0 !important; max-width: 960px;}}

  .hero {{
    position: relative; width: 100%; height: 320px;
    background: url('{HERO_IMAGE}') center/cover no-repeat;
    border-radius: 0 0 16px 16px; overflow: hidden; margin-bottom: 2rem;
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
  .hero-sub {{ color: rgba(255,255,255,0.80); font-size: 1.05rem; margin: 0; }}

  .card {{
    background: #fff; border: 1px solid #e5e7eb;
    border-radius: 12px; padding: 1.6rem 2rem; margin-bottom: 1.2rem;
  }}
  .card-title {{
    font-size: 1rem; font-weight: 700; color: #111;
    margin: 0 0 1rem; display: flex; align-items: center; gap: 8px;
  }}

  .google-btn {{
    display: inline-flex; align-items: center; gap: 12px;
    background: #fff; border: 2px solid #e5e7eb; border-radius: 8px;
    padding: 12px 24px; font-size: 1rem; font-weight: 600;
    color: #374151; text-decoration: none; cursor: pointer;
    transition: border-color 0.2s, box-shadow 0.2s;
  }}
  .google-btn:hover {{ border-color: #00A651; box-shadow: 0 0 0 3px rgba(0,166,81,0.15); }}
  .google-logo {{ width: 22px; height: 22px; }}

  .metric-row {{ display: flex; gap: 1rem; flex-wrap: wrap; margin-bottom: 1.4rem; }}
  .metric-box {{
    flex: 1; min-width: 140px;
    background: #f0faf4; border: 1px solid #bbf0d0;
    border-radius: 10px; padding: 1rem 1.2rem; text-align: center;
  }}
  .metric-value {{ font-size: 1.8rem; font-weight: 700; color: #065f46; line-height: 1; }}
  .metric-label {{ font-size: 0.78rem; color: #6b7280; margin-top: 4px; font-weight: 500; }}

  .journey-table {{ width: 100%; border-collapse: collapse; font-size: 0.88rem; }}
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
    border-radius: 20px; font-size: 0.82rem; font-weight: 500; white-space: nowrap;
  }}
  .amount-chip {{ font-weight: 600; color: #065f46; }}
  .ref-chip {{ font-size: 0.78rem; color: #9ca3af; font-family: monospace; }}
  .badge-yes {{
    background: #d1fae5; color: #065f46; border-radius: 20px;
    padding: 2px 10px; font-size: 0.78rem; font-weight: 600;
  }}
  .badge-no {{
    background: #f3f4f6; color: #9ca3af; border-radius: 20px;
    padding: 2px 10px; font-size: 0.78rem;
  }}
  .log-box {{
    background: #0d1117; color: #58d68d; font-family: monospace;
    font-size: 0.8rem; border-radius: 8px; padding: 1rem 1.2rem;
    height: 220px; overflow-y: auto; white-space: pre-wrap;
    border: 1px solid #1e3a2e;
  }}

  div.stButton > button {{
    background: #00A651 !important; color: #fff !important;
    border: none !important; border-radius: 8px !important;
    font-weight: 600 !important; padding: 0.55rem 1.4rem !important;
  }}
  div.stButton > button:hover {{ background: #008c44 !important; }}
  div.stDownloadButton > button {{
    background: #003d1a !important; color: #fff !important;
    border: none !important; border-radius: 8px !important; font-weight: 600 !important;
  }}

  .track {{ display: flex; align-items: center; gap: 0; margin: 1.5rem 0; }}
  .track-tie {{
    width: 18px; height: 10px; background: #003d1a;
    margin: 0 4px; border-radius: 2px; flex-shrink: 0;
  }}
  .track-line {{ flex: 1; height: 3px; background: #003d1a; }}

  .signed-in-bar {{
    display: flex; align-items: center; gap: 10px;
    background: #f0faf4; border: 1px solid #bbf0d0;
    border-radius: 8px; padding: 10px 16px; font-size: 0.88rem;
    color: #065f46; font-weight: 500;
  }}
</style>
""", unsafe_allow_html=True)


# ── Helpers ───────────────────────────────────────────────────────────────────
def rail_divider():
    ties = "".join(['<div class="track-tie"></div>'] * 22)
    st.markdown(f'<div class="track">{ties}</div>', unsafe_allow_html=True)


def make_flow():
    """Build an OAuth Flow from secrets."""
    client_config = {
        "web": {
            "client_id": CLIENT_ID,
            "client_secret": CLIENT_SECRET,
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "redirect_uris": [REDIRECT_URI],
        }
    }
    return Flow.from_client_config(client_config, scopes=SCOPES, redirect_uri=REDIRECT_URI)


def get_auth_url():
    flow = make_flow()
    auth_url, state = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
    )
    st.session_state["oauth_state"] = state
    return auth_url


def exchange_code(code):
    """Exchange the auth code for credentials and store in session."""
    flow = make_flow()
    flow.fetch_token(code=code)
    creds = flow.credentials
    st.session_state["google_creds"] = {
        "token": creds.token,
        "refresh_token": creds.refresh_token,
        "token_uri": creds.token_uri,
        "client_id": creds.client_id,
        "client_secret": creds.client_secret,
        "scopes": list(creds.scopes) if creds.scopes else SCOPES,
    }


def get_service():
    """Rebuild a Gmail service from stored session credentials."""
    c = st.session_state["google_creds"]
    creds = Credentials(
        token=c["token"],
        refresh_token=c["refresh_token"],
        token_uri=c["token_uri"],
        client_id=c["client_id"],
        client_secret=c["client_secret"],
        scopes=c["scopes"],
    )
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
        st.session_state["google_creds"]["token"] = creds.token
    return build("gmail", "v1", credentials=creds)


def run_scan(service, passenger_name, months, log_lines):
    """Run the full Gmail scan, return (weeks, bookings) or raise."""
    end_date = datetime.now()
    start_date = end_date - timedelta(days=months * 30)
    after_str = start_date.strftime("%Y/%m/%d")
    target_norm = norm_name(passenger_name)

    log_lines.append(f"🔍 Searching from {start_date.strftime('%d %b %Y')} to today")
    log_lines.append(f"👤 Passenger: {passenger_name}\n")

    # Step 1 — cancellations
    log_lines.append("Step 1/2 — Finding cancelled bookings...")
    cancelled_refs = set()
    cancel_msgs = search_messages(service, f"from:{IRISH_RAIL_SENDER} subject:cancelled after:{after_str}")
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
    log_lines.append(f"   {len(cancelled_refs)} cancelled booking(s) excluded.\n")

    # Step 2 — booking confirmations
    log_lines.append("Step 2/2 — Fetching booking confirmation emails...")
    booking_query = (
        f"from:{IRISH_RAIL_SENDER} "
        f"subject:\"Thank you for booking\" "
        f"after:{after_str}"
    )
    booking_msgs = search_messages(service, booking_query)
    log_lines.append(f"   Found {len(booking_msgs)} booking email(s).\n")

    bookings = []
    excluded_cancelled = 0

    for i, msg in enumerate(booking_msgs):
        log_lines.append(f"   [{i+1}/{len(booking_msgs)}] Reading...")
        try:
            body, email_date = get_message_body(service, msg["id"])
            booking = parse_booking_email(body, email_date, msg["id"], target_norm)
            if not booking:
                continue
            if booking["booking_ref"] and booking["booking_ref"] in cancelled_refs:
                excluded_cancelled += 1
                log_lines.append(f"   🚫 [{booking['booking_ref']}] Cancelled — skipped")
                continue
            bookings.append(booking)
            route = f"{booking['origin']} → {booking['destination']}" if booking["origin"] else "route unknown"
            log_lines.append(f"   ✅ [{booking['booking_ref']}] {booking['travel_date_str']} | {route} | {booking['amount_display']}")
        except Exception as e:
            log_lines.append(f"   ⚠️  Error: {e}")

    log_lines.append(f"\n✅ {len(bookings)} journey(s) found. {excluded_cancelled} cancelled excluded.\n")

    if not bookings:
        raise ValueError("No bookings found. Check the passenger name spelling and try again.")

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

    log_lines.append("📊 Generating Excel report...")
    return weeks, bookings


def build_excel(weeks, passenger_name):
    tmp = "irish_rail_tmp.xlsx"
    max_j = max((len(w["bookings"]) for w in weeks), default=1)
    write_excel(weeks, max_j, passenger_name, tmp)
    with open(tmp, "rb") as f:
        data = f.read()
    os.remove(tmp)
    return data


# ── Session state defaults ────────────────────────────────────────────────────
for key, default in [
    ("google_creds", None),
    ("scan_done", False),
    ("scan_result", {}),
    ("scan_logs", []),
]:
    if key not in st.session_state:
        st.session_state[key] = default


# ── OAuth callback handling ───────────────────────────────────────────────────
params = st.query_params
if "code" in params and st.session_state["google_creds"] is None:
    try:
        exchange_code(params["code"])
        st.query_params.clear()
        st.rerun()
    except Exception as e:
        st.error(f"OAuth error: {e}")
        st.stop()


# ── Hero ──────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero">
  <div class="hero-overlay">
    <div class="hero-tag">🚆 Spend Tracker</div>
    <div class="hero-title">Irish Rail<br>Travel Report</div>
    <div class="hero-sub">Connect your Gmail, see every journey, know what you spent.</div>
  </div>
</div>
""", unsafe_allow_html=True)


# ── Step 1 — Sign in ──────────────────────────────────────────────────────────
st.markdown('<div class="card"><div class="card-title">1 &nbsp; Sign in with Google</div>', unsafe_allow_html=True)

if st.session_state["google_creds"]:
    st.markdown('<div class="signed-in-bar">✅ &nbsp; Gmail connected — you\'re signed in.</div>', unsafe_allow_html=True)
    col_signout, _ = st.columns([2, 6])
    with col_signout:
        if st.button("Sign out", key="signout"):
            st.session_state["google_creds"] = None
            st.session_state["scan_done"] = False
            st.session_state["scan_result"] = {}
            st.rerun()
else:
    auth_url = get_auth_url()
    st.markdown(f"""
    <p style="color:#6b7280; font-size:0.88rem; margin-bottom:1rem;">
      Click below to grant read-only access to your Gmail.
      Your emails never leave your browser session.
    </p>
    <a href="{auth_url}" class="google-btn">
      <svg class="google-logo" viewBox="0 0 48 48">
        <path fill="#EA4335" d="M24 9.5c3.54 0 6.71 1.22 9.21 3.6l6.85-6.85C35.9 2.38 30.47 0 24 0 14.62 0 6.51 5.38 2.56 13.22l7.98 6.19C12.43 13.72 17.74 9.5 24 9.5z"/>
        <path fill="#4285F4" d="M46.98 24.55c0-1.57-.15-3.09-.38-4.55H24v9.02h12.94c-.58 2.96-2.26 5.48-4.78 7.18l7.73 6c4.51-4.18 7.09-10.36 7.09-17.65z"/>
        <path fill="#FBBC05" d="M10.53 28.59c-.48-1.45-.76-2.99-.76-4.59s.27-3.14.76-4.59l-7.98-6.19C.92 16.46 0 20.12 0 24c0 3.88.92 7.54 2.56 10.78l7.97-6.19z"/>
        <path fill="#34A853" d="M24 48c6.48 0 11.93-2.13 15.89-5.81l-7.73-6c-2.15 1.45-4.92 2.3-8.16 2.3-6.26 0-11.57-4.22-13.47-9.91l-7.98 6.19C6.51 42.62 14.62 48 24 48z"/>
        <path fill="none" d="M0 0h48v48H0z"/>
      </svg>
      Sign in with Google
    </a>
    """, unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)

rail_divider()


# ── Step 2 — Options (only shown when signed in) ──────────────────────────────
if st.session_state["google_creds"]:
    st.markdown('<div class="card"><div class="card-title">2 &nbsp; Search Options</div>', unsafe_allow_html=True)

    col1, col2 = st.columns([3, 1])
    with col1:
        passenger_name = st.text_input(
            "Passenger name",
            placeholder="e.g. Séamus O'Brien",
            help="Exactly as it appears on your Irish Rail ticket.",
        )
    with col2:
        months = st.number_input("Months to look back", min_value=1, max_value=60, value=13)

    st.markdown('</div>', unsafe_allow_html=True)

    rail_divider()

    # ── Step 3 — Run ──────────────────────────────────────────────────────────
    st.markdown('<div class="card"><div class="card-title">3 &nbsp; Run the Scan</div>', unsafe_allow_html=True)

    run_col, _ = st.columns([2, 4])
    with run_col:
        run_btn = st.button("🚆 Scan Gmail now", use_container_width=True)

    st.markdown('</div>', unsafe_allow_html=True)

    if run_btn:
        if not passenger_name.strip():
            st.error("Please enter a passenger name.")
        else:
            log_lines = []
            result = {}
            with st.spinner("Scanning your Gmail inbox…"):
                try:
                    service = get_service()
                    weeks, bookings = run_scan(service, passenger_name.strip(), months, log_lines)
                    excel_data = build_excel(weeks, passenger_name.strip())
                    result = {
                        "weeks": weeks,
                        "bookings": bookings,
                        "excel": excel_data,
                        "passenger_name": passenger_name.strip(),
                    }
                    st.session_state["scan_done"] = True
                    st.session_state["scan_result"] = result
                    st.session_state["scan_logs"] = log_lines
                except ValueError as e:
                    st.error(str(e))
                    log_lines.append(f"\n❌ {e}")
                except Exception as e:
                    st.error(f"Unexpected error: {e}")
                    log_lines.append(f"\n❌ {e}")

            log_text = "\n".join(log_lines)
            st.markdown(f'<div class="log-box">{log_text}</div>', unsafe_allow_html=True)

            if st.session_state["scan_done"]:
                st.rerun()


# ── Results ───────────────────────────────────────────────────────────────────
if st.session_state["scan_done"] and st.session_state["scan_result"]:
    res = st.session_state["scan_result"]
    weeks = res["weeks"]
    bookings = res["bookings"]
    passenger = res["passenger_name"]

    rail_divider()

    total_spend = sum(b["amount_raw"] for b in bookings)
    weeks_travelled = sum(1 for w in weeks if w["bookings"])
    total_journeys = len(bookings)
    avg_per_trip = total_spend / total_journeys if total_journeys else 0

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
    for week in reversed(weeks):
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
                journeys_html += f"""
                <tr>
                  <td style="padding-left:1.8rem;color:#6b7280;font-size:0.82rem;">{day}</td>
                  <td><span class="route-pill">{route}</span></td>
                  <td class="amount-chip">{b['amount_display']}</td>
                  <td class="ref-chip">{b['booking_ref'] or '—'}</td>
                </tr>"""
            rows_html += f"""
            <tr style="background:#f6fff9;">
              <td style="font-weight:600;color:#111;">{label}</td>
              <td>{badge}</td><td>{spend_cell}</td><td></td>
            </tr>{journeys_html}"""
        else:
            rows_html += f"""
            <tr>
              <td style="color:#9ca3af;">{label}</td>
              <td><span class="badge-no">— No travel</span></td>
              <td style="color:#9ca3af;font-style:italic;font-size:0.82rem;">—</td>
              <td></td>
            </tr>"""

    st.markdown(f"""
    <table class="journey-table">
      <thead><tr>
        <th>Week</th><th>Status</th><th>Spend</th><th>Booking ref</th>
      </tr></thead>
      <tbody>{rows_html}</tbody>
    </table>
    """, unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

    rail_divider()
    if st.button("🔄 Run a new scan"):
        st.session_state["scan_done"] = False
        st.session_state["scan_result"] = {}
        st.rerun()


# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown("""
<div style="text-align:center;margin-top:3rem;padding-top:1.2rem;
            border-top:1px solid #e5e7eb;color:#9ca3af;font-size:0.78rem;">
  🚆 Irish Rail Spend Tracker &nbsp;·&nbsp; Read-only Gmail access &nbsp;·&nbsp;
  <a href="https://github.com/sebmarchand44/irish-rail-spend-tracker"
     style="color:#00A651;text-decoration:none;">GitHub</a>
</div>
""", unsafe_allow_html=True)
