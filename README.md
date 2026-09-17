# 🚆 Irish Rail Spend Tracker

Analyse your Irish Rail travel and spending directly from your Gmail inbox.

Scans your booking confirmation emails from `iebookinginfo@irishrail.ie`, filters by passenger name, excludes cancelled bookings, and produces a **week-by-week Excel spreadsheet** showing when you travelled, your route, and what you paid — with your share calculated correctly even when travelling with others.

---

## Example output

| Week | Travelled? | Week Spend | J1 Date | J1 Route | J1 My Amount | J1 Booking Ref |
|---|---|---|---|---|---|---|
| 04 Aug 2025 – 10 Aug 2025 | ✅ Yes | €30.00 | Mon 04 Aug | Thurles → Dublin Heuston | €30.00 | 81234567 |
| 11 Aug 2025 – 17 Aug 2025 | — | — | | | | |
| 18 Aug 2025 – 24 Aug 2025 | ✅ Yes | €15.00 | Wed 20 Aug | Thurles → Dublin Heuston | €15.00 (1/2 of €30.00) | 81345678 |

---

## Features

- ✅ Filters by any passenger name (handles accented characters)
- 🚫 Automatically detects and excludes cancelled bookings
- 👥 Pro-rates the ticket cost by number of passengers (so you only see your share)
- 📅 Week-by-week view — every week shown, even ones with no travel
- 📊 Excel output with a summary sheet (total spend, average per trip, etc.)
- 🔒 Runs entirely locally — your data never leaves your machine

---

## Requirements

- Python 3.8 or later
- A Gmail account that receives Irish Rail booking emails
- A Google Cloud project with the Gmail API enabled (free — see setup below)

---

## Setup

### Step 1 — Clone the repo

```bash
git clone https://github.com/sebmarchand44/irish-rail-spend-tracker.git
cd irish-rail-spend-tracker
```

### Step 2 — Install dependencies

```bash
pip install -r requirements.txt
```

### Step 3 — Set up Gmail API credentials

This is a one-time setup. You need to create a free Google Cloud project to allow the script to read your Gmail.

1. Go to [https://console.cloud.google.com](https://console.cloud.google.com)
2. Click **"Select a project"** → **"New Project"**
3. Name it anything (e.g. `irish-rail-tracker`) → click **Create**
4. In the left menu go to **APIs & Services** → **Library**
5. Search for **"Gmail API"** → click it → click **Enable**
6. Go to **APIs & Services** → **OAuth consent screen**
   - Choose **External** → click **Create**
   - Fill in App name (e.g. `Irish Rail Tracker`) and your email → click **Save and Continue**
   - Skip Scopes → click **Save and Continue**
   - Under **Test users**, click **Add users** → add your Gmail address → click **Save**
7. Go to **APIs & Services** → **Credentials**
   - Click **Create Credentials** → **OAuth client ID**
   - Application type: **Desktop app**
   - Name it anything → click **Create**
8. Click **Download JSON** on the credential that was just created
9. Rename the downloaded file to `credentials.json` and place it in the project folder

### Step 4 — Run the script

```bash
python main.py
```

The first time you run it, a browser window will open asking you to sign in to Google and grant read-only Gmail access. After that, your credentials are cached in `token.pickle` and you won't be asked again.

---

## Usage

### Basic usage (interactive)

```bash
python main.py
```

You'll be prompted to enter the passenger name.

### With command-line options

```bash
# Specify passenger name directly
python main.py --name "Sebastien Marchand"

# Look back 6 months instead of 13
python main.py --name "Sebastien Marchand" --months 6

# Custom output filename
python main.py --name "Sebastien Marchand" --output my_travel_2026.xlsx

# All options together
python main.py -n "Sebastien Marchand" -m 13 -o travel_report.xlsx
```

### Options

| Option | Short | Default | Description |
|---|---|---|---|
| `--name` | `-n` | (prompt) | Passenger name to filter for |
| `--months` | `-m` | `13` | How many months to look back |
| `--output` | `-o` | `irish_rail_travel.xlsx` | Output Excel filename |

---

## Privacy & security

- The script uses **read-only** Gmail access (`gmail.readonly` scope) — it cannot send, delete or modify any emails
- Your OAuth token is stored locally in `token.pickle` — **never share this file**
- `credentials.json` and `token.pickle` are both in `.gitignore` so they can't be accidentally committed
- No data is sent anywhere — everything runs locally on your machine

---

## How it works

1. **Finds cancellation emails** — searches for emails from Irish Rail with "cancelled" in the subject, extracts the booking reference from each one
2. **Finds booking confirmation emails** — searches for "Thank you for booking" emails from Irish Rail
3. **Filters by passenger** — reads each booking email and checks the Passengers section for the target name
4. **Excludes cancelled bookings** — any booking ref found in step 1 is dropped
5. **Pro-rates the amount** — divides the total paid by the number of passengers to get your individual share
6. **Builds a week grid** — every week from the start date to today gets a row, whether you travelled or not
7. **Writes Excel** — outputs a formatted spreadsheet with a summary sheet

---

## Troubleshooting

**"credentials.json not found"**
→ Follow Step 3 in the setup instructions above.

**"Access blocked: This app's request is invalid"**
→ Make sure you added your Gmail address as a Test User in the OAuth consent screen (Step 3.6 above).

**Passenger not found / zero results**
→ Check the exact spelling of the name as it appears on your Irish Rail ticket. Try a partial name (e.g. just the surname).

**Some bookings seem to be missing**
→ The script looks back 13 months by default. Use `--months 24` to go further back.

---

## Contributing

Pull requests welcome. Some ideas for future improvements:

- Support for multiple passengers in one report
- Monthly summary chart (using matplotlib or plotly)
- Support for other email providers (Outlook, etc.)
- Tax saver break-even calculator

---

## Licence

MIT — do whatever you like with it.

---

*Built by [@sebmarchand44](https://github.com/sebmarchand44)*
