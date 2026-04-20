"""
Clinical Sample Monitor - Google Sheets Sync
Synchronizes the SQLite master database to Google Sheets
for team members who need a spreadsheet view.

Setup required (one-time):
    pip3 install gspread google-auth
    Place your Google service account credentials JSON at:
    config/google_credentials.json
"""

import sqlite3
import os
import json
from datetime import datetime

# ── Paths ──────────────────────────────────────────────────
BASE_DIR    = os.path.dirname(os.path.dirname(__file__))
DB_PATH     = os.path.join(BASE_DIR, "database", "clinical_samples.db")
CREDS_PATH  = os.path.join(BASE_DIR, "config", "google_credentials.json")
CONFIG_PATH = os.path.join(BASE_DIR, "config", "sheets_config.json")


# ── Load config ────────────────────────────────────────────
def load_config():
    """Load Google Sheet ID and sheet tab names from config."""
    if not os.path.exists(CONFIG_PATH):
        default = {
            "spreadsheet_id": "YOUR_GOOGLE_SHEET_ID_HERE",
            "tabs": {
                "samples_overview": "Sample Overview",
                "sample_status":    "Sample Status",
                "qc_metrics":       "QC Metrics",
                "billing_summary":  "Billing Summary",
                "projects":         "Projects",
                "clients":          "Clients"
            }
        }
        os.makedirs(os.path.dirname(CONFIG_PATH), exist_ok=True)
        with open(CONFIG_PATH, "w") as f:
            json.dump(default, f, indent=2)
        print(f"⚠️  Config created at {CONFIG_PATH}")
        print("    Please fill in your Google Sheet ID before syncing.")
        return default
    with open(CONFIG_PATH) as f:
        return json.load(f)


# ── Google Sheets connection ───────────────────────────────
def get_sheet_client():
    """Authenticate and return a gspread client."""
    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except ImportError:
        print("❌ Missing packages. Run: pip3 install gspread google-auth")
        return None

    if not os.path.exists(CREDS_PATH):
        print(f"❌ Credentials not found at: {CREDS_PATH}")
        print("   See SETUP.md for instructions on creating a service account.")
        return None

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds  = Credentials.from_service_account_file(CREDS_PATH, scopes=scopes)
    client = gspread.authorize(creds)
    return client


# ── Data extractors from SQLite ────────────────────────────
def get_samples_overview(conn):
    """Main sample overview — most commonly viewed sheet."""
    cursor = conn.cursor()
    cursor.execute("""
        SELECT
            s.lab_id,
            s.sample_id,
            s.project_id,
            c.client_name,
            s.service_type_id,
            s.sample_type,
            s.cap_id,
            s.pickup_datetime,
            s.lab_in_datetime,
            s.lab_in_operator,
            s.sr_condition,
            s.receiving_remarks,
            -- Latest status
            (SELECT sst.status
             FROM sample_status_tracking sst
             WHERE sst.sample_id = s.sample_id
             ORDER BY sst.timestamp DESC LIMIT 1) AS current_status,
            (SELECT sst.timestamp
             FROM sample_status_tracking sst
             WHERE sst.sample_id = s.sample_id
             ORDER BY sst.timestamp DESC LIMIT 1) AS status_updated,
            -- Notification email date
            (SELECT sr.notification_email_date
             FROM sequencing_runs sr
             WHERE sr.sample_id = s.sample_id
             ORDER BY sr.run_number DESC LIMIT 1) AS notification_email_date,
            s.transit_date,
            s.awb,
            s.created_date
        FROM samples s
        JOIN clients c ON s.client_id = c.client_id
        ORDER BY s.created_date DESC
    """)
    rows = cursor.fetchall()
    headers = [
        "LabID", "SampleID", "ProjectID", "Client", "ServiceType",
        "SampleType", "CapID", "PickupDatetime", "LabInDatetime",
        "LabInOperator", "SRCondition", "ReceivingRemarks",
        "CurrentStatus", "StatusUpdated", "NotificationEmailDate",
        "TransitDate", "AWB", "CreatedDate"
    ]
    return headers, rows


def get_sample_status(conn):
    """Full status audit trail."""
    cursor = conn.cursor()
    cursor.execute("""
        SELECT
            sst.tracking_id,
            s.lab_id,
            s.sample_id,
            s.project_id,
            sst.status,
            sst.timestamp,
            sst.operator,
            sst.remarks
        FROM sample_status_tracking sst
        JOIN samples s ON sst.sample_id = s.sample_id
        ORDER BY sst.timestamp DESC
        LIMIT 5000
    """)
    rows = cursor.fetchall()
    headers = ["TrackingID", "LabID", "SampleID", "ProjectID",
               "Status", "Timestamp", "Operator", "Remarks"]
    return headers, rows


def get_qc_metrics(conn):
    """QC metrics per sample."""
    cursor = conn.cursor()
    cursor.execute("""
        SELECT
            s.lab_id,
            s.sample_id,
            s.project_id,
            c.client_name,
            q.run_id,
            q.qc_conclusion,
            q.qc_failreason,
            q.mapping_ratio,
            q.target_depth,
            q.base_Q30,
            q.gc_rate,
            q.insert_size,
            q.dup_ratio,
            q.contamination_level,
            q.gender_analysed,
            q.gender_record,
            q.pcr_qc,
            q.created_date
        FROM qc_metrics q
        JOIN samples s ON q.sample_id = s.sample_id
        JOIN clients c ON s.client_id = c.client_id
        ORDER BY q.created_date DESC
    """)
    rows = cursor.fetchall()
    headers = [
        "LabID", "SampleID", "ProjectID", "Client", "RunID",
        "QC_Conclusion", "QC_FailReason", "MappingRatio", "TargetDepth",
        "BaseQ30", "GC_Rate", "InsertSize", "DupRatio", "ContaminationLevel",
        "GenderAnalysed", "GenderRecord", "PCR_QC", "CreatedDate"
    ]
    return headers, rows


def get_billing_summary(conn):
    """Billing overview per client and period."""
    cursor = conn.cursor()
    cursor.execute("""
        SELECT
            b.billing_id,
            c.client_name,
            s.lab_id,
            s.sample_id,
            b.service_type_id,
            b.billing_period_start,
            b.billing_period_end,
            b.notification_email_date,
            b.billing_units,
            b.unit_price,
            b.currency,
            b.total_amount,
            b.pdf_row_label,
            CASE WHEN b.billing_email_sent = 1 THEN 'Yes' ELSE 'No' END AS billing_email_sent,
            b.billing_email_date,
            b.remarks
        FROM billing b
        JOIN samples s  ON b.sample_id  = s.sample_id
        JOIN clients c  ON b.client_id  = c.client_id
        ORDER BY b.billing_period_start DESC, c.client_name
    """)
    rows = cursor.fetchall()
    headers = [
        "BillingID", "Client", "LabID", "SampleID", "ServiceType",
        "PeriodStart", "PeriodEnd", "NotificationEmailDate",
        "BillingUnits", "UnitPrice", "Currency", "TotalAmount",
        "PDF_RowLabel", "BillingEmailSent", "BillingEmailDate", "Remarks"
    ]
    return headers, rows


def get_projects(conn):
    cursor = conn.cursor()
    cursor.execute("""
        SELECT
            p.project_id,
            c.client_name,
            p.service_type_id,
            p.collection_batch,
            p.submission_form_format,
            p.created_date,
            COUNT(s.sample_id) AS sample_count
        FROM projects p
        JOIN clients c ON p.client_id = c.client_id
        LEFT JOIN samples s ON s.project_id = p.project_id
        GROUP BY p.project_id
        ORDER BY p.created_date DESC
    """)
    rows = cursor.fetchall()
    headers = ["ProjectID", "Client", "ServiceType", "CollectionBatch",
               "FormFormat", "CreatedDate", "SampleCount"]
    return headers, rows


def get_clients(conn):
    cursor = conn.cursor()
    cursor.execute("""
        SELECT
            client_id, client_name, contact_email,
            billing_currency, billing_period_day_start,
            requires_additional_qc, data_confidentiality_level,
            lab_id_prefix, is_active, created_date
        FROM clients
        ORDER BY client_name
    """)
    rows = cursor.fetchall()
    headers = [
        "ClientID", "ClientName", "ContactEmail",
        "BillingCurrency", "BillingPeriodDayStart",
        "RequiresAdditionalQC", "ConfidentialityLevel",
        "LabIDPrefix", "IsActive", "CreatedDate"
    ]
    return headers, rows


# ── Sheet writer ───────────────────────────────────────────
def write_to_sheet(spreadsheet, tab_name, headers, rows):
    """Clear and rewrite a sheet tab with fresh data."""
    try:
        worksheet = spreadsheet.worksheet(tab_name)
    except Exception:
        worksheet = spreadsheet.add_worksheet(title=tab_name, rows=5000, cols=50)

    # Build data: header row + all data rows
    data = [headers] + [list(r) for r in rows]

    # Clear and update
    worksheet.clear()
    if data:
        worksheet.update(data, value_input_option="USER_ENTERED")

    # Bold header row
    worksheet.format("1:1", {"textFormat": {"bold": True}})
    print(f"   ✅ '{tab_name}' — {len(rows)} rows written")


# ── Main sync function ─────────────────────────────────────
def sync_to_sheets(tabs_to_sync=None):
    """
    Sync SQLite → Google Sheets.
    tabs_to_sync: list of tab keys to sync, or None for all.
    """
    print("\n" + "=" * 50)
    print("  Google Sheets Sync")
    print(f"  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 50)

    config = load_config()
    if config["spreadsheet_id"] == "YOUR_GOOGLE_SHEET_ID_HERE":
        print("\n⚠️  Please set your spreadsheet_id in config/sheets_config.json")
        return False

    client = get_sheet_client()
    if not client:
        return False

    try:
        spreadsheet = client.open_by_key(config["spreadsheet_id"])
        print(f"\n📊 Connected to: {spreadsheet.title}")
    except Exception as e:
        print(f"❌ Cannot open spreadsheet: {e}")
        return False

    conn = sqlite3.connect(DB_PATH)
    tabs  = config["tabs"]
    synced = 0

    tab_functions = {
        "samples_overview": (get_samples_overview, tabs.get("samples_overview")),
        "sample_status":    (get_sample_status,    tabs.get("sample_status")),
        "qc_metrics":       (get_qc_metrics,       tabs.get("qc_metrics")),
        "billing_summary":  (get_billing_summary,  tabs.get("billing_summary")),
        "projects":         (get_projects,          tabs.get("projects")),
        "clients":          (get_clients,           tabs.get("clients")),
    }

    print("\n📤 Syncing tabs:")
    for key, (func, tab_name) in tab_functions.items():
        if tabs_to_sync and key not in tabs_to_sync:
            continue
        if not tab_name:
            continue
        try:
            headers, rows = func(conn)
            write_to_sheet(spreadsheet, tab_name, headers, rows)
            synced += 1
        except Exception as e:
            print(f"   ❌ '{tab_name}' failed: {e}")

    conn.close()
    print(f"\n✅ Sync complete — {synced} tabs updated")
    print(f"🔗 https://docs.google.com/spreadsheets/d/{config['spreadsheet_id']}")
    return True


if __name__ == "__main__":
    sync_to_sheets()
