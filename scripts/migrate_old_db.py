"""
Clinical Sample Monitor - Old Google Sheet Migration Script
Reads your existing Google Sheet and imports data into the new SQLite database.

Usage:
    1. Export your existing Google Sheet as CSV → save to exports/old_database.csv
    2. Run: python3 scripts/migrate_old_db.py
    3. Review the migration report in logs/migration_report.txt
"""

import sqlite3
import csv
import os
import json
from datetime import datetime

BASE_DIR    = os.path.dirname(os.path.dirname(__file__))
DB_PATH     = os.path.join(BASE_DIR, "database", "clinical_samples.db")
CSV_PATH    = os.path.join(BASE_DIR, "exports", "old_database.csv")
LOG_PATH    = os.path.join(BASE_DIR, "logs", "migration_report.txt")

# ── Column mapping: old sheet column → new field ──────────
# Adjust these if your column headers differ slightly
OLD_COLUMN_MAP = {
    # Sample identity
    "LabID":                    "lab_id",
    "Lab Status":               "lab_status_old",       # → mapped to status tracking
    "SR Condition":             "sr_condition",
    "Sample Receiving Remarks": "receiving_remarks",
    "ProjectID":                "project_id",
    "Cap":                      "pickup_batch",
    "Pickup Batch":             "pickup_batch",
    "PickUp DateTime":          "pickup_datetime",
    "Sample Receiving Time":    "lab_in_datetime",
    "Sample Receiving Operator":"lab_in_operator",
    "CapID":                    "cap_id",
    "Sample Type":              "sample_type",
    "SampleID":                 "sample_id",            # col BF

    # Transit
    "Transit Date":             "transit_date",
    "AWB":                      "awb",
    "TW Receiving":             "tw_receiving",

    # Sequencing run
    "PipelineID":               "pipeline_id",
    "Index Plate":              "index_plate",
    "Index Position":           "index_position",
    "Index1":                   "index1",
    "Index2":                   "index2",
    "Seq Start DateTime":       "seq_start_datetime",
    "Seq RDSC":                 "seq_rdsc",
    "Order Operator":           "order_operator",
    "Order Time":               "order_time",
    "Process Time":             "process_time",
    "Process Operator":         "process_operator",
    "Process Remarks":          "process_remarks",
    "SendOut Datetime":         "sendout_datetime",
    "Sendout Operator":         "sendout_operator",
    "Run Name":                 "run_name",
    "Repeat Start Point":       "repeat_start_point",
    "Days of Life":             "days_of_life",
    "Repeat Opearator":         "repeat_operator",
    "Repeat Datetime":          "repeat_datetime",
    "NovaSeq \nRun Date":       "novaseq_run_date",
    "SNPQC status":             "snpqc_status",
    "Data upload date":         "data_upload_date",
    "Notification email date":  "notification_email_date",
    "Date for SFTP deletion":   "sftp_deletion_date",
    "Date for final deletion":  "final_deletion_date",

    # QC metrics
    "qc_conclusion":            "qc_conclusion",
    "qc_failreason":            "qc_failreason",
    "bases_all_trimmed":        "bases_all_trimmed",
    "mapping_ratio":            "mapping_ratio",
    "target_depth":             "target_depth",
    "base_Q30":                 "base_Q30",
    "gc_rate":                  "gc_rate",
    "insert_size":              "insert_size",
    "dup_ratio":                "dup_ratio",
    "titv_ratio":               "titv_ratio",
    "variant_ratio":            "variant_ratio",
    "hethom_ratio":             "hethom_ratio",
    "snpindel_ratio":           "snpindel_ratio",
    "gender_analysed":          "gender_analysed",
    "gender_record":            "gender_record",
    "contamination_level":      "contamination_level",
    "ontarget_ratio":           "ontarget_ratio",
    "bases_target_0x":          "bases_target_0x",
    "bases_target_lt_20x":      "bases_target_lt_20x",
    "pcr_qc":                   "pcr_qc",

    # Run metrics
    "runID":                    "run_id",
    "runQ30%":                  "run_q30_pct",
    "runPhixErrorRate%":        "phix_error_rate",
    "runBarcodeContaminationRate": "barcode_contamination_rate",
    "yield_gbp_upload":         "yield_gbp",
    "q30_pct_upload":           "q30_pct_upload",
    "depth_x_upload":           "depth_x_upload",
}

# SNP marker columns (key-value stored in snp_markers table)
SNP_MARKER_COLS = [
    "rs1042713_wes", "rs1042713_pcr",
    "rs1801133_wes", "rs1801133_pcr",
    "rs8192678_wes", "rs8192678_pcr",
    "rs4343_wes",    "rs4343_pcr",
    "rs713598_wes",  "rs713598_pcr",
]


def safe_float(val):
    try:
        return float(val) if val and val.strip() else None
    except (ValueError, AttributeError):
        return None


def safe_str(val):
    if val is None:
        return None
    v = str(val).strip()
    return v if v else None


def migrate(
    default_client_id="CLIENT001",
    default_client_name="Default Client",
    dry_run=False
):
    print("=" * 55)
    print("  Old Database Migration Tool")
    print(f"  Mode: {'DRY RUN (no changes)' if dry_run else 'LIVE'}")
    print("=" * 55)

    if not os.path.exists(CSV_PATH):
        print(f"\n❌ CSV not found at: {CSV_PATH}")
        print("   Export your Google Sheet as CSV and save it there.")
        return

    os.makedirs(os.path.dirname(LOG_PATH), exist_ok=True)

    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA foreign_keys = OFF")  # Off during bulk import
    cursor = conn.cursor()

    stats = {
        "total_rows": 0,
        "samples_imported": 0,
        "samples_skipped": 0,
        "runs_imported": 0,
        "qc_imported": 0,
        "snp_imported": 0,
        "errors": []
    }

    # Ensure default client exists
    cursor.execute(
        "INSERT OR IGNORE INTO clients (client_id, client_name) VALUES (?, ?)",
        (default_client_id, default_client_name)
    )

    with open(CSV_PATH, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)

        for row_num, row in enumerate(reader, start=2):
            stats["total_rows"] += 1

            # Map columns
            mapped = {}
            for old_col, new_field in OLD_COLUMN_MAP.items():
                if old_col in row:
                    mapped[new_field] = safe_str(row[old_col])

            # Determine IDs
            # In new schema: sample_id = customer name, lab_id = internal
            lab_id   = mapped.get("lab_id")
            sample_id = mapped.get("sample_id") or lab_id  # fallback to lab_id if no SampleID col
            project_id = mapped.get("project_id")

            if not lab_id and not sample_id:
                stats["samples_skipped"] += 1
                stats["errors"].append(f"Row {row_num}: No LabID or SampleID — skipped")
                continue

            # If no separate SampleID exists in old DB, use LabID as sample_id
            # and mark lab_id as the same (old logic)
            if not sample_id:
                sample_id = lab_id

            try:
                if not dry_run:
                    # ── Ensure project exists ─────────────────
                    if project_id:
                        cursor.execute(
                            "INSERT OR IGNORE INTO projects (project_id, client_id, service_type_id) VALUES (?, ?, ?)",
                            (project_id, default_client_id, "WES")
                        )

                    # ── Insert sample ─────────────────────────
                    cursor.execute("""
                        INSERT OR IGNORE INTO samples (
                            sample_id, lab_id, project_id, client_id,
                            cap_id, sample_type, sr_condition, receiving_remarks,
                            pickup_batch, pickup_datetime, lab_in_datetime,
                            lab_in_operator, transit_date, awb, tw_receiving,
                            created_date
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        sample_id,
                        lab_id,
                        project_id,
                        default_client_id,
                        mapped.get("cap_id"),
                        mapped.get("sample_type"),
                        mapped.get("sr_condition"),
                        mapped.get("receiving_remarks"),
                        mapped.get("pickup_batch"),
                        mapped.get("pickup_datetime"),
                        mapped.get("lab_in_datetime"),
                        mapped.get("lab_in_operator"),
                        mapped.get("transit_date"),
                        mapped.get("awb"),
                        mapped.get("tw_receiving"),
                        datetime.now().isoformat()
                    ))
                    stats["samples_imported"] += 1

                    # ── Insert sequencing run ─────────────────
                    run_id = mapped.get("run_id")
                    cursor.execute("""
                        INSERT INTO sequencing_runs (
                            sample_id, run_id, run_number,
                            novaseq_run_date, seq_start_datetime, seq_rdsc,
                            pipeline_id, index_plate, index_position, index1, index2,
                            run_name, order_operator, order_time,
                            process_time, process_operator, process_remarks,
                            sendout_datetime, sendout_operator,
                            repeat_start_point, days_of_life,
                            repeat_operator, repeat_datetime,
                            data_upload_date, notification_email_date,
                            sftp_deletion_date, final_deletion_date,
                            snpqc_status
                        ) VALUES (?, ?, 1, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        sample_id,
                        run_id,
                        mapped.get("novaseq_run_date"),
                        mapped.get("seq_start_datetime"),
                        mapped.get("seq_rdsc"),
                        mapped.get("pipeline_id"),
                        mapped.get("index_plate"),
                        mapped.get("index_position"),
                        mapped.get("index1"),
                        mapped.get("index2"),
                        mapped.get("run_name"),
                        mapped.get("order_operator"),
                        mapped.get("order_time"),
                        mapped.get("process_time"),
                        mapped.get("process_operator"),
                        mapped.get("process_remarks"),
                        mapped.get("sendout_datetime"),
                        mapped.get("sendout_operator"),
                        mapped.get("repeat_start_point"),
                        mapped.get("days_of_life"),
                        mapped.get("repeat_operator"),
                        mapped.get("repeat_datetime"),
                        mapped.get("data_upload_date"),
                        mapped.get("notification_email_date"),
                        mapped.get("sftp_deletion_date"),
                        mapped.get("final_deletion_date"),
                        mapped.get("snpqc_status"),
                    ))
                    run_record_id = cursor.lastrowid
                    stats["runs_imported"] += 1

                    # ── Insert QC metrics ─────────────────────
                    cursor.execute("""
                        INSERT INTO qc_metrics (
                            sample_id, run_id, run_record_id,
                            qc_conclusion, qc_failreason,
                            bases_all_trimmed, mapping_ratio, target_depth, base_Q30,
                            gc_rate, insert_size, dup_ratio, titv_ratio, variant_ratio,
                            hethom_ratio, snpindel_ratio, gender_analysed, gender_record,
                            contamination_level, ontarget_ratio,
                            bases_target_0x, bases_target_lt_20x, pcr_qc
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        sample_id, run_id, run_record_id,
                        mapped.get("qc_conclusion"), mapped.get("qc_failreason"),
                        safe_float(mapped.get("bases_all_trimmed")),
                        safe_float(mapped.get("mapping_ratio")),
                        safe_float(mapped.get("target_depth")),
                        safe_float(mapped.get("base_Q30")),
                        safe_float(mapped.get("gc_rate")),
                        safe_float(mapped.get("insert_size")),
                        safe_float(mapped.get("dup_ratio")),
                        safe_float(mapped.get("titv_ratio")),
                        safe_float(mapped.get("variant_ratio")),
                        safe_float(mapped.get("hethom_ratio")),
                        safe_float(mapped.get("snpindel_ratio")),
                        mapped.get("gender_analysed"), mapped.get("gender_record"),
                        safe_float(mapped.get("contamination_level")),
                        safe_float(mapped.get("ontarget_ratio")),
                        safe_float(mapped.get("bases_target_0x")),
                        safe_float(mapped.get("bases_target_lt_20x")),
                        mapped.get("pcr_qc"),
                    ))
                    stats["qc_imported"] += 1

                    # ── Insert run metrics ────────────────────
                    if run_id:
                        cursor.execute("""
                            INSERT OR IGNORE INTO run_metrics (
                                run_id, run_q30_pct, phix_error_rate,
                                barcode_contamination_rate, yield_gbp,
                                q30_pct_upload, depth_x_upload
                            ) VALUES (?, ?, ?, ?, ?, ?, ?)
                        """, (
                            run_id,
                            safe_float(mapped.get("run_q30_pct")),
                            safe_float(mapped.get("phix_error_rate")),
                            safe_float(mapped.get("barcode_contamination_rate")),
                            safe_float(mapped.get("yield_gbp")),
                            safe_float(mapped.get("q30_pct_upload")),
                            safe_float(mapped.get("depth_x_upload")),
                        ))

                    # ── Insert SNP markers ────────────────────
                    for marker in SNP_MARKER_COLS:
                        val = safe_str(row.get(marker))
                        if val:
                            cursor.execute("""
                                INSERT INTO snp_markers (sample_id, run_record_id, marker_name, marker_value)
                                VALUES (?, ?, ?, ?)
                            """, (sample_id, run_record_id, marker, val))
                            stats["snp_imported"] += 1

            except Exception as e:
                stats["errors"].append(f"Row {row_num} (LabID={lab_id}): {e}")

    if not dry_run:
        conn.execute("PRAGMA foreign_keys = ON")
        conn.commit()
    conn.close()

    # ── Print & save report ───────────────────────────────
    report = [
        "=" * 55,
        "  Migration Report",
        f"  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"  Mode: {'DRY RUN' if dry_run else 'LIVE'}",
        "=" * 55,
        f"  Total rows processed : {stats['total_rows']}",
        f"  Samples imported     : {stats['samples_imported']}",
        f"  Samples skipped      : {stats['samples_skipped']}",
        f"  Sequencing runs      : {stats['runs_imported']}",
        f"  QC records           : {stats['qc_imported']}",
        f"  SNP markers          : {stats['snp_imported']}",
        f"  Errors               : {len(stats['errors'])}",
    ]
    if stats["errors"]:
        report.append("\n  Error Details:")
        for e in stats["errors"]:
            report.append(f"    - {e}")

    report_text = "\n".join(report)
    print("\n" + report_text)

    with open(LOG_PATH, "w") as f:
        f.write(report_text)
    print(f"\n📄 Report saved to: {LOG_PATH}")


if __name__ == "__main__":
    import sys
    dry = "--dry-run" in sys.argv
    migrate(dry_run=dry)
