-- ============================================================
-- Clinical Sample Testing Monitor - Master Database Schema
-- Version: 1.0
-- ============================================================

PRAGMA foreign_keys = ON;

-- ============================================================
-- TABLE 1: service_types
-- Flexible list of service types (WES, WGS, future types)
-- ============================================================
CREATE TABLE IF NOT EXISTS service_types (
    service_type_id     TEXT PRIMARY KEY,       -- e.g. 'WES', 'WGS'
    service_name        TEXT NOT NULL,
    description         TEXT,
    is_active           INTEGER DEFAULT 1,      -- 1=active, 0=retired
    created_date        TEXT DEFAULT (date('now'))
);

-- Seed default service types
INSERT OR IGNORE INTO service_types VALUES ('WES', 'Whole Exome Sequencing', 'Standard WES service', 1, date('now'));
INSERT OR IGNORE INTO service_types VALUES ('WGS', 'Whole Genome Sequencing', 'Standard WGS service', 1, date('now'));


-- ============================================================
-- TABLE 2: clients
-- One row per client, all billing config stored here
-- ============================================================
CREATE TABLE IF NOT EXISTS clients (
    client_id                   TEXT PRIMARY KEY,
    client_name                 TEXT NOT NULL,
    contact_email               TEXT,
    billing_currency            TEXT DEFAULT 'HKD',     -- 'HKD' or 'USD'
    billing_period_day_start    INTEGER DEFAULT 1,       -- Day of month billing starts
    custom_billing_period       TEXT,                    -- Override: 'YYYY-MM-DD to YYYY-MM-DD'
    requires_additional_qc      INTEGER DEFAULT 0,       -- 1=yes, 0=no
    sftp_path                   TEXT,
    data_confidentiality_level  TEXT DEFAULT 'general',  -- 'general' or 'confidential'
    lab_id_prefix               TEXT,                    -- Prefix used for LabID generation
    notes                       TEXT,
    is_active                   INTEGER DEFAULT 1,
    created_date                TEXT DEFAULT (date('now'))
);


-- ============================================================
-- TABLE 3: client_billing_rates
-- Per-client, per-service billing rates
-- Kept separate so rates can change over time
-- ============================================================
CREATE TABLE IF NOT EXISTS client_billing_rates (
    rate_id             INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id           TEXT NOT NULL REFERENCES clients(client_id),
    service_type_id     TEXT NOT NULL REFERENCES service_types(service_type_id),
    billing_mode        TEXT NOT NULL DEFAULT 'per_unit',
    -- billing_mode options:
    --   'per_unit'      = 1 sample = 1 billing unit (WES default)
    --   '3x_wes'        = 1 WGS = 3 WES billing units (legacy WGS)
    --   'independent'   = 1 WGS = 1 WGS billing unit (future WGS)
    unit_price          REAL NOT NULL DEFAULT 0,
    currency            TEXT DEFAULT 'HKD',
    effective_from      TEXT DEFAULT (date('now')),
    effective_to        TEXT,                            -- NULL = still active
    notes               TEXT
);


-- ============================================================
-- TABLE 4: projects
-- One row per project / batch
-- ============================================================
CREATE TABLE IF NOT EXISTS projects (
    project_id              TEXT PRIMARY KEY,
    client_id               TEXT NOT NULL REFERENCES clients(client_id),
    service_type_id         TEXT NOT NULL REFERENCES service_types(service_type_id),
    collection_batch        TEXT,
    custom_billing_start    TEXT,   -- Override billing period start date
    custom_billing_end      TEXT,   -- Override billing period end date
    submission_form_format  TEXT,   -- Which form template this client uses
    notes                   TEXT,
    created_date            TEXT DEFAULT (date('now')),
    is_active               INTEGER DEFAULT 1
);


-- ============================================================
-- TABLE 5: samples
-- One row per sample. SampleID = customer name. LabID = internal.
-- ============================================================
CREATE TABLE IF NOT EXISTS samples (
    sample_id               TEXT PRIMARY KEY,   -- Customer's original sample name
    lab_id                  TEXT UNIQUE,        -- Internal ID: prefix + sample_id
    project_id              TEXT NOT NULL REFERENCES projects(project_id),
    client_id               TEXT NOT NULL REFERENCES clients(client_id),
    service_type_id         TEXT REFERENCES service_types(service_type_id),
    sample_type             TEXT,               -- e.g. Blood, FFPE, DNA
    cap_id                  TEXT,               -- Capture group ID (8 samples per capture)
    sr_condition            TEXT,               -- Sample receiving condition
    receiving_remarks       TEXT,
    pickup_batch            TEXT,
    pickup_datetime         TEXT,               -- ISO datetime
    lab_in_datetime         TEXT,               -- Sample received by lab datetime
    lab_in_operator         TEXT,
    transit_date            TEXT,
    awb                     TEXT,               -- Airway bill number
    tw_receiving            TEXT,               -- TW receiving info
    recollection_of         TEXT REFERENCES samples(sample_id),  -- Links to original if re-collection
    submission_form_format  TEXT,               -- Format of submission form for this sample
    created_date            TEXT DEFAULT (datetime('now')),
    created_by              TEXT
);


-- ============================================================
-- TABLE 6: sample_status_tracking
-- Full audit trail of every status change per sample
-- ============================================================
CREATE TABLE IF NOT EXISTS sample_status_tracking (
    tracking_id     INTEGER PRIMARY KEY AUTOINCREMENT,
    sample_id       TEXT NOT NULL REFERENCES samples(sample_id),
    status          TEXT NOT NULL,
    -- Valid statuses:
    --  1.  sample_ordered
    --  2.  sample_arrived_registration
    --  3.  sample_received_by_lab
    --  4.  qc_passed_email_sent
    --  5.  processing_started
    --  6.  processing_finished
    --  7.  sequencing_in_progress
    --  8.  sequencing_finished
    --  9.  data_analysis_finished
    --  10. additional_qc_in_progress    (optional, per client setting)
    --  11. additional_qc_passed         (optional, per client setting)
    --  12. data_ready_for_review
    --  13. data_uploaded_to_sftp
    --  14. sftp_deletion_date_set
    --  15. final_deletion_date_set
    --  16. notification_email_sent
    --  17. billing_email_sent
    timestamp       TEXT DEFAULT (datetime('now')),
    operator        TEXT,
    remarks         TEXT
);


-- ============================================================
-- TABLE 7: sequencing_runs
-- Tracks all sequencing runs per sample (including repeats)
-- ============================================================
CREATE TABLE IF NOT EXISTS sequencing_runs (
    run_record_id           INTEGER PRIMARY KEY AUTOINCREMENT,
    sample_id               TEXT NOT NULL REFERENCES samples(sample_id),
    lab_project_id          TEXT,               -- Internal lab project ID (may differ from project_id)
    sequencing_project_id   TEXT,               -- Sequencer project ID (may differ from lab_project_id)
    run_id                  TEXT,               -- Sequencing run ID
    run_number              INTEGER DEFAULT 1,  -- 1=original, 2,3...=repeats
    novaseq_run_date        TEXT,
    seq_start_datetime      TEXT,
    seq_rdsc                TEXT,
    pipeline_id             TEXT,
    index_plate             TEXT,
    index_position          TEXT,
    index1                  TEXT,
    index2                  TEXT,
    run_name                TEXT,
    order_operator          TEXT,
    order_time              TEXT,
    process_time            TEXT,
    process_operator        TEXT,
    process_remarks         TEXT,
    sendout_datetime        TEXT,
    sendout_operator        TEXT,
    repeat_start_point      TEXT,
    days_of_life            TEXT,
    repeat_operator         TEXT,
    repeat_datetime         TEXT,
    data_upload_date        TEXT,
    notification_email_date TEXT,
    sftp_deletion_date      TEXT,
    final_deletion_date     TEXT,
    snpqc_status            TEXT
);


-- ============================================================
-- TABLE 8: qc_metrics
-- QC results per sample per run
-- ============================================================
CREATE TABLE IF NOT EXISTS qc_metrics (
    qc_id                   INTEGER PRIMARY KEY AUTOINCREMENT,
    sample_id               TEXT NOT NULL REFERENCES samples(sample_id),
    run_id                  TEXT,
    run_record_id           INTEGER REFERENCES sequencing_runs(run_record_id),
    qc_conclusion           TEXT,               -- 'PASS' or 'FAIL'
    qc_failreason           TEXT,
    bases_all_trimmed       REAL,
    mapping_ratio           REAL,
    target_depth            REAL,
    base_Q30                REAL,
    gc_rate                 REAL,
    insert_size             REAL,
    dup_ratio               REAL,
    titv_ratio              REAL,
    variant_ratio           REAL,
    hethom_ratio            REAL,
    snpindel_ratio          REAL,
    gender_analysed         TEXT,
    gender_record           TEXT,
    contamination_level     REAL,
    ontarget_ratio          REAL,
    bases_target_0x         REAL,
    bases_target_lt_20x     REAL,
    pcr_qc                  TEXT,
    created_date            TEXT DEFAULT (datetime('now'))
);


-- ============================================================
-- TABLE 9: snp_markers
-- Flexible SNP marker results (rs IDs) per sample per run
-- Stored as key-value so new markers can be added anytime
-- ============================================================
CREATE TABLE IF NOT EXISTS snp_markers (
    marker_id       INTEGER PRIMARY KEY AUTOINCREMENT,
    sample_id       TEXT NOT NULL REFERENCES samples(sample_id),
    run_record_id   INTEGER REFERENCES sequencing_runs(run_record_id),
    marker_name     TEXT NOT NULL,   -- e.g. 'rs1042713_wes', 'rs1042713_pcr'
    marker_value    TEXT,
    created_date    TEXT DEFAULT (datetime('now'))
);


-- ============================================================
-- TABLE 10: run_metrics
-- Per sequencing run statistics
-- ============================================================
CREATE TABLE IF NOT EXISTS run_metrics (
    metric_id                   INTEGER PRIMARY KEY AUTOINCREMENT,
    run_id                      TEXT NOT NULL,
    run_q30_pct                 REAL,
    phix_error_rate             REAL,
    barcode_contamination_rate  REAL,
    yield_gbp                   REAL,
    q30_pct_upload              REAL,
    depth_x_upload              REAL,
    created_date                TEXT DEFAULT (datetime('now'))
);


-- ============================================================
-- TABLE 11: billing
-- One row per billable sample per billing period
-- ============================================================
CREATE TABLE IF NOT EXISTS billing (
    billing_id              INTEGER PRIMARY KEY AUTOINCREMENT,
    sample_id               TEXT NOT NULL REFERENCES samples(sample_id),
    client_id               TEXT NOT NULL REFERENCES clients(client_id),
    service_type_id         TEXT NOT NULL REFERENCES service_types(service_type_id),
    billing_period_start    TEXT NOT NULL,      -- YYYY-MM-DD
    billing_period_end      TEXT NOT NULL,      -- YYYY-MM-DD
    notification_email_date TEXT,               -- Cut-off reference date
    billing_mode            TEXT,               -- 'per_unit', '3x_wes', 'independent'
    billing_units           REAL,               -- WES=1, WGS=3 or 1
    unit_price              REAL,
    currency                TEXT DEFAULT 'HKD',
    total_amount            REAL,
    pickup_date             TEXT,               -- For PDF report col C
    lab_in_date             TEXT,               -- For PDF report col D
    billing_email_sent      INTEGER DEFAULT 0,
    billing_email_date      TEXT,
    invoice_pdf_path        TEXT,
    pdf_row_label           TEXT,               -- e.g. 'GSampleID_1', 'GSampleID_2'
    remarks                 TEXT,
    created_date            TEXT DEFAULT (datetime('now'))
);


-- ============================================================
-- INDEXES for fast lookups
-- ============================================================
CREATE INDEX IF NOT EXISTS idx_samples_project    ON samples(project_id);
CREATE INDEX IF NOT EXISTS idx_samples_client     ON samples(client_id);
CREATE INDEX IF NOT EXISTS idx_samples_lab_id     ON samples(lab_id);
CREATE INDEX IF NOT EXISTS idx_status_sample      ON sample_status_tracking(sample_id);
CREATE INDEX IF NOT EXISTS idx_status_status      ON sample_status_tracking(status);
CREATE INDEX IF NOT EXISTS idx_runs_sample        ON sequencing_runs(sample_id);
CREATE INDEX IF NOT EXISTS idx_qc_sample          ON qc_metrics(sample_id);
CREATE INDEX IF NOT EXISTS idx_billing_client     ON billing(client_id);
CREATE INDEX IF NOT EXISTS idx_billing_period     ON billing(billing_period_start, billing_period_end);
CREATE INDEX IF NOT EXISTS idx_snp_sample         ON snp_markers(sample_id);
