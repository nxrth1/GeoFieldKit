"""
SCRIPT 7 — AUTOMATION PIPELINE
================================================
What it does:
  - Watches a folder for new ODK export files
  - When a new file drops in, auto-runs the full pipeline:
      Script 1 (Clean) → Script 2 (Attendance) → Script 3 (Payroll)
      → Script 4 (GPS QC) → Script 5 (Report) → Script 6 (Maps)
  - Sends an email with the report attached when done
  - Logs every run with timestamps
  - Can also be run manually as a one-shot pipeline

USAGE:
  # Run pipeline ONCE on an existing file:
      python script_7_automation.py --file ODK_Raw_Export.xlsx

  # Watch a folder and auto-run on any new .xlsx dropped in:
      python script_7_automation.py --watch ./incoming

  # Dry run (test without emailing):
      python script_7_automation.py --file ODK_Raw_Export.xlsx --no-email

Requirements:
  pip install watchdog schedule
"""

import os
import sys
import time
import shutil
import logging
import argparse
import subprocess
from datetime import datetime
from pathlib import Path

# ── CONFIG — EDIT THESE ──────────────────────────────────────────────────
WATCH_FOLDER   = "./incoming"          # folder to watch for new ODK exports
OUTPUT_FOLDER  = "./outputs"           # where processed outputs go
LOG_FILE       = "pipeline.log"
SCRIPTS = [
    "script_1_cleaning.py",
    "script_2_attendance.py",
    "script_3_payroll.py",
    "script_4_gps_qc.py",
    "script_5_reporting.py",
    "script_6_geospatial.py",
]

# Email config (only needed if sending emails)
EMAIL_CONFIG = {
    "smtp_host"    : "smtp.gmail.com",
    "smtp_port"    : 587,
    "sender_email" : "your_email@gmail.com",
    "sender_pass"  : "your_app_password",       # Use Gmail App Password (not regular password)
    "recipients"   : ["supervisor@example.com", "manager@example.com"],
    "subject"      : "📊 ODK Field Data Report — {date}",
}

# Files to attach in the email
EMAIL_ATTACHMENTS = [
    "cleaned_data.xlsx",
    "attendance_report.xlsx",
    "payroll_report.xlsx",
    "quality_control_report.xlsx",
    "full_report.xlsx",
]

# ── LOGGING SETUP ────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  [%(levelname)s]  %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE),
        logging.StreamHandler(sys.stdout),
    ]
)
log = logging.getLogger(__name__)

# ── PIPELINE RUNNER ──────────────────────────────────────────────────────
def run_pipeline(input_file: str, send_email: bool = True):
    """Run all scripts in sequence on the given input file."""
    start_time = datetime.now()
    log.info(f"{'='*60}")
    log.info(f"[START] PIPELINE STARTED — Input: {input_file}")
    log.info(f"{'='*60}")

    # Copy input file to working directory as the standard name
    target = "ODK_Raw_Export.xlsx"
    if os.path.abspath(input_file) != os.path.abspath(target):
        shutil.copy(input_file, target)
        log.info(f"📋 Copied {input_file} → {target}")

    # Run each script
    results = {}
    for script in SCRIPTS:
        if not os.path.exists(script):
            log.warning(f"[WARN]  {script} not found — skipping")
            results[script] = "SKIPPED"
            continue

        log.info(f"\n▶️  Running {script}...")
        try:
            result = subprocess.run(
                [sys.executable, script],
                capture_output=True, text=True, timeout=300
            )
            if result.returncode == 0:
                log.info(f"   [DONE] {script} completed")
                if result.stdout:
                    for line in result.stdout.strip().split("\n"):
                        log.info(f"      {line}")
                results[script] = "SUCCESS"
            else:
                log.error(f"   ❌ {script} FAILED")
                log.error(f"      {result.stderr[-500:]}")
                results[script] = "FAILED"
        except subprocess.TimeoutExpired:
            log.error(f"   ⏱️  {script} TIMED OUT after 5 minutes")
            results[script] = "TIMEOUT"
        except Exception as e:
            log.error(f"   💥 {script} ERROR: {e}")
            results[script] = f"ERROR: {e}"

    # Move outputs to output folder
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    output_files = []
    for f in EMAIL_ATTACHMENTS + ["points_export.geojson"]:
        if os.path.exists(f):
            dest = os.path.join(OUTPUT_FOLDER, f)
            shutil.copy(f, dest)
            output_files.append(dest)
            log.info(f"📁 Saved output: {dest}")

    # Summary
    elapsed = (datetime.now() - start_time).seconds
    successes = sum(1 for v in results.values() if v == "SUCCESS")
    failures  = sum(1 for v in results.values() if "FAIL" in str(v) or "ERROR" in str(v))

    log.info(f"\n{'='*60}")
    log.info(f"📊 PIPELINE COMPLETE in {elapsed}s")
    log.info(f"   Scripts run: {len(SCRIPTS)}")
    log.info(f"   [DONE] Success: {successes}  |  ❌ Failed: {failures}")
    log.info(f"{'='*60}\n")

    # Send email
    if send_email and failures == 0:
        send_report_email(output_files, results, elapsed)
    elif failures > 0:
        log.warning("[WARN]  Skipping email — pipeline had failures")

    return results

# ── EMAIL SENDER ─────────────────────────────────────────────────────────
def send_report_email(output_files: list, results: dict, elapsed_sec: int):
    """Send email with reports attached."""
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text      import MIMEText
    from email.mime.base      import MIMEBase
    from email import encoders

    log.info("📧 Sending email report...")
    try:
        msg = MIMEMultipart()
        msg["From"]    = EMAIL_CONFIG["sender_email"]
        msg["To"]      = ", ".join(EMAIL_CONFIG["recipients"])
        msg["Subject"] = EMAIL_CONFIG["subject"].format(
            date=datetime.now().strftime("%Y-%m-%d %H:%M"))

        # Body
        body = f"""
Dear Team,

The automated ODK field data pipeline has completed successfully.

📊 Summary:
  • Run time: {elapsed_sec} seconds
  • Date: {datetime.now().strftime('%Y-%m-%d %H:%M')}
  • Status: {'[DONE] All scripts passed' if all(v=='SUCCESS' for v in results.values()) else '[WARN] Some scripts had issues'}

📎 Reports attached:
  • cleaned_data.xlsx — clean ODK data
  • attendance_report.xlsx — clock-in/out analysis
  • payroll_report.xlsx — pay calculations
  • quality_control_report.xlsx — GPS QC (PASS/REVIEW/FAIL)
  • full_report.xlsx — executive summary with charts

Please review and let me know if you have any questions.

This is an automated message. Do not reply to this email.
        """
        msg.attach(MIMEText(body, "plain"))

        # Attach files
        for filepath in output_files:
            if os.path.exists(filepath):
                with open(filepath, "rb") as f:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header("Content-Disposition",
                                f"attachment; filename={os.path.basename(filepath)}")
                msg.attach(part)

        # Send
        with smtplib.SMTP(EMAIL_CONFIG["smtp_host"], EMAIL_CONFIG["smtp_port"]) as server:
            server.starttls()
            server.login(EMAIL_CONFIG["sender_email"], EMAIL_CONFIG["sender_pass"])
            server.sendmail(
                EMAIL_CONFIG["sender_email"],
                EMAIL_CONFIG["recipients"],
                msg.as_string()
            )
        log.info(f"   [DONE] Email sent to: {', '.join(EMAIL_CONFIG['recipients'])}")

    except Exception as e:
        log.error(f"   ❌ Email failed: {e}")
        log.error("   💡 Check your EMAIL_CONFIG settings and ensure Gmail App Password is set")

# ── FOLDER WATCHER ───────────────────────────────────────────────────────
def watch_folder(watch_path: str, send_email: bool = True):
    """Watch a folder and run the pipeline when a new .xlsx file appears."""
    try:
        from watchdog.observers import Observer
        from watchdog.events    import FileSystemEventHandler
    except ImportError:
        log.error("watchdog not installed. Run: pip install watchdog")
        sys.exit(1)

    processed = set()

    class ODKHandler(FileSystemEventHandler):
        def on_created(self, event):
            if event.is_directory:
                return
            path = event.src_path
            if path.endswith(".xlsx") and path not in processed:
                processed.add(path)
                log.info(f"📥 New file detected: {path}")
                time.sleep(2)  # Wait for file to finish writing
                run_pipeline(path, send_email=send_email)

    os.makedirs(watch_path, exist_ok=True)
    log.info(f"👀 Watching folder: {os.path.abspath(watch_path)}")
    log.info(f"   Drop any ODK .xlsx export into this folder to trigger the pipeline")
    log.info(f"   Press Ctrl+C to stop\n")

    observer = Observer()
    observer.schedule(ODKHandler(), watch_path, recursive=False)
    observer.start()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
        log.info("🛑 Watcher stopped")
    observer.join()

# ── MAIN ─────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="ODK Field Data Automation Pipeline"
    )
    parser.add_argument("--file",     type=str, help="Run pipeline on a specific file")
    parser.add_argument("--watch",    type=str, help="Watch a folder for new files", nargs="?", const=WATCH_FOLDER)
    parser.add_argument("--no-email", action="store_true", help="Skip sending email")
    args = parser.parse_args()

    send_email = not args.no_email

    if args.file:
        if not os.path.exists(args.file):
            log.error(f"File not found: {args.file}")
            sys.exit(1)
        run_pipeline(args.file, send_email=send_email)

    elif args.watch:
        watch_folder(args.watch, send_email=send_email)

    else:
        # Default: run on ODK_Raw_Export.xlsx if it exists
        default_file = "ODK_Raw_Export.xlsx"
        if os.path.exists(default_file):
            log.info(f"No arguments given — running on {default_file}")
            run_pipeline(default_file, send_email=send_email)
        else:
            print("\nUSAGE:")
            print("  python script_7_automation.py --file  ODK_Raw_Export.xlsx")
            print("  python script_7_automation.py --watch ./incoming")
            print("  python script_7_automation.py --file  ODK_Raw_Export.xlsx --no-email\n")
            parser.print_help()
