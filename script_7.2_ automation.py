"""
SCRIPT 7 — AUTOMATION PIPELINE (Robust Version)
================================================
- Watches a folder for new ODK export files
- Runs the full pipeline for each new file:
    Script 1 → Script 2 → Script 3 → Script 4 → Script 5 → Script 6
- Saves outputs with unique timestamps
- Sends email reports even if some scripts fail
- Logs every run
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

# ── FORCE UTF-8 OUTPUT ─────────────────────────────
try:
    if sys.stdout.encoding.lower() != "utf-8":
        sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

# ── CONFIG ─────────────────────────────────────────
WATCH_FOLDER   = "./incoming"
OUTPUT_FOLDER  = "./outputs"
LOG_FILE       = "pipeline.log"

SCRIPTS = [
    "script_1_cleaning.py",
    "script_2_attendance.py",
    "script_3_payroll.py",
    "script_4_gps_qc.py",
    "script_5_reporting.py",
    "script_6_geospatial.py",
]

EMAIL_CONFIG = {
    "smtp_host"    : "smtp.gmail.com",
    "smtp_port"    : 587,
    "sender_email" : os.environ.get("PIPELINE_EMAIL") or "your_email@gmail.com",
    "sender_pass"  : os.environ.get("PIPELINE_EMAIL_PASS") or "your_app_password",
    "recipients"   : ["supervisor@example.com", "manager@example.com"],
    "subject"      : "ODK Field Data Report — {date}",
}

EMAIL_ATTACHMENTS = [
    "cleaned_data.xlsx",
    "attendance_report.xlsx",
    "payroll_report.xlsx",
    "quality_control_report.xlsx",
    "full_report.xlsx",
]

# ── LOGGING SETUP ────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  [%(levelname)s]  %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ]
)
log = logging.getLogger(__name__)

# ── PIPELINE RUNNER ──────────────────────────────
def run_pipeline(input_file: str, send_email: bool = True):
    start_time = datetime.now()
    log.info("=" * 60)
    log.info(f"PIPELINE STARTED — Input: {input_file}")
    log.info("=" * 60)

    # Create a unique timestamped target
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    target = f"ODK_Raw_Export_{timestamp}.xlsx"
    shutil.copy(input_file, target)
    log.info(f"Copied {input_file} → {target}")

    results = {}

    for script in SCRIPTS:
        if not os.path.exists(script):
            log.warning(f"{script} not found — skipping")
            results[script] = "SKIPPED"
            continue

        log.info(f"Running {script}...")
        try:
            result = subprocess.run(
                [sys.executable, script],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=300
            )

            if result.returncode == 0:
                log.info(f"{script} completed successfully")
                results[script] = "SUCCESS"
            else:
                log.error(f"{script} FAILED")
                log.error(result.stderr)
                results[script] = f"FAILED: {result.stderr.strip()[:1000]}"

        except subprocess.TimeoutExpired:
            log.error(f"{script} TIMED OUT after 5 minutes")
            results[script] = "TIMEOUT"

        except Exception as e:
            log.error(f"{script} ERROR: {e}")
            results[script] = f"ERROR: {e}"

    # ── Move Outputs ─────────────────────────────
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    output_files = []

    for f in EMAIL_ATTACHMENTS + ["points_export.geojson"]:
        if os.path.exists(f):
            dest_name = f"{Path(f).stem}_{timestamp}{Path(f).suffix}"
            dest = os.path.join(OUTPUT_FOLDER, dest_name)
            shutil.copy(f, dest)
            output_files.append(dest)
            log.info(f"Saved output: {dest}")

    # ── Summary ─────────────────────────────────
    elapsed = (datetime.now() - start_time).seconds
    successes = sum(1 for v in results.values() if v == "SUCCESS")
    failures = len(results) - successes

    log.info("=" * 60)
    log.info(f"PIPELINE COMPLETE in {elapsed}s")
    log.info(f"Scripts run: {len(SCRIPTS)} | Success: {successes} | Failed: {failures}")
    log.info("=" * 60)

    if send_email:
        send_report_email(output_files, results, elapsed)

    return results

# ── EMAIL SENDER ────────────────────────────────
def send_report_email(output_files: list, results: dict, elapsed_sec: int):
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase
    from email import encoders

    log.info("Sending email report...")

    try:
        msg = MIMEMultipart()
        msg["From"] = EMAIL_CONFIG["sender_email"]
        msg["To"] = ", ".join(EMAIL_CONFIG["recipients"])
        msg["Subject"] = EMAIL_CONFIG["subject"].format(
            date=datetime.now().strftime("%Y-%m-%d %H:%M")
        )

        body = f"""
Dear Team,

The automated ODK field data pipeline has completed.

Run time: {elapsed_sec} seconds
Date: {datetime.now().strftime('%Y-%m-%d %H:%M')}

Script results:
"""
        for s, r in results.items():
            body += f"- {s}: {r}\n"

        body += "\nReports attached where available.\n\nThis is an automated message."

        msg.attach(MIMEText(body, "plain"))

        for filepath in output_files:
            if os.path.exists(filepath):
                with open(filepath, "rb") as f:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header(
                    "Content-Disposition",
                    f"attachment; filename={os.path.basename(filepath)}"
                )
                msg.attach(part)

        with smtplib.SMTP(EMAIL_CONFIG["smtp_host"], EMAIL_CONFIG["smtp_port"]) as server:
            server.starttls()
            server.login(
                EMAIL_CONFIG["sender_email"],
                EMAIL_CONFIG["sender_pass"]
            )
            server.sendmail(
                EMAIL_CONFIG["sender_email"],
                EMAIL_CONFIG["recipients"],
                msg.as_string()
            )

        log.info("Email sent successfully")

    except Exception as e:
        log.error(f"Email failed: {e}")

# ── FOLDER WATCHER ─────────────────────────────
def watch_folder(watch_path: str, send_email: bool = True):
    try:
        from watchdog.observers import Observer
        from watchdog.events import FileSystemEventHandler
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
                # Wait until file is stable
                stable = False
                prev_size = -1
                while not stable:
                    try:
                        curr_size = os.path.getsize(path)
                        if curr_size == prev_size:
                            stable = True
                        else:
                            prev_size = curr_size
                            time.sleep(1)
                    except FileNotFoundError:
                        return

                processed.add(path)
                log.info(f"New file detected: {path}")
                run_pipeline(path, send_email=send_email)

    os.makedirs(watch_path, exist_ok=True)
    log.info(f"Watching folder: {os.path.abspath(watch_path)}")

    observer = Observer()
    observer.schedule(ODKHandler(), watch_path, recursive=False)
    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
        log.info("Watcher stopped")
    observer.join()

# ── MAIN ──────────────────────────────────────
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="ODK Field Data Automation Pipeline")
    parser.add_argument("--file", type=str, help="Run pipeline on specific file")
    parser.add_argument("--watch", type=str, help="Watch folder", nargs="?", const=WATCH_FOLDER)
    parser.add_argument("--no-email", action="store_true")
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
        default_file = "ODK_Raw_Export.xlsx"
        if os.path.exists(default_file):
            run_pipeline(default_file, send_email=send_email)
        else:
            parser.print_help()