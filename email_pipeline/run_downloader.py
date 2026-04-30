"""
Entry point for the email automation pipeline.
1) Download report attachments via POP3
2) Send failure alert via SMTP if download is incomplete
"""
import os
import sys
import logging
from dotenv import load_dotenv

from email_pipeline import config
from email_pipeline.data_downloader import ReportEmailDownloader, DownloadResult
from email_pipeline.smtp_notifier import EmailNotifier


def setup_logging():
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s | %(levelname)-7s | %(name)s | %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S',
    )

def main(send_failure: bool = True) -> 'DownloadResult':
    setup_logging()
    log = logging.getLogger('run_email')

    load_dotenv(config.ENV_PATH)

    pop3_user = os.environ.get(config.POP3_USER_ENV)
    pop3_pass = os.environ.get(config.POP3_PASS_ENV)
    smtp_user = os.environ.get(config.SMTP_USER_ENV)
    smtp_pass = os.environ.get(config.SMTP_PASS_ENV)

    if not all([pop3_user, pop3_pass, smtp_user, smtp_pass]):
        log.error("Missing credentials in .env")
        return DownloadResult(success=False, error_message="Missing credentials")

    # Step 1: Download attachments
    downloader = ReportEmailDownloader(pop3_user, pop3_pass)
    result = downloader.download_reports()

    if result.success:
        log.info(f"✅ Downloaded {len(result.downloaded_files)} files")
        for f in result.downloaded_files:
            log.info(f"   - {f}")
        return result

    # Step 2: Send failure alert if download incomplete
    log.warning(f"❌ Download incomplete: {result.error_message}")
    if send_failure:
        notifier = EmailNotifier(smtp_user, smtp_pass)
        if notifier.send_failure_alert(received_times=result.received_times):
            log.info("Failure notification sent")

    return result

if __name__ == '__main__':
    sys.exit(main())
