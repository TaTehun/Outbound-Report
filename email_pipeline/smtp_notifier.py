"""
SMTP email notification module.
"""
import os
import smtplib
import logging
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.utils import formatdate
from email import encoders
from pathlib import Path
from typing import List, Optional
from dotenv import load_dotenv

from email_pipeline import config
from utils.date_utils import today_str
from utils import pivot_utils

logger = logging.getLogger(__name__)


class EmailNotifier:
    """SMTP email sender."""

    def __init__(self, username: str, password: str):
        self.username = username
        self.password = password

    def send(
        self,
        subject: str,
        body: str,
        to: List[str],
        cc: List[str] = None,
        sender: str = None,
        attachment_path: Optional[Path] = None,
    ) -> bool:
        """Send an email. Returns True on success."""
        cc = cc or []
        sender = sender or config.NOTIFICATION_FROM

        # MIME structure: related > alternative > html
        msgRoot = MIMEMultipart('related')
        msgRoot['From'] = sender
        msgRoot['To'] = ','.join(to)
        msgRoot['Cc'] = ','.join(cc)
        msgRoot['Date'] = formatdate(localtime=True)
        msgRoot['Subject'] = subject

        msgAlternative = MIMEMultipart('alternative')
        msgRoot.attach(msgAlternative)
        msgAlternative.attach(MIMEText(body, 'html', 'utf-8'))

        if attachment_path and Path(attachment_path).exists():
            with open(attachment_path, 'rb') as f:
                # Explicit xlsb MIME type so clients recognize the extension correctly
                part = MIMEBase('application', 'vnd.ms-excel.sheet.binary.macroEnabled.12')
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header(
                'Content-Disposition',
                'attachment',
                filename=Path(attachment_path).name,
            )
            msgRoot.attach(part)

        recipients = to + cc

        try:
            with smtplib.SMTP(config.SMTP_SERVER, config.SMTP_PORT) as smtp:
                smtp.login(self.username, self.password)
                smtp.sendmail(sender, recipients, msgRoot.as_string())
            logger.info(f"Mail sent to {recipients}")
            return True
        except Exception as e:
            logger.error(f"Failed to send mail: {e}", exc_info=True)
            return False

    def send_failure_alert(self, received_times: dict = None) -> bool:
        """Send failure alert with missing/received file details."""
        received_times = received_times or {}
        all_targets = [t["name"] for t in config.EMAIL_TARGETS]
        missing  = [name for name in all_targets if name not in received_times]
        received = [name for name in all_targets if name in received_times]

        lines = ["Dear All,", ""]
        if missing:
            lines.append("Today, we did not receive the file for")
            for name in missing:
                lines.append(f"ㆍ {name}")
            lines.append("")
        if received:
            lines.append("The following files were received today:")
            for name in received:
                lines.append(f"ㆍ {name} (Received Date: {received_times[name]})")
            lines.append("")
        lines.append("Thank you")

        body = (
            '<BODY style="font-size:11pt;font-family:Calibri">'
            + "<br>".join(lines)
            + "</BODY>"
        )
        return self.send(
            subject=config.FAILURE_SUBJECT,
            body=body,
            to=config.NOTIFICATION_TO,
            cc=config.NOTIFICATION_CC,
        )


def send_email(output_file: Path, received_times: dict, is_updated: bool = False, pivot_data: list = None) -> bool:
    """
    Send the consolidated report email with the xlsb file attached.

    output_file    : xlsb file path returned by combine_excel()
    received_times : {target name → CST received time} from run_downloader()
    is_updated     : if True, prepends "(Updated) " to the subject
    pivot_data     : pivot table data to embed as HTML table in the email body
    """
    load_dotenv(config.ENV_PATH)
    smtp_user = os.environ.get(config.SMTP_USER_ENV)
    smtp_pass = os.environ.get(config.SMTP_PASS_ENV)

    if not smtp_user or not smtp_pass:
        logger.error("SMTP credentials missing — email not sent")
        return False

    pivot_html = pivot_utils.to_html(pivot_data) if pivot_data else ""

    body = (
        '<BODY style="font-size:11pt;font-family:Calibri">'
        + "Dear All,<br><br>"
        + f"Please find the consolidated Daily Outbound Report for {today_str('h')} Attached.<br><br>"
        + pivot_html
        + "<br>Thank you"
        + "</BODY>"
    )
    subject = config.REPORT_SUBJECT.format(today_str('e'))
    if is_updated:
        subject = "(Updated) " + subject

    notifier = EmailNotifier(smtp_user, smtp_pass)
    success = notifier.send(
        subject=subject,
        body=body,
        to=config.REPORT_TO,
        cc=config.REPORT_CC,
        attachment_path=output_file,
    )
    logger.info(f"Attaching: {output_file} | exists: {output_file.exists()}")
    return success

def send_skip_email(downloaded_files: list) -> None:
    """Send notification that 2nd run files matched 1st run — report not resent."""
    load_dotenv(config.ENV_PATH)
    smtp_user = os.environ.get(config.SMTP_USER_ENV)
    smtp_pass = os.environ.get(config.SMTP_PASS_ENV)

    if not smtp_user or not smtp_pass:
        logger.error("SMTP credentials missing — skip email not sent")
        return

    file_names = [Path(f).name for f in downloaded_files]
    lines = (
        ["Dear All,", ""]
        + ["The following files from the second run were identical to the first run.",
           "No updated report has been sent.", ""]
        + [f"ㆍ {name}" for name in file_names]
        + ["", "Thank you"]
    )

    body = (
        '<BODY style="font-size:11pt;font-family:Calibri">'
        + "<br>".join(lines)
        + "</BODY>"
    )
    notifier = EmailNotifier(smtp_user, smtp_pass)
    notifier.send(
        subject=f"[No Update] Consolidated Manifest – {today_str('e')}",
        body=body,
        to=config.NOTIFICATION_TO,
        cc=config.NOTIFICATION_CC,
    )
