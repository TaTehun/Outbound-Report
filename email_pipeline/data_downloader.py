"""
Downloads report attachments from a POP3 mail server.
"""
import poplib
import logging
import pytz

from email.parser import Parser
from dataclasses import dataclass, field
from typing import List, Optional

from email_pipeline import config
from email.utils import parsedate_to_datetime
from datetime import datetime, timedelta

logger = logging.getLogger(__name__)
texas_tz = pytz.timezone('America/Chicago')


@dataclass
class DownloadResult:
    """Result of a download run."""
    success: bool
    downloaded_files: List[str] = field(default_factory=list)
    received_times: dict = field(default_factory=dict)  # {target name → CST received time}
    error_message: Optional[str] = None


class ReportEmailDownloader:
    """
    Connects to a POP3 server and downloads attachments matching configured targets.

    Usage:
        downloader = ReportEmailDownloader(user, password)
        result = downloader.download_reports()
        if result.success:
            print(result.downloaded_files)
    """

    def __init__(self, user_email: str, user_password: str):
        self.user_email = user_email
        self.user_password = user_password
        self.connection: Optional[poplib.POP3_SSL] = None
        self.downloaded_files: List[str] = []

    # ------------------------------------------------------------
    # Connection Management
    # ------------------------------------------------------------
    def __enter__(self):
        self.connect()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    def connect(self) -> None:
        """Connect and login to the POP3 server via SSL."""
        if self.connection is not None:
            return

        logger.info(f"Connecting to {config.POP3_SERVER}:{config.POP3_PORT}")
        self.connection = poplib.POP3_SSL(config.POP3_SERVER, config.POP3_PORT)

        welcome = self.connection.getwelcome()
        logger.debug(f"Server welcome: {welcome}")

        self.connection.user(self.user_email)
        self.connection.pass_(self.user_password)
        logger.info("POP3 login successful")

    def close(self) -> None:
        """Close the POP3 connection."""
        if self.connection is not None:
            try:
                self.connection.quit()
                logger.info("POP3 connection closed")
            except Exception as e:
                logger.warning(f"Error during quit: {e}")
            finally:
                self.connection = None

    # ------------------------------------------------------------
    # Entry Point
    # ------------------------------------------------------------
    def download_reports(self) -> DownloadResult:
        """Download attachments matching EMAIL_TARGETS. Stops early when all found."""
        try:
            self._clean_data_folder()
            self.connect()

            count, size = self.connection.stat()
            logger.info(f"Mailbox: {count} messages, {size} bytes")

            return self._search_and_download()

        except Exception as e:
            logger.error(f"Download failed: {e}", exc_info=True)
            return DownloadResult(success=False, error_message=str(e))
        finally:
            self.close()

    # ------------------------------------------------------------
    # Internal Helpers
    # ------------------------------------------------------------
    def _clean_data_folder(self) -> None:
        """Create today's date folders (no cleanup needed with date-based folder structure)."""
        config.DATA_DIR.mkdir(parents=True, exist_ok=True)
        config.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

        # for f in config.DATA_DIR.iterdir():
        #     if not f.name.startswith(config.MASTER_FILE_PREFIX):
        #         logger.info(f"Removing old file: {f}")
        #         f.unlink()

        # for f in config.OUTPUT_DIR.iterdir():
        #     if f.is_file():
        #         logger.info(f"Removing old output: {f}")
        #         f.unlink()

    def _search_and_download(self) -> DownloadResult:
        """Scan emails newest-first and download matching attachments."""
        _, mails_list, _ = self.connection.list()
        recent_mails = list(reversed(mails_list[-config.EMAIL_SEARCH_COUNT:]))

        matched_contains: set = set()  # tracks completed targets (list → tuple for set compatibility)
        received_times: dict = {}

        today = datetime.now(texas_tz).date()

        for mail in recent_mails:
            email_index = mail.split()[0].decode() if isinstance(mail, bytes) \
                else str(mail)[2:].split(' ')[0].strip()

            # Fetch headers only first — stop if we've passed today's emails
            headers = self._fetch_headers(email_index)
            try:
                email_date = parsedate_to_datetime(headers.get('Date', '')).astimezone(texas_tz).date()
                if email_date < today:
                    logger.info(f"Reached mail older than today at index {email_index}, stopping search")
                    break
            except Exception:
                pass  # If date parse fails, continue scanning

            matched = self._matches_criteria(headers, exclude=matched_contains)

            if matched:
                # Full fetch only for matched emails
                msg = self._fetch_email(email_index)
                email_date = parsedate_to_datetime(msg.get('Date', '')).astimezone(texas_tz)
                logger.info(f"Match found at index {email_index} | Date: {email_date.strftime('%Y-%m-%d %H:%M %Z')} | Subject: {msg.get('Subject', '')}")
                self._save_attachments(msg)

                target_name = matched.get("name", str(matched["contains"]))
                received_times[target_name] = email_date.strftime('%m/%d/%Y %H:%M CST')

                matched_contains.add(tuple(matched["contains"]))

                if self._is_complete(matched_contains):
                    logger.info("All expected attachments downloaded")
                    return DownloadResult(
                        success=True,
                        downloaded_files=self.downloaded_files,
                        received_times=received_times,
                    )

        # Partial result — include whatever was received
        return DownloadResult(
            success=False,
            downloaded_files=self.downloaded_files,
            received_times=received_times,
            error_message=f"Expected {config.EXPECTED_ATTACHMENT_COUNT} files, got {len(self.downloaded_files)}"
        )

    def _fetch_headers(self, email_index: str):
        """Fetch and parse headers only (fast, used for subject/sender/date check)."""
        _, lines, _ = self.connection.top(email_index, 0)  # 0 body lines = headers only
        msg_content = b'\r\n'.join(lines).decode('utf-8', 'replace')
        return Parser().parsestr(msg_content)

    def _fetch_email(self, email_index: str):
        """Fetch and parse the full email (used for attachment download)."""
        _, lines, _ = self.connection.retr(email_index)
        msg_content = b'\r\n'.join(lines).decode('utf-8', 'replace')
        return Parser().parsestr(msg_content)

    def _matches_criteria(self, msg, exclude: set = None) -> dict | None:
        """
        Check if an email matches any EMAIL_TARGET.
        Returns the matched target dict, or None if no match.
        exclude: set of already-downloaded targets (as tuples) to skip.
        """
        from_addr = msg.get('From', '')
        subject = msg.get('Subject', '').strip()

        try:
            email_date = parsedate_to_datetime(msg.get('Date', '')).astimezone(texas_tz).date()
            """
            # 오늘 날짜 이메일만 처리
            if email_date != datetime.now(texas_tz).date():
                return None
            """
            # 5일전 이메일까지 포함
            cutoff_date = (datetime.now(texas_tz) - timedelta(days=5)).date()
            if email_date < cutoff_date:
                return None

        except Exception:
            pass

        for t in config.EMAIL_TARGETS:
            if exclude and tuple(t["contains"]) in exclude:
                continue

            if (
                (t.get("sender") is None or t.get("sender") in from_addr)  # sender check (None = skip)
                and all(c in subject for c in t["contains"])                # all keywords must match
                and (t.get("dates") is None or any(d in subject for d in t["dates"]))  # date check (None = skip)
            ):
                return t

        return None

    def _save_attachments(self, msg) -> None:
        """Recursively find and save all attachments from an email."""
        content_type = msg.get_content_type().lower()
        # logger.info(f"Part content-type: {content_type}") - For debugging

        if content_type.startswith('multipart'):
            for part in msg.get_payload():
                self._save_attachments(part)
            return

        if not (content_type.startswith('application') or content_type == 'text/csv'):
            return

        filename = self._extract_filename(msg)
        if not filename:
            return

        save_path = config.DATA_DIR / filename
        save_path.write_bytes(msg.get_payload(decode=True))
        self.downloaded_files.append(str(save_path))
        logger.info(f"Saved attachment: {save_path}")

    @staticmethod
    def _extract_filename(msg) -> Optional[str]:
        """Extract filename from Content-Disposition header, handles encoded names."""
        filename = msg.get_filename()
        if not filename:
            return None
        return filename.replace('\r\n', '').replace('\t', ' ')

    def _is_complete(self, matched_contains: set) -> bool:
        """Return True if all expected targets have been downloaded."""
        return len(matched_contains) >= config.EXPECTED_ATTACHMENT_COUNT
