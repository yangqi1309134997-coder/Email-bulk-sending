import smtplib
import time
import random
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os
from typing import List, Tuple
from ..models.sender import Sender
from ..utils.security import decrypt_password
from ..config import settings


class EmailSender:
    """SMTP 邮件发送引擎"""

    MAX_RETRIES = 3
    SMTP_TIMEOUT = 30

    def build_message(
        self,
        sender_email: str,
        recipient_email: str,
        recipient_name: str,
        subject: str,
        body_html: str,
        attachments: List[str] = None,
    ) -> MIMEMultipart:
        msg = MIMEMultipart()
        msg["From"] = sender_email
        msg["To"] = recipient_email

        # Personalize subject and body
        name = recipient_name or "朋友"
        msg["Subject"] = subject.replace("{name}", name).replace("{email}", recipient_email)

        personalized_body = body_html.replace("{name}", name).replace("{email}", recipient_email)

        # Add tracking pixel
        tracking_pixel = f'<img src="{settings.TRACKING_DOMAIN}/track/open/{{LOG_ID}}" width="1" height="1" alt="" style="display:none;">'
        personalized_body = personalized_body.replace("</body>", f"{tracking_pixel}</body>")
        if "</body>" not in personalized_body:
            personalized_body += tracking_pixel

        msg.attach(MIMEText(personalized_body, "html", "utf-8"))

        # Add attachments
        if attachments:
            for path in attachments:
                if os.path.exists(path):
                    try:
                        with open(path, "rb") as f:
                            part = MIMEBase("application", "octet-stream")
                            part.set_payload(f.read())
                        encoders.encode_base64(part)
                        part.add_header("Content-Disposition", f'attachment; filename="{os.path.basename(path)}"')
                        msg.attach(part)
                    except Exception:
                        pass

        return msg

    def send(self, sender: Sender, recipient_email: str, recipient_name: str, subject: str, body_html: str, attachments: List[str] = None) -> Tuple[bool, str]:
        """Send email with retry logic. Returns (success, error_message)."""
        for attempt in range(self.MAX_RETRIES + 1):
            try:
                pwd = decrypt_password(sender.password)
                server = smtplib.SMTP(sender.smtp_server, sender.smtp_port, timeout=self.SMTP_TIMEOUT)
                if sender.use_tls:
                    server.starttls()
                server.login(sender.email, pwd)

                msg = self.build_message(sender.email, recipient_email, recipient_name, subject, body_html, attachments)
                server.send_message(msg)
                server.quit()
                return True, ""
            except smtplib.SMTPAuthenticationError as e:
                return False, f"Auth failed: {str(e)}"
            except smtplib.SMTPResponseException as e:
                if e.smtp_code in (552, 554):
                    return False, f"Rate limited: {str(e)}"
                if attempt < self.MAX_RETRIES:
                    time.sleep(2 ** attempt)
                    continue
                return False, f"SMTP error: {str(e)}"
            except Exception as e:
                if attempt < self.MAX_RETRIES:
                    time.sleep(2 ** attempt)
                    continue
                return False, f"Send failed: {str(e)}"

        return False, "Max retries exceeded"


email_sender = EmailSender()