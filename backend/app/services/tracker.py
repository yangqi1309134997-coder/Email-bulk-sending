import re
from urllib.parse import urlparse
from ..config import settings


def inject_tracking_pixel(body_html: str, log_id: int) -> str:
    pixel = f'<img src="{settings.TRACKING_DOMAIN}/track/open/{log_id}" width="1" height="1" alt="" style="display:none;">'
    if "</body>" in body_html:
        return body_html.replace("</body>", f"{pixel}</body>")
    return body_html + pixel


def replace_links_with_tracking(body_html: str, log_id: int) -> str:
    """Replace all href links with tracking redirects."""
    pattern = r'href="(https?://[^"]+)"'

    def replace_match(match):
        original_url = match.group(1)
        # Skip tracking URLs to avoid infinite loops
        if "/track/" in original_url:
            return match.group(0)
        tracking_url = f'{settings.TRACKING_DOMAIN}/track/click/{log_id}?url={original_url}'
        return f'href="{tracking_url}"'

    return re.sub(pattern, replace_match, body_html)


def is_valid_redirect_url(url: str) -> bool:
    """Basic validation to prevent open redirect attacks."""
    try:
        parsed = urlparse(url)
        return parsed.scheme in ("http", "https") and bool(parsed.netloc)
    except Exception:
        return False