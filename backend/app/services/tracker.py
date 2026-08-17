import html
import ipaddress
from html.parser import HTMLParser
from urllib.parse import urlencode, urlsplit

from ..config import settings
from ..utils.security import create_tracking_signature


def _tracking_url(path: str, query: dict[str, str]) -> str:
    base = settings.TRACKING_DOMAIN.rstrip("/")
    return f"{base}{path}?{urlencode(query)}"


def inject_tracking_pixel(body_html: str, log_id: int) -> str:
    signature = create_tracking_signature(log_id, "open")
    source = _tracking_url(f"/track/open/{int(log_id)}", {"sig": signature})
    pixel = (
        f'<img src="{html.escape(source, quote=True)}" width="1" height="1" '
        'alt="" style="display:none;">'
    )
    lower_body = body_html.lower()
    closing_index = lower_body.rfind("</body>")
    if closing_index >= 0:
        return body_html[:closing_index] + pixel + body_html[closing_index:]
    return body_html + pixel


class _TrackingLinkParser(HTMLParser):
    def __init__(self, log_id: int):
        super().__init__(convert_charrefs=False)
        self.log_id = int(log_id)
        self.parts: list[str] = []

    def _rewrite_attrs(self, attrs: list[tuple[str, str | None]]) -> str:
        rendered = []
        signature = create_tracking_signature(self.log_id, "click")
        for name, value in attrs:
            if value is None:
                rendered.append(name)
                continue
            rewritten = value
            if name.lower() == "href":
                parsed = urlsplit(value)
                tracking_host = urlsplit(settings.TRACKING_DOMAIN).hostname
                is_tracking_link = (
                    parsed.hostname == tracking_host and parsed.path.startswith("/track/")
                )
                if parsed.scheme.lower() in {"http", "https"} and not is_tracking_link:
                    rewritten = _tracking_url(
                        f"/track/click/{self.log_id}",
                        {"url": value, "sig": signature},
                    )
            rendered.append(f'{name}="{html.escape(rewritten, quote=True)}"')
        return (" " + " ".join(rendered)) if rendered else ""

    def handle_starttag(self, tag, attrs):
        self.parts.append(f"<{tag}{self._rewrite_attrs(attrs)}>")

    def handle_startendtag(self, tag, attrs):
        self.parts.append(f"<{tag}{self._rewrite_attrs(attrs)}/>")

    def handle_endtag(self, tag):
        self.parts.append(f"</{tag}>")

    def handle_data(self, data):
        self.parts.append(data)

    def handle_entityref(self, name):
        self.parts.append(f"&{name};")

    def handle_charref(self, name):
        self.parts.append(f"&#{name};")

    def handle_comment(self, data):
        self.parts.append(f"<!--{data}-->")

    def handle_decl(self, decl):
        self.parts.append(f"<!{decl}>")

    def handle_pi(self, data):
        self.parts.append(f"<?{data}>")


def replace_links_with_tracking(body_html: str, log_id: int) -> str:
    """Rewrite HTTP(S) links using a real HTML parser and signed query data."""
    parser = _TrackingLinkParser(log_id)
    try:
        parser.feed(body_html)
        parser.close()
        return "".join(parser.parts)
    except Exception:
        return body_html


def is_valid_redirect_url(url: str) -> bool:
    """Allow public HTTP(S) destinations while rejecting local/private targets."""
    try:
        if any(ord(character) < 32 for character in url):
            return False
        parsed = urlsplit(url)
        if parsed.scheme.lower() not in {"http", "https"}:
            return False
        if not parsed.hostname or parsed.username or parsed.password:
            return False
        _ = parsed.port
        hostname = parsed.hostname.rstrip(".").lower()
        if hostname == "localhost" or hostname.endswith((".localhost", ".local")):
            return False
        try:
            address = ipaddress.ip_address(hostname)
        except ValueError:
            if "." not in hostname:
                return False
        else:
            if not address.is_global:
                return False
        return True
    except (TypeError, ValueError):
        return False
