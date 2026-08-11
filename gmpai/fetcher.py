"""HTTP access layer: stdlib-only fetching, link discovery and PDF sniffing."""

from __future__ import annotations

import re
import ssl
import time
import urllib.error
import urllib.request
from dataclasses import dataclass
from html.parser import HTMLParser
from typing import Callable
from urllib.parse import urljoin

USER_AGENT = (
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/124.0 Safari/537.36 gmp-ai-regs/1.0"
)
DEFAULT_TIMEOUT = 60
DEFAULT_RETRIES = 3
PDF_MAGIC = b"%PDF-"
PDF_CONTENT_TYPES = ("application/pdf", "application/octet-stream", "binary/octet-stream")


class FetchError(Exception):
    """Raised when a URL cannot be retrieved after all retries."""


@dataclass
class Response:
    url: str
    status: int
    content_type: str
    body: bytes

    @property
    def looks_like_pdf(self) -> bool:
        if self.body.startswith(PDF_MAGIC):
            return True
        main_type = self.content_type.split(";")[0].strip().lower()
        return main_type in PDF_CONTENT_TYPES and PDF_MAGIC in self.body[:1024]

    @property
    def looks_like_html(self) -> bool:
        main_type = self.content_type.split(";")[0].strip().lower()
        return main_type in ("text/html", "application/xhtml+xml") or self.body[:512].lstrip().lower().startswith(
            (b"<!doctype html", b"<html")
        )


class _LinkParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.links: list[tuple[str, str]] = []
        self._current_href: str | None = None
        self._text_parts: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag != "a":
            return
        href = dict(attrs).get("href")
        if href:
            self._current_href = href
            self._text_parts = []

    def handle_data(self, data: str) -> None:
        if self._current_href is not None:
            self._text_parts.append(data)

    def handle_endtag(self, tag: str) -> None:
        if tag == "a" and self._current_href is not None:
            text = " ".join("".join(self._text_parts).split())
            self.links.append((self._current_href, text))
            self._current_href = None
            self._text_parts = []


def extract_links(html: str, base_url: str) -> list[tuple[str, str]]:
    """Return (absolute_url, link_text) pairs in document order."""
    parser = _LinkParser()
    parser.feed(html)
    parser.close()
    return [(urljoin(base_url, href), text) for href, text in parser.links]


def find_matching_links(html: str, base_url: str, pattern: str, limit: int = 5) -> list[str]:
    """Return absolute URLs whose href or link text matches `pattern`."""
    regex = re.compile(pattern)
    matches: list[str] = []
    for url, text in extract_links(html, base_url):
        if regex.search(url) or regex.search(text):
            if url not in matches:
                matches.append(url)
        if len(matches) >= limit:
            break
    return matches


def _decode(body: bytes) -> str:
    for encoding in ("utf-8", "cp1252", "latin-1"):
        try:
            return body.decode(encoding)
        except UnicodeDecodeError:
            continue
    return body.decode("utf-8", errors="replace")


def fetch(
    url: str,
    timeout: int = DEFAULT_TIMEOUT,
    retries: int = DEFAULT_RETRIES,
    max_bytes: int | None = None,
    sleep: Callable[[float], None] = time.sleep,
    opener: Callable[..., object] | None = None,
) -> Response:
    """GET a URL with retries and exponential backoff.

    Proxies are taken from the standard HTTP(S)_PROXY environment variables via
    urllib's default handlers, which matters inside corporate networks.
    """
    request_opener = opener or urllib.request.urlopen
    context = ssl.create_default_context()
    last_error: Exception | None = None

    for attempt in range(1, max(1, retries) + 1):
        request = urllib.request.Request(
            url,
            headers={
                "User-Agent": USER_AGENT,
                "Accept": "application/pdf,text/html;q=0.9,*/*;q=0.8",
                "Accept-Language": "en",
            },
        )
        try:
            kwargs = {"timeout": timeout}
            if opener is None:
                kwargs["context"] = context
            with request_opener(request, **kwargs) as response:  # type: ignore[misc]
                body = response.read(max_bytes) if max_bytes else response.read()
                return Response(
                    url=response.geturl(),
                    status=getattr(response, "status", 200) or 200,
                    content_type=response.headers.get("Content-Type", "") if response.headers else "",
                    body=body,
                )
        except urllib.error.HTTPError as exc:
            last_error = exc
            # Client errors other than rate limiting will not fix themselves.
            if exc.code not in (408, 429) and exc.code < 500:
                break
        except (urllib.error.URLError, TimeoutError, ssl.SSLError, OSError) as exc:
            last_error = exc

        if attempt < retries:
            sleep(2 ** attempt)

    raise FetchError(f"{url}: {last_error}") from last_error


def fetch_text(url: str, **kwargs) -> str:
    return _decode(fetch(url, **kwargs).body)
