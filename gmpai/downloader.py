"""Download regulation documents and maintain a manifest of what was archived."""

from __future__ import annotations

import hashlib
import json
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterable, Sequence

from .catalog import Document
from .fetcher import DEFAULT_RETRIES, DEFAULT_TIMEOUT, FetchError, Response, fetch, find_matching_links

MANIFEST_NAME = "manifest.json"

# Outcomes
OK = "ok"
UNCHANGED = "unchanged"
UPDATED = "updated"
SKIPPED = "skipped"
FAILED = "failed"


@dataclass
class DownloadResult:
    id: str
    title: str
    authority: str
    status: str
    filename: str | None = None
    path: str | None = None
    resolved_url: str | None = None
    sha256: str | None = None
    previous_sha256: str | None = None
    bytes: int | None = None
    content_type: str | None = None
    downloaded_at: str | None = None
    document_status: str | None = None
    document_date: str | None = None
    landing_page: str | None = None
    attempts: list[str] = field(default_factory=list)
    error: str | None = None

    @property
    def succeeded(self) -> bool:
        return self.status in (OK, UPDATED, UNCHANGED, SKIPPED)


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _now() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def load_manifest(out_dir: Path) -> dict:
    path = Path(out_dir) / MANIFEST_NAME
    if not path.exists():
        return {"generated_at": None, "documents": {}}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {"generated_at": None, "documents": {}}
    data.setdefault("documents", {})
    return data


def save_manifest(out_dir: Path, results: Iterable[DownloadResult]) -> Path:
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    manifest = load_manifest(out_dir)
    documents = manifest.get("documents", {})
    for result in results:
        if result.status == FAILED:
            # Keep the previous good record, but note the failure.
            entry = documents.get(result.id, {"id": result.id, "title": result.title})
            entry["last_error"] = result.error
            entry["last_attempt_at"] = _now()
            documents[result.id] = entry
            continue
        entry = {k: v for k, v in asdict(result).items() if v not in (None, [], "")}
        entry.pop("error", None)
        documents[result.id] = entry
    manifest["documents"] = documents
    manifest["generated_at"] = _now()
    path = out_dir / MANIFEST_NAME
    path.write_text(json.dumps(manifest, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    return path


def candidate_urls(doc: Document, prefer_discovery: bool = False, **fetch_kwargs) -> list[str]:
    """Build the ordered list of URLs to try for a document.

    The direct URL is tried first; the landing page is then scraped for a link
    matching the document's discover pattern, which is what keeps the catalog
    working when a regulator re-publishes a file under a new URL.
    """
    direct = [doc.direct_url] if doc.direct_url else []
    discovered: list[str] = []
    if doc.discover_pattern:
        try:
            response = fetch(doc.landing_page, **fetch_kwargs)
            discovered = find_matching_links(
                response.body.decode("utf-8", errors="replace"), response.url, doc.discover_pattern
            )
        except FetchError:
            discovered = []
    ordered = discovered + direct if prefer_discovery else direct + discovered
    seen: set[str] = set()
    return [u for u in ordered if not (u in seen or seen.add(u))]


def _validate_pdf(response: Response) -> str | None:
    """Return an error message if the response is not a usable PDF."""
    if response.looks_like_pdf:
        return None
    if response.looks_like_html:
        return (
            f"got an HTML page instead of a PDF (HTTP {response.status}) — the URL may have moved "
            "or the regulator's site is behind a consent/blocking page"
        )
    return f"unexpected content type {response.content_type!r} (HTTP {response.status})"


def download_document(
    doc: Document,
    out_dir: Path,
    manifest: dict | None = None,
    force: bool = False,
    dry_run: bool = False,
    prefer_discovery: bool = False,
    timeout: int = DEFAULT_TIMEOUT,
    retries: int = DEFAULT_RETRIES,
) -> DownloadResult:
    out_dir = Path(out_dir)
    manifest = manifest if manifest is not None else load_manifest(out_dir)
    previous = manifest.get("documents", {}).get(doc.id, {})
    previous_sha = previous.get("sha256")

    target_dir = out_dir / doc.authority
    target = target_dir / doc.filename

    result = DownloadResult(
        id=doc.id,
        title=doc.title,
        authority=doc.authority,
        status=FAILED,
        filename=doc.filename,
        path=str(target),
        document_status=doc.status,
        document_date=doc.document_date,
        landing_page=doc.landing_page,
        previous_sha256=previous_sha,
    )

    if target.exists() and previous_sha and not force:
        result.status = SKIPPED
        result.sha256 = previous_sha
        result.bytes = target.stat().st_size
        result.resolved_url = previous.get("resolved_url")
        result.downloaded_at = previous.get("downloaded_at")
        return result

    fetch_kwargs = {"timeout": timeout, "retries": retries}
    urls = candidate_urls(doc, prefer_discovery=prefer_discovery, **fetch_kwargs)
    if not urls:
        result.error = "no candidate URL: direct_url is unset and nothing on the landing page matched the discover pattern"
        return result

    if dry_run:
        result.status = OK
        result.resolved_url = urls[0]
        result.attempts = urls
        result.path = None
        return result

    errors: list[str] = []
    for url in urls:
        result.attempts.append(url)
        try:
            response = fetch(url, **fetch_kwargs)
        except FetchError as exc:
            errors.append(str(exc))
            continue

        problem = _validate_pdf(response)
        if problem:
            errors.append(f"{url}: {problem}")
            continue

        digest = sha256_bytes(response.body)
        target_dir.mkdir(parents=True, exist_ok=True)
        tmp = target.with_suffix(target.suffix + ".part")
        tmp.write_bytes(response.body)
        tmp.replace(target)

        result.status = UPDATED if previous_sha and previous_sha != digest else OK
        if previous_sha and previous_sha == digest:
            result.status = UNCHANGED
        result.resolved_url = response.url
        result.sha256 = digest
        result.bytes = len(response.body)
        result.content_type = response.content_type
        result.downloaded_at = _now()
        return result

    result.error = "; ".join(errors) or "download failed"
    return result


def download_all(
    documents: Sequence[Document],
    out_dir: Path,
    force: bool = False,
    dry_run: bool = False,
    prefer_discovery: bool = False,
    timeout: int = DEFAULT_TIMEOUT,
    retries: int = DEFAULT_RETRIES,
    on_result=None,
) -> list[DownloadResult]:
    out_dir = Path(out_dir)
    manifest = load_manifest(out_dir)
    results: list[DownloadResult] = []
    for doc in documents:
        result = download_document(
            doc,
            out_dir,
            manifest=manifest,
            force=force,
            dry_run=dry_run,
            prefer_discovery=prefer_discovery,
            timeout=timeout,
            retries=retries,
        )
        results.append(result)
        if on_result:
            on_result(result)
    if not dry_run:
        save_manifest(out_dir, results)
    return results
