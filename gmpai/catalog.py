"""Loading and filtering of the regulation source catalog."""

from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable, Sequence

CATALOG_PATH = Path(__file__).with_name("data") / "catalog.json"

REQUIRED_FIELDS = ("id", "title", "authority", "category", "status", "filename", "landing_page")


class CatalogError(Exception):
    """Raised when the catalog file is missing or malformed."""


@dataclass(frozen=True)
class Document:
    id: str
    title: str
    title_ko: str
    authority: str
    issuer: str
    category: str
    status: str
    document_date: str
    reference: str
    landing_page: str
    direct_url: str | None
    discover_pattern: str | None
    filename: str
    notes: str
    language: str

    @property
    def urls_hint(self) -> str:
        return self.direct_url or self.landing_page


@dataclass(frozen=True)
class Catalog:
    schema_version: int
    last_reviewed: str
    documents: tuple[Document, ...]

    def __iter__(self):
        return iter(self.documents)

    def __len__(self) -> int:
        return len(self.documents)

    def get(self, doc_id: str) -> Document:
        for doc in self.documents:
            if doc.id == doc_id:
                return doc
        raise KeyError(doc_id)

    def filter(
        self,
        ids: Sequence[str] | None = None,
        authority: str | None = None,
        category: str | None = None,
        status: str | None = None,
    ) -> list[Document]:
        selected = list(self.documents)
        if ids:
            wanted = list(ids)
            unknown = [i for i in wanted if not any(d.id == i for d in selected)]
            if unknown:
                raise CatalogError(f"unknown document id(s): {', '.join(unknown)}")
            selected = [d for d in selected if d.id in wanted]
        if authority:
            selected = [d for d in selected if d.authority.lower() == authority.lower()]
        if category:
            selected = [d for d in selected if d.category.lower() == category.lower()]
        if status:
            selected = [d for d in selected if d.status.lower() == status.lower()]
        return selected

    @property
    def authorities(self) -> list[str]:
        return sorted({d.authority for d in self.documents})

    @property
    def categories(self) -> list[str]:
        return sorted({d.category for d in self.documents})


def _document_from_dict(raw: dict[str, Any]) -> Document:
    missing = [f for f in REQUIRED_FIELDS if not raw.get(f)]
    if missing:
        raise CatalogError(f"document {raw.get('id', '<no id>')} is missing field(s): {', '.join(missing)}")

    discover = raw.get("discover") or None
    pattern = None
    if discover:
        pattern = discover.get("pattern")
        if pattern:
            try:
                re.compile(pattern)
            except re.error as exc:  # pragma: no cover - guarded by test
                raise CatalogError(f"document {raw['id']} has an invalid discover pattern: {exc}") from exc

    if not raw.get("direct_url") and not pattern:
        raise CatalogError(f"document {raw['id']} has neither a direct_url nor a discover pattern")

    return Document(
        id=raw["id"],
        title=raw["title"],
        title_ko=raw.get("title_ko", raw["title"]),
        authority=raw["authority"],
        issuer=raw.get("issuer", ""),
        category=raw["category"],
        status=raw["status"],
        document_date=raw.get("document_date", ""),
        reference=raw.get("reference", ""),
        landing_page=raw["landing_page"],
        direct_url=raw.get("direct_url") or None,
        discover_pattern=pattern,
        filename=raw["filename"],
        notes=raw.get("notes", ""),
        language=raw.get("language", "en"),
    )


def load_catalog(path: Path | str | None = None) -> Catalog:
    catalog_path = Path(path) if path else CATALOG_PATH
    try:
        raw = json.loads(catalog_path.read_text(encoding="utf-8"))
    except FileNotFoundError as exc:
        raise CatalogError(f"catalog not found: {catalog_path}") from exc
    except json.JSONDecodeError as exc:
        raise CatalogError(f"catalog is not valid JSON ({catalog_path}): {exc}") from exc

    documents = [_document_from_dict(d) for d in raw.get("documents", [])]
    if not documents:
        raise CatalogError(f"catalog {catalog_path} contains no documents")

    seen: set[str] = set()
    for doc in documents:
        if doc.id in seen:
            raise CatalogError(f"duplicate document id in catalog: {doc.id}")
        seen.add(doc.id)

    return Catalog(
        schema_version=int(raw.get("schema_version", 1)),
        last_reviewed=raw.get("last_reviewed", ""),
        documents=tuple(documents),
    )


def group_by_authority(documents: Iterable[Document]) -> dict[str, list[Document]]:
    grouped: dict[str, list[Document]] = {}
    for doc in documents:
        grouped.setdefault(doc.authority, []).append(doc)
    return dict(sorted(grouped.items()))
