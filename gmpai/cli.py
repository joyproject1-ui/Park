"""Command line interface: list / download / verify / index."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .catalog import Catalog, CatalogError, load_catalog
from .downloader import (
    FAILED,
    OK,
    SKIPPED,
    UNCHANGED,
    UPDATED,
    candidate_urls,
    download_all,
    load_manifest,
)
from .fetcher import DEFAULT_RETRIES, DEFAULT_TIMEOUT, FetchError, fetch
from .index import write_indexes

DEFAULT_OUT_DIR = "downloads"

MARKS = {OK: "[받음]", UPDATED: "[갱신]", UNCHANGED: "[동일]", SKIPPED: "[보유]", FAILED: "[실패]"}


def _add_selection_args(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--id", dest="ids", action="append", help="특정 문서 ID만 선택 (반복 지정 가능)")
    parser.add_argument("--authority", help="발행 기관 필터 (EU, FDA)")
    parser.add_argument("--category", help="분류 필터 (gmp-ai, computerised-systems, ...)")
    parser.add_argument("--status", help="상태 필터 (draft, final, in-force, discussion-paper)")


def _add_network_args(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--timeout", type=int, default=DEFAULT_TIMEOUT, help=f"요청 타임아웃 초 (기본 {DEFAULT_TIMEOUT})")
    parser.add_argument("--retries", type=int, default=DEFAULT_RETRIES, help=f"재시도 횟수 (기본 {DEFAULT_RETRIES})")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="gmpai",
        description="EU(EC/EMA)와 FDA의 GMP·AI 관련 규정 원문을 공식 사이트에서 내려받습니다.",
    )
    parser.add_argument("--catalog", help="사용할 카탈로그 JSON 경로 (기본: 내장 카탈로그)")
    sub = parser.add_subparsers(dest="command", required=True)

    p_list = sub.add_parser("list", help="카탈로그에 등록된 문서 목록 출력")
    _add_selection_args(p_list)
    p_list.add_argument("--long", action="store_true", help="URL과 비고까지 함께 출력")

    p_download = sub.add_parser("download", help="문서 PDF 내려받기")
    _add_selection_args(p_download)
    _add_network_args(p_download)
    p_download.add_argument("-o", "--out", default=DEFAULT_OUT_DIR, help=f"저장 디렉터리 (기본 {DEFAULT_OUT_DIR})")
    p_download.add_argument("--force", action="store_true", help="이미 받은 파일도 다시 내려받기")
    p_download.add_argument("--dry-run", action="store_true", help="실제 저장 없이 사용할 URL만 확인")
    p_download.add_argument(
        "--prefer-discovery",
        action="store_true",
        help="직접 URL보다 출처 페이지에서 찾은 링크를 우선 사용",
    )

    p_verify = sub.add_parser("verify", help="등록된 링크가 아직 살아 있는지 점검")
    _add_selection_args(p_verify)
    _add_network_args(p_verify)

    p_index = sub.add_parser("index", help="INDEX.md와 docs/index.html 다시 생성")
    _add_selection_args(p_index)
    p_index.add_argument("--markdown", default="INDEX.md", help="Markdown 목록 경로")
    p_index.add_argument("--html", default="docs/index.html", help="HTML 목록 경로")

    p_status = sub.add_parser("status", help="이미 내려받은 파일 상태 확인")
    p_status.add_argument("-o", "--out", default=DEFAULT_OUT_DIR, help=f"저장 디렉터리 (기본 {DEFAULT_OUT_DIR})")

    return parser


def _select(catalog: Catalog, args: argparse.Namespace):
    return catalog.filter(
        ids=getattr(args, "ids", None),
        authority=getattr(args, "authority", None),
        category=getattr(args, "category", None),
        status=getattr(args, "status", None),
    )


def cmd_list(catalog: Catalog, args: argparse.Namespace) -> int:
    docs = _select(catalog, args)
    if not docs:
        print("조건에 맞는 문서가 없습니다.")
        return 1
    for doc in docs:
        print(f"{doc.id:<38} {doc.authority:<4} {doc.status:<16} {doc.document_date or '-':<10} {doc.title_ko}")
        if args.long:
            print(f"{'':<38} 원문: {doc.title}")
            print(f"{'':<38} 직접 링크: {doc.direct_url or '(출처 페이지에서 탐색)'}")
            print(f"{'':<38} 출처 페이지: {doc.landing_page}")
            if doc.notes:
                print(f"{'':<38} 비고: {doc.notes}")
            print()
    print(f"\n총 {len(docs)}건 (카탈로그 검토일 {catalog.last_reviewed})")
    return 0


def cmd_download(catalog: Catalog, args: argparse.Namespace) -> int:
    docs = _select(catalog, args)
    if not docs:
        print("조건에 맞는 문서가 없습니다.")
        return 1

    out_dir = Path(args.out)
    print(f"{len(docs)}건을 처리합니다 → {out_dir.resolve()}\n")

    def report(result) -> None:
        mark = MARKS.get(result.status, result.status)
        line = f"{mark} {result.id}"
        if result.status == FAILED:
            print(f"{line}\n      {result.error}")
            return
        if args.dry_run:
            print(f"{line} → {result.resolved_url}")
            return
        size = f"{(result.bytes or 0) / 1024:.0f} KB"
        print(f"{line} ({size}) → {result.path}")
        if result.status == UPDATED:
            print(f"      내용이 바뀌었습니다: {(result.previous_sha256 or '')[:12]} → {(result.sha256 or '')[:12]}")

    results = download_all(
        docs,
        out_dir,
        force=args.force,
        dry_run=args.dry_run,
        prefer_discovery=args.prefer_discovery,
        timeout=args.timeout,
        retries=args.retries,
        on_result=report,
    )

    failed = [r for r in results if r.status == FAILED]
    updated = [r for r in results if r.status == UPDATED]
    print(f"\n성공 {len(results) - len(failed)}건, 실패 {len(failed)}건", end="")
    print(f", 내용 변경 {len(updated)}건" if updated else "")
    if not args.dry_run:
        print(f"이력: {out_dir / 'manifest.json'}")
    if failed:
        print("\n실패한 문서는 출처 페이지에서 직접 확인하세요:")
        for r in failed:
            print(f"  - {r.id}: {r.landing_page}")
        return 1
    return 0


def cmd_verify(catalog: Catalog, args: argparse.Namespace) -> int:
    docs = _select(catalog, args)
    problems = 0
    for doc in docs:
        urls = candidate_urls(doc, timeout=args.timeout, retries=1)
        if not urls:
            print(f"[없음] {doc.id}: 사용할 수 있는 URL을 찾지 못했습니다 ({doc.landing_page})")
            problems += 1
            continue
        url = urls[0]
        try:
            response = fetch(url, timeout=args.timeout, retries=args.retries, max_bytes=2048)
        except FetchError as exc:
            print(f"[실패] {doc.id}: {exc}")
            problems += 1
            continue
        if response.looks_like_pdf:
            print(f"[정상] {doc.id} → {response.url}")
        else:
            kind = "HTML 페이지" if response.looks_like_html else response.content_type or "알 수 없음"
            print(f"[주의] {doc.id}: PDF가 아닌 응답({kind}) → {response.url}")
            problems += 1
    print(f"\n점검 {len(docs)}건 중 문제 {problems}건")
    return 1 if problems else 0


def cmd_index(catalog: Catalog, args: argparse.Namespace) -> int:
    docs = _select(catalog, args)
    written = write_indexes(catalog, docs, markdown_path=args.markdown, html_path=args.html)
    for path in written:
        print(f"생성: {path}")
    return 0


def cmd_status(catalog: Catalog, args: argparse.Namespace) -> int:
    out_dir = Path(args.out)
    manifest = load_manifest(out_dir)
    entries = manifest.get("documents", {})
    if not entries:
        print(f"{out_dir}에 내려받은 기록이 없습니다. `python -m gmpai download`를 먼저 실행하세요.")
        return 1
    print(f"기록 생성 시각: {manifest.get('generated_at')}\n")
    for doc in catalog.documents:
        entry = entries.get(doc.id)
        if not entry:
            print(f"[미보유] {doc.id}")
            continue
        if entry.get("last_error"):
            print(f"[오류] {doc.id}: {entry['last_error']}")
            continue
        exists = Path(entry.get("path", "")).exists()
        mark = "[보유]" if exists else "[파일없음]"
        print(f"{mark} {doc.id}  {entry.get('downloaded_at', '-')}  sha256:{(entry.get('sha256') or '')[:12]}")
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        catalog = load_catalog(args.catalog)
    except CatalogError as exc:
        print(f"카탈로그 오류: {exc}", file=sys.stderr)
        return 2

    handlers = {
        "list": cmd_list,
        "download": cmd_download,
        "verify": cmd_verify,
        "index": cmd_index,
        "status": cmd_status,
    }
    try:
        return handlers[args.command](catalog, args)
    except CatalogError as exc:
        print(f"오류: {exc}", file=sys.stderr)
        return 2
    except KeyboardInterrupt:
        print("\n중단했습니다.", file=sys.stderr)
        return 130


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
