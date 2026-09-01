"""예시 입력 파일 생성 (`python -m pqr demo`).

실제 데이터를 연결하기 전에 형식을 확인하고 전체 흐름을 돌려보기 위한 것입니다.
난수 시드가 고정되어 있어 언제 만들어도 같은 파일이 나옵니다.
"""

import csv
import datetime as _dt
import os
import random

PRODUCTS = [
    ("HP-101", "히알로타인 점안액 0.1%", "점안제", "무균 1공장", "김현우"),
    ("HP-110", "레보클린 점안액 0.5%", "점안제", "무균 1공장", "이서연"),
    ("HP-201", "로수바틴 정 10mg", "정제", "내용고형 2공장", "최민서"),
    ("HP-210", "에스오메졸 장용정 20mg", "정제", "내용고형 2공장", "정하늘"),
    ("HP-301", "세프트리 주 1g", "주사제", "주사 3공장", "윤도현"),
    ("HP-401", "암브록솔 시럽", "시럽제", "내용고형 2공장", "이서연"),
]

TESTS = {
    "점안제": [("함량", "%", 95.0, 105.0, 100.0, 1.3), ("pH", "", 6.5, 8.0, 7.2, 0.18),
             ("삼투압", "mOsm/kg", 260.0, 330.0, 295.0, 9.0)],
    "정제": [("함량", "%", 95.0, 105.0, 99.6, 1.5), ("용출률", "%", 80.0, None, 92.0, 3.4),
           ("경도", "N", 60.0, 140.0, 98.0, 9.5)],
    "주사제": [("함량", "%", 95.0, 105.0, 100.2, 1.1), ("pH", "", 6.0, 8.0, 6.9, 0.15),
            ("불용성이물", "개", None, 6.0, 2.1, 0.9)],
    "시럽제": [("함량", "%", 95.0, 105.0, 99.1, 1.6), ("pH", "", 4.0, 6.0, 5.1, 0.2),
            ("점도", "cP", 100.0, 300.0, 190.0, 22.0)],
}

PERIOD_FROM = _dt.date(2025, 7, 1)
PERIOD_TO = _dt.date(2026, 6, 30)
def _stages():
    """예시 자료의 단계 이름은 config 를 따릅니다.

    두 곳에 따로 적어 두면 config 만 고쳤을 때 예시의 단계 이력이 어느 단계에도
    붙지 않아 리드타임이 조용히 비어 버립니다 — 실제로 겪은 일입니다.
    """
    from . import build
    return list(build.load_config()["stages"])


STAGES = _stages()


def _write(path, header, rows):
    with open(path, "w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.writer(handle)
        writer.writerow(header)
        writer.writerows(rows)


def _product_folder(code, name):
    return "%s %s" % (code, name)


def write_samples(out_dir, layout="tree"):
    """예시 입력 파일을 만듭니다.

    layout="tree"  담당자가 쓰는 방식 — 제품 폴더마다 그 제품 자료를 넣습니다.
    layout="flat"  폴더 하나에 모든 제품 자료를 몰아넣습니다.
    """
    os.makedirs(out_dir, exist_ok=True)
    rng = random.Random(20260827)
    tables = {}          # dataset -> (헤더, {제품코드: [행, ...]})
    common = {}          # dataset -> (헤더, [행, ...])

    # 1. 제품 마스터 (공통)
    rows = []
    for index, (code, name, form, site, owner) in enumerate(PRODUCTS):
        # 보고서 3항(대상 제품)이 요구하는 값들도 마스터가 들고 있어야 합니다.
        rows.append([code, name, form, site, owner, PERIOD_FROM.isoformat(),
                     PERIOD_TO.isoformat(),
                     (PERIOD_TO + _dt.timedelta(days=60 + index * 12)).isoformat(),
                     STAGES[index % len(STAGES)],
                     "전문의약품" if index % 3 else "일반의약품",
                     "제 %d 호" % (4000 + index * 7),
                     (PERIOD_FROM - _dt.timedelta(days=900 + index * 31)).isoformat(),
                     "제조일로부터 %d개월" % (24 + (index % 3) * 12),
                     "기밀용기, 실온(1~30℃)보관"])
    common["products_제품마스터"] = (
        ["제품코드", "제품명", "제형", "공장", "담당자",
         "평가시작일", "평가종료일", "마감일", "단계",
         "제품분류", "허가번호", "허가일자", "사용기한", "보관조건"], rows)

    # 2. 배치 시험성적서 · 공정관리
    batch_rows = {}
    for code, name, form, _, _ in PRODUCTS:
        rows = []
        for batch_index in range(rng.randint(14, 26)):
            batch_no = "%s%03d" % (code.split("-")[1], batch_index + 1)
            mfg = PERIOD_FROM + _dt.timedelta(days=rng.randint(0, 360))
            rows.append([code, name, batch_no, mfg.isoformat(), "원자재",
                         "성상", "", "", "", "", "적합"])
            for test_name, unit, lsl, usl, center, sd in TESTS[form]:
                value = rng.gauss(center, sd)
                # 규격 이탈 배치를 의도적으로 몇 건 넣어 OOS · 경향 판정을 확인할 수 있게 합니다.
                if batch_index == 3 and test_name == "함량" and code in ("HP-201", "HP-301"):
                    value = (lsl or 0) - 1.2
                verdict = "적합"
                if lsl is not None and value < lsl:
                    verdict = "부적합"
                if usl is not None and value > usl:
                    verdict = "부적합"
                rows.append([code, name, batch_no, mfg.isoformat(), "완제품", test_name,
                             round(value, 2), unit,
                             "" if lsl is None else lsl, "" if usl is None else usl, verdict])
        batch_rows[code] = rows
    tables["batches_시험성적서"] = (
        ["제품코드", "제품명", "배치번호", "제조일", "구분단계", "시험항목",
         "결과값", "단위", "규격하한", "규격상한", "판정"], batch_rows)

    # 3. 일탈 · OOS/OOT · CAPA
    deviation_rows = {}
    counter = 0
    for code, _, _, _, _ in PRODUCTS:
        rows = []
        for _ in range(rng.randint(2, 7)):
            counter += 1
            kind = rng.choice(["일탈", "일탈", "일탈", "OOS", "OOT"])
            opened = PERIOD_FROM + _dt.timedelta(days=rng.randint(0, 400))
            closed_days = rng.choice([12, 24, 38, 55, None, None])
            closed = (opened + _dt.timedelta(days=closed_days)) if closed_days else None
            has_capa = rng.random() < 0.6
            rows.append([
                code, "DV-%s-%03d" % (opened.strftime("%y%m"), counter), kind,
                rng.choice(["중대", "중요", "경미", "경미"]), opened.isoformat(),
                rng.choice(["충전 중 정지", "함량 규격 이탈", "환경모니터링 초과",
                            "포장 라벨 오류", "온도 기록 누락"]),
                "종결" if closed else "조사 중",
                closed.isoformat() if closed else "",
                "CAPA-%03d" % counter if has_capa else "",
                (rng.choice(["완결", "진행", "진행"]) if has_capa else ""),
            ])
        deviation_rows[code] = rows
    tables["deviations_일탈대장"] = (
        ["제품코드", "관리번호", "구분", "등급", "발생일", "제목",
         "상태", "종결일", "CAPA번호", "CAPA상태"], deviation_rows)

    # 4. 변경관리 · 불만 · 회수
    change_rows = {}
    counter = 0
    for code, _, _, _, _ in PRODUCTS:
        rows = []
        for _ in range(rng.randint(2, 6)):
            counter += 1
            kind = rng.choice(["변경", "변경", "허가변경", "불만", "반품", "확약"])
            opened = PERIOD_FROM + _dt.timedelta(days=rng.randint(0, 400))
            closed_days = rng.choice([20, 45, 70, None])
            closed = (opened + _dt.timedelta(days=closed_days)) if closed_days else None
            rows.append([code, "CH-%s-%03d" % (opened.strftime("%y%m"), counter), kind,
                         opened.isoformat(),
                         rng.choice(["원료 공급원 추가", "분석법 이관", "포장재 변경",
                                     "이물 혼입 신고", "사용기한 표기 문의"]),
                         "완료" if closed else "진행",
                         closed.isoformat() if closed else ""])
        change_rows[code] = rows
    tables["changes_변경불만대장"] = (
        ["제품코드", "관리번호", "구분", "발생일", "제목", "상태", "완료일"], change_rows)

    # 5. 안정성 모니터링
    stability_rows = {}
    for code, _, form, _, _ in PRODUCTS:
        rows = []
        test_name, unit, lsl, usl, center, sd = TESTS[form][0]
        for batch_index in range(2):
            batch_no = "%s-S%d" % (code.split("-")[1], batch_index + 1)
            start_value = center - rng.uniform(0, 0.6)
            drift = rng.choice([-0.05, -0.12, -0.22, -0.02])
            for month in (0, 3, 6, 9, 12, 18, 24):
                value = start_value + drift * month + rng.gauss(0, 0.25)
                rows.append([code, batch_no, "장기(25℃/60%RH)", month, test_name,
                             round(value, 2), unit,
                             "" if lsl is None else lsl, "" if usl is None else usl])
            for month in (0, 3, 6):
                value = start_value + drift * 2.4 * month + rng.gauss(0, 0.3)
                rows.append([code, batch_no, "가속(40℃/75%RH)", month, test_name,
                             round(value, 2), unit,
                             "" if lsl is None else lsl, "" if usl is None else usl])
        stability_rows[code] = rows
    tables["stability_안정성"] = (
        ["제품코드", "배치번호", "조건", "시점", "시험항목",
         "결과값", "단위", "규격하한", "규격상한"], stability_rows)

    # 6. 설비 적격성 · 위수탁 협약 (공통)
    common["qualification_적격성"] = (
        ["설비명", "구분", "제품코드", "최종적격성일", "차기예정일", "상태"],
        [["AHU-01 (무균 1공장)", "HVAC", "", "2025-09-12", "2026-09-12", "유효"],
         ["AHU-02 (내용고형 2공장)", "HVAC", "", "2025-05-20", "2026-05-20", "재적격성 진행"],
         ["WFI 시스템", "제조용수", "", "2025-11-03", "2026-11-03", "유효"],
         ["정제수 시스템", "제조용수", "", "2025-08-18", "2026-08-18", "재적격성 필요"],
         ["압축공기 시스템", "압축가스", "", "2026-01-15", "2027-01-15", "유효"],
         ["충전기 FIL-03", "설비", "HP-101", "2025-10-01", "2026-10-01", "유효"],
         ["타정기 TAB-02", "설비", "HP-201", "2025-06-11", "2026-06-11", "재적격성 진행"],
         ["A 시험소 (위수탁 시험)", "위수탁", "", "2025-03-01", "2027-03-01", "유효"]])

    # 7. 단계 진행 이력 (선택 입력 — 리드타임 분석용, 공통)
    rows = []
    for index, (code, _, _, _, _) in enumerate(PRODUCTS):
        cursor = PERIOD_TO + _dt.timedelta(days=1)
        for stage in STAGES[:(index % len(STAGES)) + 1]:
            rows.append([code, stage, cursor.isoformat()])
            cursor += _dt.timedelta(days=rng.randint(4, 30))
    common["stagelog_단계이력"] = (["제품코드", "단계", "진입일"], rows)

    written = []
    for stem, (header, rows) in sorted(common.items()):
        path = os.path.join(out_dir, stem + ".csv")
        _write(path, header, rows)
        written.append(path)

    if layout == "flat":
        for stem, (header, grouped) in sorted(tables.items()):
            merged = [row for code, _, _, _, _ in PRODUCTS for row in grouped[code]]
            path = os.path.join(out_dir, stem + ".csv")
            _write(path, header, merged)
            written.append(path)
        return written

    # 제품 폴더 구조: 제품코드 열은 빼고, 폴더 이름이 제품을 가리키게 합니다.
    for code, name, _, _, _ in PRODUCTS:
        folder = os.path.join(out_dir, _product_folder(code, name))
        os.makedirs(folder, exist_ok=True)
        for stem, (header, grouped) in sorted(tables.items()):
            path = os.path.join(folder, stem.split("_", 1)[1] + ".csv")
            _write(path, header[1:], [row[1:] for row in grouped[code]])
            written.append(path)
    return written
