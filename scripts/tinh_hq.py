"""Tính HQ cho VinaPro và VinaKey1 từ báo cáo đã xuất."""

import argparse
import csv
import json
import re
import unicodedata
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

import xlrd
from openpyxl import load_workbook


POINT_FIELDS = (
    "unit", "channel", "point_code", "account", "points", "revenue_points",
    "production_points", "ptm_revenue", "reuse_revenue", "hq_revenue", "production",
)
REFERENCE_COLUMNS = {
    "VINAPRO": {"points": 7, "revenue_points": 8, "production_points": 11,
                "hq_revenue": 9, "production": 12},
    "VINAKEY1": {"points": 16, "revenue_points": 17, "production_points": 20,
                 "hq_revenue": 18, "production": 21},
}
REUSE_WEIGHT = {"VINAPRO": Decimal("0.8"), "VINAKEY1": Decimal("0.5")}


def key(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value).strip().upper()


def header(value):
    text = unicodedata.normalize("NFKD", key(value).replace("Đ", "D"))
    return re.sub(r"[^A-Z0-9]", "", text)


def money(value):
    if value is None or value == "":
        return Decimal(0)
    amount = Decimal(str(value))
    if not amount.is_finite():
        raise ValueError("Giá trị doanh thu không hữu hạn")
    return amount


def workbook_rows(path, sheet=None):
    if path.suffix.lower() == ".xls":
        book = xlrd.open_workbook(str(path), on_demand=True)
        try:
            tab = book.sheet_by_name(sheet) if sheet else book.sheet_by_index(0)
            yield from (tuple(tab.row_values(i)) for i in range(tab.nrows))
        finally:
            book.release_resources()
        return
    book = load_workbook(path, read_only=True, data_only=True, keep_links=False)
    try:
        yield from (book[sheet] if sheet else book.worksheets[0]).values
    finally:
        book.close()


def table(rows, columns, label):
    for row_number, row in enumerate(rows, 1):
        if row_number > 15:
            break
        names = [header(value) for value in row]
        positions = {}
        for field, aliases in columns.items():
            matches = [index for index, name in enumerate(names) if name in aliases]
            if len(matches) == 1:
                positions[field] = matches[0]
        if len(positions) == len(columns):
            for number, record in enumerate(rows, row_number + 1):
                values = {field: record[i] if i < len(record) else None
                          for field, i in positions.items()}
                if any(value not in (None, "") for value in values.values()):
                    yield number, values
            return
    raise ValueError(f"{label}: không tìm thấy tiêu đề bắt buộc ({', '.join(columns)})")


@dataclass(frozen=True)
class Point:
    code: str
    account: str
    unit: str
    channel: str


def catalog(path, sheet):
    columns = {
        "code": {"MADIEMBAN", "POINTCODE"}, "eload": {"SOELOAD", "ELOAD"},
        "hrm": {"MAHRM", "HRM"}, "unit": {"MADV", "MADONVI", "UNIT"},
        "channel": {"LOAIKENH", "KENH", "CHANNEL"}, "subject": {"DOITUONG", "SUBJECT"},
    }
    for _, row in table(workbook_rows(path, sheet), columns, "Danh mục SMCS"):
        channel = header(row["channel"])
        if header(row["subject"]) == "DIEMUYQUYEN":
            channel = "VINAKEY1"
        if channel in REUSE_WEIGHT:
            yield Point(key(row["code"]), key(row["hrm" if channel == "VINAPRO" else "eload"]),
                        key(row["unit"]), channel)


def reference_catalog(path):
    for channel, sheet, code, account, unit, last in (
        ("VINAPRO", "Hiệu quả_VINAPRO", 2, 2, 3, 29),
        ("VINAKEY1", "Hiệu quả_VINAKEY1", 1, 4, 2, 203),
    ):
        for number, row in enumerate(workbook_rows(path, sheet), 1):
            if 3 <= number <= last and key(row[account]):
                yield Point(key(row[code]), key(row[account]), key(row[unit]), channel)


def unique_index(points, field):
    candidates = defaultdict(set)
    for point in points:
        value = getattr(point, field)
        if value:
            candidates[value].add(point)
    return {value: next(iter(items)) if len(items) == 1 else None
            for value, items in candidates.items()}


def activation_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value or "").strip().split(" ")[0]
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", text):
        return date.fromisoformat(text)
    return datetime.strptime(text, "%d/%m/%Y").date()


def get_ims(path, sheet):
    columns = {"account": {"TAIKHOANDKY"}, "amount": {"GIAGOI"},
               "type": {"LOAITB"}, "phone": {"SOMAY"}, "unit": {"DONVI"}}
    return table(workbook_rows(path, sheet), columns, "IMS73")


def get_bris(path, sheet):
    columns = {"account": {"USERELOADDKTTTB"}, "phone": {"SOTB"},
               "activated": {"NGAYKICHHOAT"}}
    return table(workbook_rows(path, sheet), columns, "BRIS")


def lookup(path, sheet, key_col, value_col):
    for number, row in enumerate(workbook_rows(path, sheet), 1):
        if number > 1 and key(row[key_col]) and row[value_col] not in (None, ""):
            yield key(row[key_col]), row[value_col]


def calculate(args, points):
    accounts = unique_index(points, "account")
    codes = unique_index((p for p in points if p.channel == "VINAKEY1"), "code")
    td49 = {account for account, marker in lookup(args.td49, "ALL_TD49", 0, 1)
            if key(marker) == "TD49"} if args.td49 else set()
    audit = Counter()
    exceptions = []
    amount_by_point = defaultdict(Decimal)
    reuse_by_point = defaultdict(Decimal)
    production = Counter()

    def resolve(account, source, number):
        index = codes if account.startswith("NAN") else accounts
        point = index.get(account)
        if point and point.account and point.unit and accounts.get(point.account) == point:
            if index is codes:
                audit[f"{source}_nan_resolved"] += 1
            return point
        reason = "missing_account" if not account else "unmatched_or_ambiguous_account"
        if account.startswith("NAN"):
            reason = "unresolved_nan"
        audit[f"{source}_{reason}"] += 1
        exceptions.append({"source": source, "row": number, "reason": reason})
        return None

    if args.sim_tsd:
        for account, amount in lookup(args.sim_tsd, "Sim TSD", 5, 9):
            point = accounts.get(account)
            if point and point.channel == "VINAKEY1":
                reuse_by_point[point] += money(amount)

    for number, row in get_ims(args.ims, args.ims_sheet):
        audit["ims_rows"] += 1
        kind = key(row["type"])
        is_reuse = kind == "HIEN HUU" and key(row["phone"]) in td49
        if kind != "KHM" and not is_reuse:
            continue
        point = resolve(key(row["account"]), "ims", number)
        if point is None:
            continue
        if kind == "KHM":
            amount_by_point[point] += money(row["amount"])
        elif point.channel == "VINAPRO":
            reuse_by_point[point] += money(row["amount"])
        if key(row["unit"]) != point.unit:
            audit["ims_transaction_unit_differs"] += 1

    period_phones = set()
    for number, row in get_bris(args.bris, args.bris_sheet):
        if not key(row["phone"]):
            audit["bris_non_detail_rows"] += 1
            continue
        audit["bris_detail_rows"] += 1
        try:
            day = activation_date(row["activated"])
        except ValueError:
            audit["bris_invalid_activation_date"] += 1
            exceptions.append({"source": "bris", "row": number, "reason": "invalid_activation_date"})
            continue
        if (args.bris_from and day < args.bris_from) or (args.bris_to and day > args.bris_to):
            audit["bris_outside_period"] += 1
            continue
        period_phones.add(key(row["phone"]))
        audit["bris_in_period"] += 1
        point = resolve(key(row["account"]), "bris", number)
        if point:
            production[point] += 1

    detail = []
    for point in sorted(points, key=lambda p: (p.unit, p.channel, p.code, p.account)):
        if not point.unit or not point.account or accounts.get(point.account) != point:
            audit["invalid_catalog_points"] += 1
            continue
        ptm = amount_by_point[point]
        reuse = reuse_by_point[point]
        hq_revenue = ptm + reuse * REUSE_WEIGHT[point.channel]
        detail.append(dict(zip(POINT_FIELDS, (
            point.unit, point.channel, point.code, point.account, 1,
            int(hq_revenue > 0), int(production[point] > 0), ptm,
            reuse if (args.td49 if point.channel == "VINAPRO" else args.sim_tsd) else None,
            hq_revenue, production[point],
        ))))
    return detail, exceptions, audit, period_phones


def summarize(detail):
    totals = {}
    for row in detail:
        for group in ((row["unit"], row["channel"]), ("TOTAL", row["channel"])):
            result = totals.setdefault(group, {name: 0 for name in POINT_FIELDS[4:]})
            for name in result:
                value = row[name]
                if value is None:
                    result[name] = None
                else:
                    result[name] += value
    return totals


def compare(path, totals):
    comparisons = []
    for number, row in enumerate(workbook_rows(path, "HQ"), 1):
        if not 5 <= number <= 23:
            continue
        unit = key(row[0]) if number != 23 else "TOTAL"
        for channel, fields in REFERENCE_COLUMNS.items():
            current = totals.get((unit, channel), {})
            for metric, column in fields.items():
                computed = current.get(metric, 0)
                expected = row[column]
                comparisons.append({
                    "unit": unit, "channel": channel, "metric": metric,
                    "computed": computed, "reference": expected,
                    "difference": computed - money(expected) if isinstance(expected, (int, float)) else None,
                })
    return comparisons


def reference_bris(path, detail, period_phones):
    """Kiểm tra riêng khóa BRIS đã chuẩn hóa trong workbook, giữ nguyên kết quả BRIS thô."""
    accounts = Counter()
    coverage = Counter()
    for number, row in enumerate(workbook_rows(path, "P01_SL PTM"), 1):
        if number < 5 or not key(row[6]):
            continue
        accounts[key(row[11])] += 1
        coverage["reference_bris_rows"] += 1
        coverage["reference_bris_phone_in_raw_period"] += int(key(row[6]) in period_phones)
    adjusted = [
        dict(row, production=accounts[row["account"]],
             production_points=int(accounts[row["account"]] > 0))
        for row in detail
    ]
    return adjusted, coverage


def write_csv(path, rows, fields):
    with path.open("w", encoding="utf-8-sig", newline="") as output:
        writer = csv.DictWriter(output, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--ims", type=Path, required=True)
    parser.add_argument("--ims-sheet", default="temp1")
    parser.add_argument("--bris", type=Path, required=True)
    parser.add_argument("--bris-sheet", default="Sheet1")
    parser.add_argument("--bris-from", type=date.fromisoformat)
    parser.add_argument("--bris-to", type=date.fromisoformat)
    source = parser.add_mutually_exclusive_group(required=True)
    source.add_argument("--smcs", type=Path, help="Danh mục SMCS đã bổ sung HRM và mã đơn vị")
    source.add_argument("--week46-catalog", type=Path, help="Danh mục thay thế chỉ dùng thử tuần 46")
    parser.add_argument("--smcs-sheet")
    parser.add_argument("--td49", type=Path, help="Workbook có sheet ALL_TD49")
    parser.add_argument("--sim-tsd", type=Path, help="Workbook có sheet Sim TSD")
    parser.add_argument("--compare", type=Path, help="Workbook có HQ mẫu tuần 46")
    parser.add_argument("--diagnose-reference-bris", action="store_true")
    parser.add_argument("--output", type=Path, required=True, help="Thư mục đầu ra mới")
    args = parser.parse_args()
    if args.bris_from and args.bris_to and args.bris_from > args.bris_to:
        parser.error("--bris-from phải trước hoặc bằng --bris-to")
    if args.diagnose_reference_bris and not args.compare:
        parser.error("--diagnose-reference-bris yêu cầu --compare")
    points = set(catalog(args.smcs, args.smcs_sheet) if args.smcs
                 else reference_catalog(args.week46_catalog))
    if not points:
        parser.error("Danh mục không có điểm VinaPro hoặc VinaKey1")
    detail, exceptions, audit, period_phones = calculate(args, points)
    if not detail:
        parser.error("Danh mục không có điểm đủ mã tài khoản và đơn vị")
    totals = summarize(detail)
    report = [dict(unit=unit, channel=channel, **values)
              for (unit, channel), values in sorted(totals.items())]
    for row in report:
        row["revenue_point_rate"] = Decimal(row["revenue_points"]) / row["points"]
        row["production_point_rate"] = Decimal(row["production_points"]) / row["points"]
    args.output.mkdir(parents=True, exist_ok=False)
    write_csv(args.output / "hq.csv", report,
              ("unit", "channel", *POINT_FIELDS[4:], "revenue_point_rate", "production_point_rate"))
    write_csv(args.output / "points.csv", detail, POINT_FIELDS)
    write_csv(args.output / "exceptions.csv", exceptions, ("source", "row", "reason"))
    metadata = {
        "catalog": "smcs" if args.smcs else "week46_reference",
        "ims_period": "as_exported", "bris_from": args.bris_from, "bris_to": args.bris_to,
        "reuse_sources": {"VINAPRO": bool(args.td49), "VINAKEY1": bool(args.sim_tsd)},
        "counts": dict(audit),
    }
    fields = ("unit", "channel", "metric", "computed", "reference", "difference")
    if args.compare:
        write_csv(args.output / "comparison.csv", compare(args.compare, totals), fields)
    if args.diagnose_reference_bris:
        adjusted, coverage = reference_bris(args.compare, detail, period_phones)
        metadata["reference_bris_diagnostic"] = dict(coverage)
        write_csv(args.output / "comparison_reference_bris.csv",
                  compare(args.compare, summarize(adjusted)), fields)
    (args.output / "run.json").write_text(
        json.dumps(metadata, ensure_ascii=False, indent=2, default=str) + "\n", encoding="utf-8",
    )
    print(f"Đã ghi kết quả vào {args.output}")


if __name__ == "__main__":
    main()
