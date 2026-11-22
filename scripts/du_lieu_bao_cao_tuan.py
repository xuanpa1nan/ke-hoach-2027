"""Chuẩn bị số liệu HQ/BSC và cấu trúc báo cáo tuần để xuất tệp Office."""

import argparse
import csv
import json
from collections import Counter, defaultdict
from pathlib import Path

from tinh_hq import key, money, table, workbook_rows


def ratio(numerator, denominator):
    return numerator / denominator if denominator else None


def score(value, low, high, offset):
    if value is None:
        return None
    if value < low:
        return 1
    return 10 * value - offset if value < high else 5


def numeric(value):
    return value if isinstance(value, (int, float)) else None


def display(value):
    return "Thiếu dữ liệu" if value is None else value


def report(hq_dir, reference, template, ims):
    with (hq_dir / "hq.csv").open(encoding="utf-8-sig") as stream:
        hq = {(r["unit"], r["channel"]): r for r in csv.DictReader(stream)}
    meta = json.loads((hq_dir / "run.json").read_text(encoding="utf-8"))
    bsc = list(workbook_rows(reference, "BSC"))
    reference_hq = list(workbook_rows(reference, "HQ"))
    units = [key(r[0]) for r in reference_hq[4:22] if key(r[0])]
    units += sorted({unit for unit, _ in hq} - set(units) - {"TOTAL"})
    units.append("TOTAL")
    missing = [
        ["SMCS", "Danh mục điểm dùng bản tuần 46 trích từ workbook. Chưa có tệp SMCS xuất trực tiếp."],
        ["CTV / BSC I.8.1", "Tính lại từ các bảng CTV XHH, All_CTV XHH và danh sách nhân viên lưu trong workbook tuần 46."],
        ["Kế hoạch BSC", "Dùng giá trị lưu sẵn ở cột E, K và N của BSC tuần 46. Chưa đối chiếu tệp kế hoạch gốc."],
        ["Điều kiện chi lương", "Chưa tính lại từ bảng định mức và quy chế. Các ô tương ứng được đánh dấu thiếu."],
        ["So sánh kỳ trước", "Thiếu bộ số tuần 45 và tiến độ kế hoạch. Chưa viết nhận định tăng/giảm theo kỳ trước."],
        ["Phạm vi báo cáo tuần", "Các dịch vụ ngoài quản trị kênh, cùng các chỉ tiêu chiến dịch 120 ngày và DigiShop, cần nguồn riêng."],
        ["Đối soát TD49", "Tính VinaPro trên toàn bộ tệp IMS. Bản mẫu tuần 46 chỉ có công thức tra TD49 đến dòng 5091."],
        ["Tái sử dụng VinaKey1", "Sheet Sim TSD của tuần 46 chưa có dòng đủ khóa và giá gói để kiểm chứng doanh thu tái sử dụng."],
        ["BRIS", f"Lọc theo ngày kích hoạt từ {meta.get('bris_from')} đến {meta.get('bris_to')}. Dùng khóa BRIS thô."],
        ["CTKD", "Doanh thu VinaPro/VinaKey1 tính từ IMS. Kế hoạch, điểm đạt chuẩn và các chỉ tiêu CTKD khác lấy số lưu trong sheet TH tuần 46."],
        ["Chiến dịch 120 ngày", "Tệp mang tên tuần 46 có doanh thu tuần trùng bản tuần 47 tại 19 đơn vị sau làm tròn. Cần xác nhận kỳ trước khi điền."],
    ]
    ctv_orders = Counter(key(r[13]) for r in list(workbook_rows(reference, "CTV XHH"))[1:] if key(r[13]))
    managed = defaultdict(set)
    for row in list(workbook_rows(reference, "All_CTV XHH"))[1:]:
        if key(row[7]) and key(row[3]):
            managed[key(row[7])].add(key(row[3]))
    staff = defaultdict(dict)
    for row in bsc[2:]:
        unit, employee = key(row[20]), key(row[21])
        if unit and employee:
            accounts = managed[employee]
            share = ratio(sum(bool(ctv_orders[account]) for account in accounts), len(accounts))
            staff[unit][employee] = share is not None and share >= .20
    plans = {key(row[0]): (numeric(row[4]), numeric(row[10]), numeric(row[13]))
             for row in bsc[3:21] if key(row[0])}
    plans["TOTAL"] = tuple(numeric(bsc[21][i]) for i in (4, 10, 13))
    # Mỗi cột đích D:M, N, R:Y giữ định nghĩa từ nhóm CTKD tương ứng trong TH.
    th_columns = dict(zip(range(3, 13), range(8, 19)))
    th_columns.update({7: 13, 8: 14, 9: 15, 10: 16, 11: 17, 12: 18})
    th_columns.update(dict(zip((13, *range(17, 25)), (20, *range(24, 32)))))
    th_rows = list(workbook_rows(reference, "TH"))
    th_by_unit = {key(row[1]): row for row in th_rows[5:24] if key(row[1])}
    th_by_unit["TOTAL"] = th_rows[24]
    transaction_ptm = defaultdict(float)
    ims_columns = {"unit": {"DONVI"}, "channel": {"DOITUONG"},
                   "kind": {"LOAITB"}, "amount": {"GIAGOI"}}
    for _, row in table(workbook_rows(ims, "temp1"), ims_columns, "IMS73"):
        channel, unit = key(row["channel"]), key(row["unit"])
        if key(row["kind"]) == "KHM" and channel in {"VINAPRO", "VINAKEY1"}:
            amount = float(money(row["amount"]))
            transaction_ptm[unit, channel] += amount
            transaction_ptm["TOTAL", channel] += amount

    def measure(unit, channel, field):
        value = hq.get((unit, channel), {}).get(field)
        return float(value) if value else 0

    hq_rows, bsc_rows, program = [], [], []
    for unit in units:
        pro = [measure(unit, "VINAPRO", f) for f in (
            "points", "revenue_points", "hq_revenue", "production_points", "production", "ptm_revenue", "reuse_revenue")]
        vk = [measure(unit, "VINAKEY1", f) for f in (
            "points", "revenue_points", "hq_revenue", "production_points", "production", "ptm_revenue", "reuse_revenue")]
        label = "Toàn tỉnh" if unit == "TOTAL" else unit
        hq_rows.append([label, pro[5]+vk[5], pro[6]+vk[6], pro[4]+vk[4],
                        pro[0], pro[1], pro[2], pro[3], pro[4], "Thiếu dữ liệu",
                        vk[0], vk[1], vk[2], vk[3], vk[4], "Thiếu dữ liệu"])
        people = list(staff[unit].values()) if unit != "TOTAL" else [
            valid for group in staff.values() for valid in group.values()
        ]
        plan1, plan2, plan3 = plans.get(unit, (None, None, None))
        active_points, all_points = pro[3]+vk[3], pro[0]+vk[0]
        revenue = (pro[2]+vk[2])/1_000_000
        f1 = ratio(sum(people)/len(people), plan1) if people else None
        f2 = ratio(active_points/all_points, plan2) if all_points else None
        f3 = ratio(revenue, plan3)
        scores = [score(f1,.3,.7,2), score(f2,.8,1,5), score(f3,.8,1,5)]
        weighted = sum(s*w for s,w in zip(scores, (.3,.3,.4))) if None not in scores and unit != "TOTAL" else None
        bsc_rows.append([label, len(people), sum(people), plan1, f1, scores[0],
                         all_points, active_points, plan2, f2, scores[1],
                         plan3, revenue, f3, scores[2], weighted if unit != "TOTAL" else "Không chấm dòng tổng"])
        # KHG BSC doanh thu có phần tái sử dụng, nên giữ riêng khỏi PTM của CT2.1.
        pro_transaction = transaction_ptm[unit, "VINAPRO"]
        vk_transaction = transaction_ptm[unit, "VINAKEY1"]
        th = th_by_unit.get(unit)
        cached = {column: numeric(th[source]) for column, source in th_columns.items()} if th else {}
        program.append([label, pro_transaction, vk_transaction, pro_transaction+vk_transaction, cached])

    template_tabs = []
    from openpyxl import load_workbook
    book = load_workbook(template, read_only=True, data_only=True, keep_links=False)
    try:
        template_tabs = book.sheetnames
        template_headers = {
            name: list(book[name].iter_rows(max_row=height, values_only=True))
            for name, height in (("KQTH Kênh PRO, KEY1", 4), ("CTKD DT DD và KBH", 5))
        }
    finally:
        book.close()
    return {"hq": hq_rows, "bsc": [[display(v) for v in r] for r in bsc_rows],
            "program": program, "notes": missing, "template_tabs": template_tabs,
            "template_headers": template_headers}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--hq-dir", type=Path, required=True)
    parser.add_argument("--reference", type=Path, required=True)
    parser.add_argument("--ims", type=Path, required=True)
    parser.add_argument("--excel-template", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--week", type=int, required=True)
    args = parser.parse_args()
    data = report(args.hq_dir, args.reference, args.excel_template, args.ims)
    data["week"] = args.week
    args.output.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    main()
