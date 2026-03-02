#!/usr/bin/env python3
"""Convert 用户信息.xlsx to OTR audit template format.
No third-party deps; works by editing XLSX XML directly.
"""
import argparse
import datetime as dt
import re
import zipfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def get_sheet_path(zf: zipfile.ZipFile) -> str:
    wb = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    relmap = {r.attrib["Id"]: r.attrib["Target"] for r in rels}
    s = wb.find(f".//{{{NS}}}sheet")
    return "xl/" + relmap[s.attrib[f"{{{RNS}}}id"]]


def read_user_records(user_xlsx: str, dedupe: bool):
    zf = zipfile.ZipFile(user_xlsx)
    sheet = ET.fromstring(zf.read(get_sheet_path(zf)))
    rows = sheet.findall(f".//{{{NS}}}sheetData/{{{NS}}}row")
    items = []

    for row in rows[1:]:
        vals = {}
        for c in row.findall(f"{{{NS}}}c"):
            ref = c.attrib.get("r", "")
            m = re.match(r"[A-Z]+", ref)
            if not m:
                continue
            col = m.group(0)
            t = c.attrib.get("t")
            txt = ""
            if t == "inlineStr":
                txt = "".join((x.text or "") for x in c.findall(f".//{{{NS}}}t"))
            else:
                v = c.find(f"{{{NS}}}v")
                if v is not None:
                    txt = v.text or ""
            vals[col] = txt.strip()

        uid = vals.get("A", "").strip()
        if not uid or uid == "GEMS_ID":
            continue
        name = (vals.get("E", "") + vals.get("F", "")).strip() or (
            vals.get("C", "") + " " + vals.get("D", "")
        ).strip()
        items.append((uid, name))

    if not dedupe:
        return items

    seen = set()
    out = []
    for uid, name in items:
        if uid in seen:
            continue
        seen.add(uid)
        out.append((uid, name))
    return out


def render_with_template(template_xlsx: str, records, dealer_id: str, dealer_name: str, out_xlsx: str):
    zf = zipfile.ZipFile(template_xlsx, "r")
    sheet_path = get_sheet_path(zf)
    xml = zf.read(sheet_path).decode("utf-8")

    # Plain output (no template header/style/formula reuse)
    header = (
        '<row r="1" spans="1:8">'
        '<c r="A1" t="inlineStr"><is><t>用户ID</t></is></c>'
        '<c r="B1" t="inlineStr"><is><t>用户姓名</t></is></c>'
        '<c r="C1" t="inlineStr"><is><t>DIC是否在职</t></is></c>'
        '<c r="D1" t="inlineStr"><is><t>是否使用OTR+</t></is></c>'
        '<c r="E1" t="inlineStr"><is><t>是否已审查DIC和OTR+岗位匹配关系并做了相关调整</t></is></c>'
        '<c r="F1" t="inlineStr"><is><t>是否已审查用户OTR+权限并做了相关调整</t></is></c>'
        '<c r="G1" t="inlineStr"><is><t>经销商ID</t></is></c>'
        '<c r="H1" t="inlineStr"><is><t>经销商名称</t></is></c>'
        '</row>'
    )

    body = []
    for i, (uid, name) in enumerate(records, start=2):
        body.append(
            f'<row r="{i}" spans="1:8">'
            f'<c r="A{i}" t="inlineStr"><is><t>{escape(uid)}</t></is></c>'
            f'<c r="B{i}" t="inlineStr"><is><t>{escape(name)}</t></is></c>'
            f'<c r="C{i}" t="inlineStr"><is><t>是</t></is></c>'
            f'<c r="D{i}" t="inlineStr"><is><t>是</t></is></c>'
            f'<c r="E{i}" t="inlineStr"><is><t>是</t></is></c>'
            f'<c r="F{i}" t="inlineStr"><is><t>是</t></is></c>'
            f'<c r="G{i}" t="inlineStr"><is><t>{escape(dealer_id)}</t></is></c>'
            f'<c r="H{i}" t="inlineStr"><is><t>{escape(dealer_name)}</t></is></c>'
            f'</row>'
        )

    new_sheet = "<sheetData>" + header + "".join(body) + "</sheetData>"
    xml = re.sub(r"<sheetData>.*?</sheetData>", new_sheet, xml, flags=re.S)
    last = 1 + len(records)
    xml = re.sub(r'<dimension ref="[^"]+"\s*/>', f'<dimension ref="A1:H{last}" />', xml)

    # Remove calcChain refs to avoid Excel repair popups after XML edit
    content_types = zf.read("[Content_Types].xml").decode("utf-8")
    content_types = re.sub(
        r'<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain\+xml"\s*/>',
        "",
        content_types,
    )
    wb_rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    wb_rels = re.sub(
        r'<Relationship[^>]*Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain"[^>]*/>',
        "",
        wb_rels,
    )

    with zipfile.ZipFile(out_xlsx, "w", zipfile.ZIP_DEFLATED) as out:
        for info in zf.infolist():
            if info.filename == "xl/calcChain.xml":
                continue
            data = zf.read(info.filename)
            if info.filename == sheet_path:
                data = xml.encode("utf-8")
            elif info.filename == "[Content_Types].xml":
                data = content_types.encode("utf-8")
            elif info.filename == "xl/_rels/workbook.xml.rels":
                data = wb_rels.encode("utf-8")
            out.writestr(info, data)


def default_output_name():
    now = dt.datetime.now()
    return f"OTR 用户审核表{now.year}年{now.month:02d}月.xlsx"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--user", required=True, help="用户信息.xlsx")
    ap.add_argument("--template", required=True, help="OTR模板.xlsx")
    ap.add_argument("--out", default=default_output_name(), help="输出文件路径")
    ap.add_argument("--dedupe", action="store_true", default=True, help="按用户ID去重(默认开)")
    ap.add_argument("--no-dedupe", action="store_true", help="关闭去重")
    ap.add_argument("--dealer-id", default="SFSZ10")
    ap.add_argument("--dealer-name", default="佛山中升睿之星汽车销售服务有限公司")
    args = ap.parse_args()

    dedupe = False if args.no_dedupe else True
    records = read_user_records(args.user, dedupe=dedupe)
    render_with_template(args.template, records, args.dealer_id, args.dealer_name, args.out)
    print(f"OK: {args.out} (rows={len(records)})")


if __name__ == "__main__":
    main()
