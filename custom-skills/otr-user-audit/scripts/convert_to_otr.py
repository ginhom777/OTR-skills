#!/usr/bin/env python3
"""Convert 用户信息.xlsx to OTR audit template format.
No third-party deps; works by editing XLSX XML directly.
"""
import argparse
import datetime as dt
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from xml.sax.saxutils import escape

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def get_sheet_path(zf: zipfile.ZipFile) -> str:
    wb = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    relmap = {r.attrib["Id"]: r.attrib["Target"] for r in rels}
    s = wb.find(f".//{{{NS}}}sheet")
    return "xl/" + relmap[s.attrib[f"{{{RNS}}}id"]]


def read_shared_strings(zf: zipfile.ZipFile):
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    return ["".join((x.text or "") for x in si.findall(f".//{{{NS}}}t")) for si in root.findall(f"{{{NS}}}si")]


def cell_text(c, shared_strings):
    t = c.attrib.get("t")
    if t == "inlineStr":
        return "".join((x.text or "") for x in c.findall(f".//{{{NS}}}t")).strip()
    v = c.find(f"{{{NS}}}v")
    if v is None:
        return ""
    txt = (v.text or "").strip()
    if t == "s" and txt.isdigit():
        i = int(txt)
        if 0 <= i < len(shared_strings):
            return shared_strings[i].strip()
    return txt


def read_user_records(user_xlsx: str, dedupe: bool):
    zf = zipfile.ZipFile(user_xlsx)
    shared = read_shared_strings(zf)
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
            vals[m.group(0)] = cell_text(c, shared)

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


def default_template_path():
    return str(Path(__file__).resolve().parent.parent / "examples" / "OTR-template-new.xlsx")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--user", required=True, help="用户信息.xlsx")
    ap.add_argument("--template", default=default_template_path(), help="OTR模板.xlsx（可省略，默认使用examples内置模板）")
    ap.add_argument("--out", default=default_output_name(), help="输出文件路径")
    ap.add_argument("--dedupe", action="store_true", default=True, help="按用户ID去重(默认开)")
    ap.add_argument("--no-dedupe", action="store_true", help="关闭去重")
    ap.add_argument("--dealer-id", default="SFSZ10")
    ap.add_argument("--dealer-name", default="佛山中升睿之星汽车销售服务有限公司")
    args = ap.parse_args()

    if not Path(args.user).exists():
        raise SystemExit(f"ERROR: 用户信息文件不存在: {args.user}")
    if not Path(args.template).exists():
        raise SystemExit(f"ERROR: 模板文件不存在: {args.template}")

    dedupe = False if args.no_dedupe else True
    records = read_user_records(args.user, dedupe=dedupe)
    if not records:
        raise SystemExit("ERROR: 未读取到有效用户数据。请检查用户文件是否包含A列用户ID和姓名列（E/F或C/D）。")

    empty_ab = sum(1 for uid, name in records if (not uid.strip()) or (not name.strip()))
    if empty_ab:
        raise SystemExit(f"ERROR: 检测到{empty_ab}条记录的A/B可能为空，请检查源文件编码与列映射。")

    render_with_template(args.template, records, args.dealer_id, args.dealer_name, args.out)
    print(f"OK: {args.out} (rows={len(records)})")


if __name__ == "__main__":
    main()
