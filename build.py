"""JSON テンプレートと素材画像から pptx を生成する。

使い方:
    uv run build.py ポスター.json

もとの pptx（"base"）を開き、JSON に書かれた要素だけを変更して "output" に保存する。
もとの pptx は上書きしない。

- image: "path" の画像で差し替える。画像は元の図形の中身だけを入れ替えるので、
  重なり順・枠線・影などの書式は残る。"path" が未指定またはファイルがない場合は、
  点線の枠を残す。
  "fit": "contain"（枠内に収める・縦横比維持）/ "cover"（枠を埋めてはみ出しをトリミング）
         / "stretch"（枠に合わせて伸縮）
- text: "text" の段落リストが元と異なる段落だけ書き換える。書き換えた段落は
  その段落の先頭の文字書式が全体に適用される。段落内改行は "\n"。
- box: [left, top, width, height]（cm）が元と異なれば移動・リサイズする。
"""

import argparse
import copy
import json
import re
import sys
from pathlib import Path

from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.shapes.shapetree import SlideShapeFactory
from pptx.text.text import Font
from pptx.util import Pt

from template_common import (
    cm_to_emu,
    get_box,
    get_paragraphs,
    is_empty_picture_placeholder,
    iter_shapes,
)

FRAME_COLOR = RGBColor(0x99, 0x99, 0x99)


def warn(message):
    print(f"  [警告] {message}", file=sys.stderr)


# ---------------------------------------------------------------- 共通


def apply_box(shape, box):
    if box is None or box == get_box(shape):
        return
    shape.left, shape.top, shape.width, shape.height = (cm_to_emu(v) for v in box)


def drop_unused_rels(part, rIds):
    """どこからも参照されなくなったリレーションを削除し、不要な画像を保存対象から外す。"""
    used = set(part._element.xpath("//@r:embed | //@r:link | //@r:id"))
    for rId in rIds:
        if rId and rId not in used and rId in part.rels:
            part.rels.pop(rId)


# ---------------------------------------------------------------- 画像


def blip_rIds(pic_element):
    return [v for v in pic_element.xpath(".//@r:embed | .//@r:link")]


def replace_picture(shape, image_path):
    """p:pic の画像だけを差し替える（位置・書式・重なり順は維持）。"""
    pic = shape._element
    old_rIds = blip_rIds(pic)
    _, rId = shape.part.get_or_add_image_part(str(image_path))
    blip = pic.blipFill.blip
    blip.rEmbed = rId
    # SVG 画像は extLst 内の svgBlip が優先表示されるので消す
    for ext_lst in blip.findall(qn("a:extLst")):
        blip.remove(ext_lst)
    drop_unused_rels(shape.part, old_rIds)


def apply_fit(picture, image_path, fit):
    picture.crop_left = picture.crop_right = picture.crop_top = picture.crop_bottom = 0
    if fit == "stretch":
        return

    with Image.open(image_path) as img:
        img_w, img_h = img.size
    img_ratio = img_w / img_h
    left, top, width, height = picture.left, picture.top, picture.width, picture.height
    box_ratio = width / height

    if fit == "contain":
        if box_ratio > img_ratio:
            new_width = round(height * img_ratio)
            picture.left = left + (width - new_width) // 2
            picture.width = new_width
        else:
            new_height = round(width / img_ratio)
            picture.top = top + (height - new_height) // 2
            picture.height = new_height
    elif fit == "cover":
        if box_ratio > img_ratio:
            crop = (1 - img_ratio / box_ratio) / 2
            picture.crop_top = picture.crop_bottom = crop
        else:
            crop = (1 - box_ratio / img_ratio) / 2
            picture.crop_left = picture.crop_right = crop
    else:
        warn(f'fit "{fit}" は不明です（contain / cover / stretch）。stretch として扱います')


def add_frame(slide, left, top, width, height, name, label):
    """「画像未配置」の点線の枠をスライド最前面に追加する。"""
    frame = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    frame.name = name
    frame.fill.background()
    frame.line.color.rgb = FRAME_COLOR
    frame.line.dash_style = MSO_LINE.DASH
    frame.line.width = Pt(1.5)
    frame.shadow.inherit = False
    tf = frame.text_frame
    tf.text = f"画像未配置\n{label}"
    font_size = Pt(max(18, min(frame.width, frame.height) / 12 / 12700))  # 枠の大きさに合わせる
    for p in tf.paragraphs:
        p.alignment = PP_ALIGN.CENTER
        for r in p.runs:
            r.font.color.rgb = FRAME_COLOR
            r.font.size = font_size
    return frame


def replace_with_frame(slide, shape, label):
    """画像がない場合、画像を消して同じ位置に点線の枠を残す。"""
    frame = add_frame(slide, shape.left, shape.top, shape.width, shape.height, shape.name, label)
    old_rIds = blip_rIds(shape._element)
    # add_shape はスライド最前面に追加されるので、元の画像の位置（グループ内も含む）へ移す
    shape._element.addprevious(frame._element)
    shape._element.getparent().remove(shape._element)
    drop_unused_rels(slide.part, old_rIds)


def build_image(slide, shape, element, image_root):
    name = shape.name
    apply_box(shape, element.get("box"))
    path = element.get("path")
    image_path = (image_root / path) if path else None

    if image_path is None or not image_path.is_file():
        if image_path is not None:
            warn(f"{shape.name}: 画像が見つかりません: {image_path}")
        else:
            print(f"  {shape.name}: 画像未指定のため枠を残します")
        if not is_empty_picture_placeholder(shape):
            replace_with_frame(slide, shape, path or shape.name)
        return

    if is_empty_picture_placeholder(shape):
        box = (shape.left, shape.top, shape.width, shape.height)
        picture = shape.insert_picture(str(image_path))
        # 挿入直後の画像は位置をレイアウトから継承しているので、明示的に書き込む
        picture.left, picture.top, picture.width, picture.height = box
    else:
        replace_picture(shape, image_path)
        picture = shape
    apply_fit(picture, image_path, element.get("fit", "contain"))
    print(f"  {name}: {path}")


# ---------------------------------------------------------------- テキスト


# 段落内の書式指定: <sup>…</sup>（上付き）、<color=RRGGBB>…</color>（文字色）
MARKUP = re.compile(r"<sup>(.*?)</sup>|<color=([0-9A-Fa-f]{6})>(.*?)</color>")


def split_markup(line):
    """(文字列, 上付きか, 色) の並びにする。"""
    pos = 0
    for m in MARKUP.finditer(line):
        if m.start() > pos:
            yield line[pos:m.start()], False, None
        if m.group(1) is not None:
            yield m.group(1), True, None
        else:
            yield m.group(3), False, m.group(2).upper()
        pos = m.end()
    if pos < len(line):
        yield line[pos:], False, None


def line_formats(p):
    """段落内の行（Shift+Enter 区切り）ごとに、先頭の文字書式 a:rPr を返す。"""
    formats, current = [], None
    for child in p.content_children:
        if child.tag == qn("a:br"):
            formats.append(current)
            current = None
        elif current is None and child.find(qn("a:rPr")) is not None:
            current = child.find(qn("a:rPr"))
    formats.append(current)
    fallback = next((f for f in formats if f is not None), None)
    return [f if f is not None else fallback for f in formats]


def set_paragraph_text(p, text):
    """段落の文字列を置き換える。

    書式は行ごと（段落内改行 \n で区切った行）に、もとの同じ行の先頭の文字書式を引き継ぐ
    （例: 1 行目が和文タイトル 75pt、2 行目が英文タイトル 43pt）。行がもとより多ければ最後の行の書式。
    <sup>…</sup> で囲んだ部分は上付き文字、<color=RRGGBB>…</color> はその色にする。
    """
    formats = line_formats(p)
    for child in p.content_children:
        p.remove(child)
    for i, line in enumerate(text.split("\n")):
        if i > 0:
            p.add_br()
        rPr = formats[min(i, len(formats) - 1)]
        for part, sup, color in split_markup(line):
            if not part:
                continue
            r = p.add_r()
            r.text = part
            if rPr is not None:
                r.insert(0, copy.deepcopy(rPr))
            if sup:
                r.get_or_add_rPr().set("baseline", "30000")
            elif r.rPr is not None and "baseline" in r.rPr.attrib:
                del r.rPr.attrib["baseline"]
            if color:
                font = Font(r.get_or_add_rPr())
                font.color.rgb = RGBColor.from_string(color)


def scale_fonts(shape, scale):
    """文字の大きさ（sz）を変える。

    scale が数値なら全体を scale 倍、{"28": 32, "36": 40} のような対応表なら、その大きさ（pt）の文字だけ置き換える。
    """
    for node in shape.text_frame._txBody.iter(qn("a:rPr"), qn("a:endParaRPr"), qn("a:defRPr")):
        sz = node.get("sz")
        if not sz:
            continue
        if isinstance(scale, dict):
            pt = int(sz) / 100
            new = scale.get(f"{pt:g}")
            if new is not None:
                node.set("sz", str(round(new * 100)))
        else:
            node.set("sz", str(round(int(sz) * scale / 100) * 100))


def build_text(shape, element):
    apply_box(shape, element.get("box"))
    if element.get("font_scale"):
        scale_fonts(shape, element["font_scale"])
    new_texts = element.get("text")
    if new_texts is None:
        return
    if isinstance(new_texts, str):
        new_texts = [new_texts]
    if new_texts == get_paragraphs(shape):
        return

    txBody = shape.text_frame._txBody
    ps = txBody.p_lst
    # 段落数を合わせる（増やすときは最後の段落の書式を複製）
    while len(ps) < len(new_texts):
        ps[-1].addnext(copy.deepcopy(ps[-1]))
        ps = txBody.p_lst
    for extra in ps[len(new_texts):]:
        txBody.remove(extra)
    ps = txBody.p_lst

    for p, text in zip(ps, new_texts):
        if p.text.replace("\v", "\n") != text:
            set_paragraph_text(p, text)
    print(f"  {shape.name}: テキストを更新")


# ---------------------------------------------------------------- 新しい要素


def clone_shape(slide, source, name):
    """source を複製してスライド最前面（グループの外）に置く。書式はそのまま引き継ぐ。"""
    element = copy.deepcopy(source._element)
    c_nv_pr = element.xpath("./*[1]/p:cNvPr")[0]
    c_nv_pr.set("id", str(slide.shapes._next_shape_id))
    c_nv_pr.set("name", name or source.name)
    slide.shapes._spTree.insert_element_before(element, "p:extLst")
    return SlideShapeFactory(element, slide.shapes)


def add_element(slide, element, shapes, image_root):
    """"id" のない要素を新しく追加する。"copy_from" があればその図形の書式を複製する。"""
    name = element.get("name")
    if element["type"] == "arrow":  # 矢印は box ではなく from / to で位置を決める
        add_arrow(slide, element)
        return
    box = element.get("box")
    if box is None:
        warn(f"{name or '新しい要素'}: 追加する要素には box が必要です")
        return

    source_id = element.get("copy_from")
    if source_id is not None:
        source = shapes.get(source_id)
        if source is None:
            warn(f"{name or '新しい要素'}: copy_from の id {source_id} の図形が見つかりません")
            return
        shape = clone_shape(slide, source, name)
        if element["type"] == "image":
            build_image(slide, shape, element, image_root)
        elif element["type"] == "text":
            build_text(shape, element)
        else:  # "shape": 線など。位置・サイズだけ変える
            apply_box(shape, box)
            print(f"  {shape.name}: 図形を追加")
        return

    if element["type"] == "shape":
        warn(f"{name or '新しい要素'}: type \"shape\" には copy_from が必要です")
        return
    if element["type"] == "table":
        add_table(slide, element)
        return
    left, top, width, height = (cm_to_emu(v) for v in box)
    if element["type"] == "image":
        path = element.get("path")
        image_path = (image_root / path) if path else None
        if image_path is None or not image_path.is_file():
            if image_path is not None:
                warn(f"{name}: 画像が見つかりません: {image_path}")
            add_frame(slide, left, top, width, height, name or "画像", path or name or "")
            print(f"  {name}: 画像未配置の枠を追加")
            return
        picture = slide.shapes.add_picture(str(image_path), left, top, width, height)
        if name:
            picture.name = name
        apply_fit(picture, image_path, element.get("fit", "contain"))
        print(f"  {picture.name}: {path} を追加")
    else:
        textbox = slide.shapes.add_textbox(left, top, width, height)
        if name:
            textbox.name = name
        texts = element.get("text") or [""]
        textbox.text_frame.text = "\n".join([texts] if isinstance(texts, str) else texts)
        print(f"  {textbox.name}: テキストを追加")


def add_arrow(slide, element):
    """矢印（直線コネクタ）を追加する。"from" の点から "to" の点へ向く（cm）。

    "color": 線の色（RRGGBB、既定 FFFF00）、"width_pt": 線の太さ（既定 6）
    """
    (x1, y1), (x2, y2) = element["from"], element["to"]
    line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, cm_to_emu(x1), cm_to_emu(y1), cm_to_emu(x2), cm_to_emu(y2))
    if element.get("name"):
        line.name = element["name"]
    line.line.color.rgb = RGBColor.from_string(element.get("color", "FFFF00"))
    line.line.width = Pt(element.get("width_pt", 6))
    ln = line.line._get_or_add_ln()
    tail = ln.find(qn("a:tailEnd"))
    if tail is None:
        tail = etree.SubElement(ln, qn("a:tailEnd"))
    tail.set("type", "triangle")
    print(f"  {line.name}: 矢印を追加")


def iter_groups(shapes):
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield shape
            yield from iter_groups(shape.shapes)


# 罫線だけの表スタイル（PowerPoint の「スタイルなし、表のグリッド線」）
TABLE_STYLE_GRID = "{5940675A-B579-460E-94D1-54222C63F5DA}"


def add_table(slide, element):
    """表を追加する。

    "rows": 行ごとの文字列のリスト（1 行目は見出し）。null のセルは上のセルと結合する
    "font_size": 文字の大きさ（pt、既定 28）
    "col_widths": 列幅の比（省略時は均等）
    "header_fill": 見出し行の塗り（RRGGBB、既定 D9D9D9）
    """
    rows = element["rows"]
    n_rows, n_cols = len(rows), max(len(r) for r in rows)
    left, top, width, height = (cm_to_emu(v) for v in element["box"])
    frame = slide.shapes.add_table(n_rows, n_cols, left, top, width, height)
    if element.get("name"):
        frame.name = element["name"]
    table = frame.table
    table._tbl.tblPr.find(qn("a:tableStyleId")).text = TABLE_STYLE_GRID
    table.first_row = True
    table.horz_banding = False

    ratios = element.get("col_widths") or [1] * n_cols
    for col, ratio in zip(table.columns, ratios):
        col.width = int(width * ratio / sum(ratios))
    for row in table.rows:
        row.height = int(height / n_rows)

    font_size = Pt(element.get("font_size", 28))
    header_fill = RGBColor.from_string(element.get("header_fill", "D9D9D9"))
    for r, values in enumerate(rows):
        for c in range(n_cols):
            value = values[c] if c < len(values) else None
            cell = table.cell(r, c)
            if value is None:
                continue
            cell.text = str(value)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            if r == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = header_fill
            else:
                cell.fill.background()
            for p in cell.text_frame.paragraphs:
                p.alignment = PP_ALIGN.CENTER
                for run in p.runs:
                    run.font.size = font_size
                    run.font.bold = r == 0
                    run.font.color.rgb = RGBColor(0, 0, 0)
    # null のセルは、上の値のあるセルと結合する（例: 群名を縦に結合）
    for c in range(n_cols):
        r = 1
        while r < n_rows:
            if c < len(rows[r]) and rows[r][c] is None:
                end = r
                while end + 1 < n_rows and c < len(rows[end + 1]) and rows[end + 1][c] is None:
                    end += 1
                table.cell(r - 1, c).merge(table.cell(end, c))
                r = end + 1
            else:
                r += 1
    print(f"  {frame.name}: 表を追加（{n_rows}×{n_cols}）")


def delete_shape(slide, element):
    """id の図形を消す。グループの id なら中身ごと消す。"""
    found = slide.shapes._spTree.xpath(f'.//p:cNvPr[@id="{element["id"]}"]/../..')
    if not found:
        warn(f'id {element["id"]}（{element.get("name")}）の図形が見つかりません')
        return
    target = found[0]
    rIds = target.xpath(".//@r:embed | .//@r:link")
    name = target.xpath("./*[1]/p:cNvPr/@name")[0]
    target.getparent().remove(target)
    drop_unused_rels(slide.part, rIds)
    print(f"  {name}: 削除")


# ---------------------------------------------------------------- メイン


def build(json_path, output=None):
    template = json.loads(json_path.read_text(encoding="utf-8"))
    json_dir = json_path.parent
    base_path = json_dir / template["base"]
    # -o があれば JSON の "output" より優先（別の PC で JSON を共有するとき用）
    output_path = output.resolve() if output else json_dir / template["output"]
    image_root = json_dir / template.get("image_root", ".")

    if template.get("unit", "cm") != "cm":
        sys.exit('エラー: "unit" は "cm" のみ対応しています')
    if base_path.resolve() == output_path.resolve():
        sys.exit('エラー: "output" が "base" と同じです。もとの pptx を上書きしないよう別名にしてください')

    prs = Presentation(base_path)
    for slide_spec in template["slides"]:
        index = slide_spec["index"]
        slide = prs.slides[index]
        print(f"スライド {index + 1}")
        shapes = {shape.shape_id: shape for shape, _ in iter_shapes(slide.shapes)}
        groups = {shape.shape_id: shape for shape in iter_groups(slide.shapes)}

        for element in slide_spec["elements"]:
            if element.get("id") is None:
                add_element(slide, element, shapes, image_root)
                continue
            if element["type"] == "delete":
                delete_shape(slide, element)
                continue
            if element["type"] == "group":  # グループは位置・サイズだけ変える
                group = groups.get(element["id"])
                if group is None:
                    warn(f'id {element["id"]}（{element.get("name")}）のグループが見つかりません')
                else:
                    apply_box(group, element.get("box"))
                continue
            shape = shapes.get(element["id"])
            if shape is None:
                warn(f'id {element["id"]}（{element.get("name")}）の図形が見つかりません')
                continue
            if element["type"] == "image":
                build_image(slide, shape, element, image_root)
            elif element["type"] == "text":
                build_text(shape, element)
            else:
                warn(f'type "{element["type"]}" は不明です')

    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        prs.save(output_path)
    except PermissionError:
        sys.exit(f"エラー: {output_path} が開いたままです。閉じてから再実行してください")
    print(f"\n{output_path} に保存しました")


def main():
    parser = argparse.ArgumentParser(description="JSON テンプレートから pptx を生成する")
    parser.add_argument("json", type=Path, help="extract.py で作った JSON テンプレート")
    parser.add_argument("-o", "--output", type=Path, help='出力先（JSON の "output" より優先）')
    args = parser.parse_args()
    build(args.json.resolve(), args.output)


if __name__ == "__main__":
    main()
