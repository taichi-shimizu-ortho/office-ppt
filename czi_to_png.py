"""明視野（H-E・MT など）の CZI を、ポスターに貼る PNG に書き出す。

使い方:
    uv run czi_to_png.py a.czi b.czi -o 出力フォルダ [--rotate 90]

処理の順番:
1. CZI を読む（ZEN の Bgr24 を RGB に並べ替える）
2. 縁の黒い行・列を切り落とす（撮影時に上や左が 1〜2 px 黒く残ることがある）
3. タイル境界に出る 1 px の黒い線を、上下（左右）の平均で埋める
4. 明るさを調整する
   - 白バランス: 各色の背景（--white-percentile、既定 99.5 パーセンタイル）を --white-level に合わせる
   - 黒点: 暗い側（輝度 0.5 パーセンタイル × --black）を 0 に寄せてコントラストをつける
   - ガンマ: 中間調を --gamma で締める（1 より大きいと暗く・色が濃くなる）
5. --rotate があれば時計回りに回転し、PNG に保存する

見た目を整えるための調整で、画像どうしの定量比較には使わない。
使った設定と µm/pixel は PNG のテキスト情報に残す。
"""

import argparse
import json
import sys
from pathlib import Path

import numpy as np
from PIL import Image
from PIL.PngImagePlugin import PngInfo
from pylibCZIrw import czi as pyczi

BLACK_EDGE_LEVEL = 16      # 全色がこれ未満の画素を「黒」とみなす
BLACK_EDGE_FRACTION = 0.05  # 縁の行・列のうち黒の割合がこれを超えたら切り落とす
SEAM_DEPTH = 40            # 上下（左右）の平均よりこれだけ暗ければ線の候補
SEAM_MIN_RUN = 16          # 候補がこれ以上連続したら線とみなす


def read_czi(path):
    """RGB（uint8）と µm/pixel を返す。"""
    with pyczi.open_czi(str(path)) as doc:
        metadata = doc.metadata
        raw = doc.read(plane={"C": 0, "Z": 0, "T": 0})
        pixel_type = doc.pixel_types.get(0, "")
    if raw.ndim != 3 or raw.shape[-1] != 3:
        sys.exit(f"エラー: {path.name} は明視野カラー画像ではありません（shape {raw.shape}）")
    rgb = raw[..., ::-1] if pixel_type.lower().startswith("bgr") else raw

    um_per_pixel = None
    items = metadata["ImageDocument"]["Metadata"].get("Scaling", {}).get("Items", {}).get("Distance", [])
    for item in items if isinstance(items, list) else [items]:
        if item.get("@Id") == "X" and item.get("Value"):
            um_per_pixel = float(item["Value"]) * 1e6
    return rgb, um_per_pixel


def trim_black_edges(rgb):
    """縁から黒い行・列を切り落とす。切った数を (上, 下, 左, 右) で返す。"""
    black = rgb.max(axis=2) < BLACK_EDGE_LEVEL
    top, bottom, left, right = 0, black.shape[0], 0, black.shape[1]
    while top < bottom and black[top, left:right].mean() > BLACK_EDGE_FRACTION:
        top += 1
    while bottom > top and black[bottom - 1, left:right].mean() > BLACK_EDGE_FRACTION:
        bottom -= 1
    while left < right and black[top:bottom, left].mean() > BLACK_EDGE_FRACTION:
        left += 1
    while right > left and black[top:bottom, right - 1].mean() > BLACK_EDGE_FRACTION:
        right -= 1
    trimmed = (top, rgb.shape[0] - bottom, left, rgb.shape[1] - right)
    return rgb[top:bottom, left:right], trimmed


def fix_seams(rgb):
    """1 px 幅の黒い線（タイル境界）を両隣の平均で埋める。直した線の数を返す。"""
    img = rgb.astype(np.float32)
    count = 0
    for axis_img in (img, img.transpose(1, 0, 2)):  # 横線、次に縦線（転置はビューなので img に反映される）
        lum = axis_img.mean(axis=2)
        depth = lum[1:-1] - (lum[:-2] + lum[2:]) / 2
        for y, row in enumerate(depth < -SEAM_DEPTH, start=1):
            if row.sum() < SEAM_MIN_RUN:
                continue
            # 連続区間だけを直す（核などの点状の暗さは対象外）
            edges = np.flatnonzero(np.diff(np.concatenate(([0], row.astype(np.int8), [0]))))
            for start, end in zip(edges[::2], edges[1::2]):
                if end - start >= SEAM_MIN_RUN:
                    axis_img[y, start:end] = (axis_img[y - 1, start:end] + axis_img[y + 1, start:end]) / 2
                    count += 1
    return img, count


def adjust(rgb, white_level, white_percentile, black, gamma):
    """白バランス・黒点・ガンマ。0〜1 の float を返す。"""
    white = np.percentile(rgb.reshape(-1, 3), white_percentile, axis=0)
    balanced = np.clip(rgb * (white_level / np.clip(white, 1, None)), 0, 255) / 255
    low = np.percentile(balanced.mean(axis=2), 0.5) * black
    high = min(1.0, white_level / 255 + 0.01)
    return np.clip((balanced - low) / (high - low), 0, 1) ** gamma, white


def convert(path, output_dir, args):
    rgb, um_per_pixel = read_czi(path)
    rgb, trimmed = trim_black_edges(rgb)
    img, seams = fix_seams(rgb) if not args.no_seam_fix else (rgb.astype(np.float32), 0)
    out, white = adjust(img, args.white_level, args.white_percentile, args.black, args.gamma)

    image = Image.fromarray((out * 255).round().astype(np.uint8))
    if args.rotate:
        image = image.rotate(-args.rotate, expand=True)  # PIL は反時計回りなので符号を反転

    settings = {
        "source": path.name,
        "umPerPixel": um_per_pixel,
        "trimmed": dict(zip(("top", "bottom", "left", "right"), trimmed)),
        "seamsFixed": seams,
        "whiteReference": [round(float(v), 1) for v in white],
        "whiteLevel": args.white_level,
        "black": args.black,
        "gamma": args.gamma,
        "rotate": args.rotate,
    }
    info = PngInfo()
    info.add_text("czi_to_png", json.dumps(settings, ensure_ascii=False))
    output_path = output_dir / f"{path.stem}.png"
    image.save(output_path, pnginfo=info)

    scale = f"{um_per_pixel:.3f} µm/px" if um_per_pixel else "µm/px 不明"
    cut = f"縁 {sum(trimmed)} px 切除" if sum(trimmed) else "縁の切除なし"
    print(f"{path.name} → {output_path.name}  {image.size[0]}×{image.size[1]}  {scale}  {cut}  線の補修 {seams}")


def main():
    parser = argparse.ArgumentParser(description="明視野 CZI を明るさ調整した PNG に書き出す")
    parser.add_argument("czi", nargs="+", type=Path, help="変換する CZI")
    parser.add_argument("-o", "--output-dir", type=Path, required=True, help="PNG の出力先フォルダ")
    parser.add_argument("--rotate", type=int, default=0, choices=(0, 90, 180, 270), help="時計回りの回転角（既定 0）")
    parser.add_argument("--white-level", type=float, default=245, help="背景を合わせる明るさ 0〜255（既定 245）")
    parser.add_argument("--white-percentile", type=float, default=99.5, help="背景とみなすパーセンタイル（既定 99.5）")
    parser.add_argument("--black", type=float, default=0.4, help="黒点をどれだけ詰めるか 0〜1（既定 0.4、0 で無効）")
    parser.add_argument("--gamma", type=float, default=1.3, help="中間調のガンマ（既定 1.3、1 で無効）")
    parser.add_argument("--no-seam-fix", action="store_true", help="1 px の黒い線を補修しない")
    args = parser.parse_args()

    for path in args.czi:
        if not path.is_file():
            sys.exit(f"エラー: 見つかりません: {path}")
    args.output_dir.mkdir(parents=True, exist_ok=True)
    for path in args.czi:
        convert(path, args.output_dir, args)


if __name__ == "__main__":
    main()
