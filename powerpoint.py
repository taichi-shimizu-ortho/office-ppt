import os
import glob
from pptx import Presentation
from pptx.util import Inches

# ==================== 設定エリア ====================
# 追記対象のPowerPointファイル
PPTX_PATH = "260602msc.pptx"

# 画像が入っているフォルダーのパス
IMAGE_DIR = r"C:\Users\a2189\OneDrive\デスクトップ\260602\調整後"
# ====================================================

# 画像の配置座標
POSITIONS = {
    "c1":    (Inches(0.5),  Inches(1.5), Inches(2.8), Inches(3.7)),
    "c2":    (Inches(3.6),  Inches(1.5), Inches(2.8), Inches(3.7)),
    "c3":    (Inches(6.7),  Inches(1.5), Inches(2.8), Inches(3.7)),
    "c1-3":  (Inches(9.8),  Inches(1.5), Inches(2.8), Inches(3.7))
}

def get_channel_type(filename_without_ext):
    if filename_without_ext.endswith("_c1"):
        return "c1"
    elif filename_without_ext.endswith("_c2"):
        return "c2"
    elif filename_without_ext.endswith("_c3"):
        return "c3"
    elif filename_without_ext.endswith("_c1-3"):
        return "c1-3"
    return None

def clone_shape(shape, target_slide):
    """1枚目のスライドにある図形を複製する（エラー対策版）"""
    if shape.has_text_frame:
        new_shape = target_slide.shapes.add_textbox(shape.left, shape.top, shape.width, shape.height)
        new_shape.text_frame.text = shape.text_frame.text
        
        for i, paragraph in enumerate(shape.text_frame.paragraphs):
            if i < len(new_shape.text_frame.paragraphs):
                new_para = new_shape.text_frame.paragraphs[i]
            else:
                new_para = new_shape.text_frame.add_paragraph()
            
            new_para.text = paragraph.text
            new_para.font.name = paragraph.font.name
            new_para.font.size = paragraph.font.size
            new_para.font.bold = paragraph.font.bold
            new_para.font.italic = paragraph.font.italic
            
            try:
                if paragraph.font.color and paragraph.font.color.rgb:
                    new_para.font.color.rgb = paragraph.font.color.rgb
            except AttributeError:
                pass
                
            new_para.alignment = paragraph.alignment
    else:
        if shape.shape_type != 13: 
            try:
                shape._element.getparent().append(shape._element)
            except Exception:
                pass

def main():
    if not os.path.exists(PPTX_PATH):
        print(f"エラー: 対象のPowerPointファイルが見つかりません:\n{PPTX_PATH}")
        return
    
    prs = Presentation(PPTX_PATH)
    
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    
    # --- 【重複防止】既存のスライドタイトルをすべて取得 ---
    existing_titles = set()
    for slide in prs.slides:
        for shape in slide.shapes:
            # テキストボックスがあり、かつそのテキストが空でない場合
            if shape.has_text_frame and shape.text_frame.text.strip():
                # タイトルとして追加したテキストボックス（ファイル名）を検出
                # 前後の空白を排除して記録
                existing_titles.add(shape.text_frame.text.strip())
    
    # 1枚目のスライドから、コピー元となるテキストボックス（図形）を抽出
    template_shapes = []
    if len(prs.slides) > 0:
        first_slide = prs.slides[0]
        for shape in first_slide.shapes:
            if shape.shape_type != 13: 
                template_shapes.append(shape)
    else:
        print("エラー: 1枚目のスライドが存在しません。")
        return

    # 白紙のスライドレイアウト
    try:
        slide_layout = prs.slide_layouts[6] 
    except IndexError:
        slide_layout = prs.slide_layouts[0]

    image_files = glob.glob(os.path.join(IMAGE_DIR, "*.jpg")) + glob.glob(os.path.join(IMAGE_DIR, "*.JPG"))

    if not image_files:
        print(f"警告: 指定されたフォルダにJPG画像が見つかりませんでした:\n{IMAGE_DIR}")
        return

    # 画像をグループ化
    groups = {}
    for filepath in image_files:
        filename = os.path.basename(filepath)
        name_without_ext, _ = os.path.splitext(filename)
        
        channel = get_channel_type(name_without_ext)
        
        if channel:
            suffix_len = len(channel) + 1
            base_name = name_without_ext[:-suffix_len]
            
            if base_name not in groups:
                groups[base_name] = {}
            groups[base_name][channel] = filepath

    print(f"フォルダ内の総画像セット数: {len(groups)} セット")

    # 新規追加するセットのカウンター
    added_count = 0

    for base_name, channels in sorted(groups.items()):
        # --- 【重複防止】すでにスライドが存在する場合はスキップ ---
        if base_name in existing_titles:
            # print(f"  [スキップ] すでにスライドが存在します: {base_name}")
            continue
        
        print(f"\n[新規追加] {base_name} のスライドを作成中...")
        slide = prs.slides.add_slide(slide_layout)
        added_count += 1
        
        # 1. 1枚目のスライドからコピーしたテキストボックス（DAPIなどの文字）を配置
        for shape in template_shapes:
            clone_shape(shape, slide)
        
        # 2. スライド上部にタイトル（ファイル名）を配置
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(12.33), Inches(0.8))
        tf = title_box.text_frame
        tf.text = base_name
        tf.paragraphs[0].font.size = Inches(0.3)
        tf.paragraphs[0].font.bold = True
        
        # 3. 画像を配置
        for ch in ["c1", "c2", "c3", "c1-3"]:
            if ch in channels:
                img_path = channels[ch]
                left, top, width, height = POSITIONS[ch]
                try:
                    slide.shapes.add_picture(img_path, left, top, width=width, height=height)
                    print(f"  [{ch}] 配置完了: {os.path.basename(img_path)}")
                except Exception as e:
                    print(f"  [エラー] {os.path.basename(img_path)} の配置に失敗: {e}")
            else:
                print(f"  [警告] {base_name} に対する チャンネル {ch} の画像が見つかりません。")

    if added_count > 0:
        try:
            prs.save(PPTX_PATH)
            print(f"\n完了！ 新しく {added_count} 枚のスライドを追加して上書き保存しました。")
        except PermissionError:
            print(f"\n[エラー] PowerPointファイル '{PPTX_PATH}' が開いたままになっています。")
            print("ファイルを閉じてから、もう一度スクリプトを実行してください。")
    else:
        print("\n追加が必要な新しい画像セットはありませんでした（すべて処理済みです）。")

if __name__ == "__main__":
    main()
