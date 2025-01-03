from dotenv import load_dotenv
import os
from pptx import Presentation
from pptx.util import Inches, Pt
import re
from pathlib import Path
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

def convert_marp_to_pptx(input_file: Path, output_file: Path) -> None:
    # 新しいプレゼンテーションを作成
    prs = Presentation()
    prs.slide_width = Inches(16)
    prs.slide_height = Inches(9)
    
    # Marpファイルを読み込む
    with open(input_file, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # スライドを分割
    slides = content.split('---')
    
    # フロントマターを除去
    if slides[0].strip().startswith('---'):
        slides = slides[1:]
    
    # CSSスタイルを解析
    style_definitions = {}
    css_pattern = re.compile(r'<style[^>]*>(.*?)</style>', re.DOTALL)
    css_matches = css_pattern.findall(content)
    
    for css in css_matches:
        # クラス定義を探す（例: .class-name { property: value; }）
        class_pattern = re.compile(r'\.([^{]+){([^}]+)}')
        for match in class_pattern.finditer(css):
            class_name = match.group(1).strip()
            properties = match.group(2).strip()
            style_definitions[class_name] = dict(
                prop.strip().split(':') 
                for prop in properties.split(';') 
                if ':' in prop
            )
    
    for slide_content in slides:
        # 空のスライドをスキップ
        if not slide_content.strip():
            continue
            
        # 新しいスライドを追加
        slide = prs.slides.add_slide(prs.slide_layouts[1])  # タイトルと本文のレイアウト
        
        # スライドの内容を解析
        lines = slide_content.strip().split('\n')
        
        # タイトルを探す（# で始まる最初の行）
        title = ""
        content_lines = []
        for line in lines:
            # h1（#）をスライドタイトルとして処理
            if not title and line.strip().startswith('# '):
                title = line.strip('#').strip()
            # h2-h6（## ～ ######）を本文の見出しとして処理
            elif line.strip().startswith('#'):
                heading_level = len(line.split()[0])  # #の数を数える
                content = line.strip('#').strip()
                p = text_frame.add_paragraph()
                p.text = content
                # 見出しレベルに応じてスタイルを設定
                if heading_level == 2:  # h2
                    p.font.size = Pt(32)
                    p.font.bold = True
                elif heading_level == 3:  # h3
                    p.font.size = Pt(28)
                    p.font.bold = True
                elif heading_level <= 6:  # h4-h6
                    p.font.size = Pt(24)
                    p.font.bold = True
            else:
                # 既存のdivタグ処理とその他のコンテンツ処理
                div_pattern = re.compile(r'<div\s+class=["\']([^"\']+)["\']>(.*?)</div>', re.DOTALL)
                div_match = div_pattern.match(line.strip())
                
                if div_match:
                    class_name = div_match.group(1)
                    content = div_match.group(2).strip()
                    
                    # スタイルの適用
                    p = text_frame.add_paragraph()
                    p.text = content
                    
                    if class_name in style_definitions:
                        style = style_definitions[class_name]
                        if 'color' in style:
                            # カラーコードをRGBに変換
                            color = style['color'].strip('#')
                            if len(color) == 6:
                                r = int(color[:2], 16)
                                g = int(color[2:4], 16)
                                b = int(color[4:], 16)
                                p.font.color.rgb = RGBColor(r, g, b)
                        
                        if 'font-size' in style:
                            # フォントサイズの設定（pxをポイントに変換）
                            size = style['font-size'].replace('px', '')
                            p.font.size = Pt(float(size))
                            
                        if 'text-align' in style:
                            # テキストの配置
                            align = style['text-align']
                            if align == 'center':
                                p.alignment = PP_ALIGN.CENTER
                            elif align == 'right':
                                p.alignment = PP_ALIGN.RIGHT
                            elif align == 'left':
                                p.alignment = PP_ALIGN.LEFT
                else:
                    content_lines.append(line)
        
        # タイトルを設定
        if title and slide.shapes.title:
            title_shape = slide.shapes.title
            title_shape.left = Inches(0.5)
            title_shape.top = Inches(0.5)
            title_shape.width = Inches(15)  # 16 - (0.5 * 2)
            title_shape.text = title
            
        # 本文を設定
        if content_lines and slide.placeholders[1]:
            body_shape = slide.placeholders[1]
            body_shape.left = Inches(0.5)
            body_shape.top = Inches(2)
            body_shape.width = Inches(15)  # 16 - (0.5 * 2)
            body_shape.height = Inches(6.5)  # 9 - 2 - 0.5
            
            text_frame = body_shape.text_frame
            for line in content_lines:
                p = text_frame.add_paragraph()
                # Markdownの箇条書きを変換
                if line.strip().startswith('- '):
                    p.text = line.strip('- ').strip()
                    p.level = 1
                else:
                    p.text = line.strip()
    
    # PowerPointファイルを保存
    prs.save(str(output_file))

if __name__ == "__main__":
    load_dotenv()
    work_folder: Path = Path(os.getenv("WORK_FOLDER"))
    input_file: Path = work_folder / Path("main.md")
    output_file: Path = work_folder / Path("presentation.pptx")
    convert_marp_to_pptx(input_file, output_file)

