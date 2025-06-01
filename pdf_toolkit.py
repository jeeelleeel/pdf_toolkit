# copyright (c) 2025 Tomohisa Saisho all rights reserved.
# This software is released under the MIT License. 

import os
import glob
import inspect
from enum import Enum
import logging
import pymupdf
import natsort
import os

logger =  logging.getLogger(__name__)
stream_handler = logging.StreamHandler()
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
stream_handler.setFormatter(formatter)
logger.addHandler(stream_handler)
logger.setLevel(logging.INFO)

def hello():
    """
        PDF Toolkitの初期化処理を行います。
    """   
    logger.info("Hello from PDF Toolkit!")

def create_rect(x,y,width,height):
    """  
        原点の(x,y)座標と幅・高さをを指定し、pymupdf.Rect型を生成します。
        ∵ pypdf.Rect型コンストラクタは、(x1,y1,x2,y2)の座標で矩形を特定するため直感的にわかりにくい。
    """
    return pymupdf.Rect(x,y,x+width,y+height)

def show_pdf_info(pdf_path):
    """
        PDFファイルの情報を表示します。
        引数:
            pdf_path: PDFファイルのパス
    """
    try:
        doc = pymupdf.open(pdf_path)
        logger.info(f"PDFファイル：{doc}の情報を表示します。")
        logger.info(f"    ページ数：{doc.page_count}")
        logger.info(f"    ページサイズ：{doc[0].rect}")
        logger.info(f"    ページの回転：{doc[0].rotation}")
        logger.info(f"    ページのスケール：{doc[0].get_page_scale()}")
        logger.info(f"    ページのスケール（X,Y）：{doc[0].get_page_scale_x()}, {doc[0].get_page_scale_y()}")
        logger.info(f"    ページのスケール（幅,高さ）：{doc[0].get_page_scale_width()}, {doc[0].get_page_scale_height()}")
        logger.info(f"    ページのスケール（幅,高さ）(mm)：{doc[0].get_page_scale_width()*25.4/72}, {doc[0].get_page_scale_height()*25.4/72}")
        logger.info(f"    ページのスケール（幅,高さ）(cm)：{doc[0].get_page_scale_width()*25.4/72/10}, {doc[0].get_page_scale_height()*25.4/72/10}")
        logger.info(f"    ページのスケール（幅,高さ）(inch)：{doc[0].get_page_scale_width()*25.4/72/25.4}, {doc[0].get_page_scale_height()*25.4/72/25.4}")
        logger.info(f"    ページのスケール（幅,高さ）(pixel)：{doc[0].get_page_scale_width()*25.4/72*72}, {doc[0].get_page_scale_height()*25.4/72*72}")
    except Exception as e:
        logger.error(f"エラーが発生しました： {e}")
    finally:
        if 'doc' in locals() and not doc.is_closed:
            doc.close()

def masking_to_pdf(input_pdf_path, output_pdf_path, mask_rect_array, overwrite=True):
    """ 
        input_pdf_pathで指定されたPDFファイルに対し、mask_rect_arrayで指定された矩形領域のマスキング処理を行い、output_pdf_pathで指定されたファイルに保存します。
        引数： 
            input_pdf_path: 入力PDFファイルのパス
            output_pdf_path: 出力PDFファイルのパス
            mask_rect_array: マスキングする矩形領域のリスト
            overwrite: 出力ファイルが既に存在する場合に上書きするかどうか
    """
    try:
        if(not os.path.exists(input_pdf_path)):
            logger.fatal(f"ファイル：{input_pdf_path}は存在しません。{inspect.currentframe().f_code.co_name}を終了します。")
            return
        if(input_pdf_path == output_pdf_path):
            logger.fatal(f"入力ファイルと同一のファイルが出力先に指定されています。{inspect.currentframe().f_code.co_name}を終了します。")
            return
        if(not overwrite and os.path.exists(output_pdf_path)):
            logger.fatal(f"出力ファイル：{output_pdf_path}は既に存在します。上書きしません。{inspect.currentframe().f_code.co_name}を終了します。")
            return

        doc = pymupdf.open(input_pdf_path)
        logger.debug(f"PDFファイル：{doc}にマスキング処理を行います...")

        for page in doc.pages():
            if(page.rotation != 0):
                # ページの回転を解除
                page.remove_rotation()
                
            for mask_rect in mask_rect_array:
                logger.debug(f"マスキング:{mask_rect} to {page}")
                if isinstance(mask_rect, pymupdf.Rect) and mask_rect.is_valid:
                    page.add_redact_annot(mask_rect, fill=(0, 0, 0))
                else:
                    logger.error(f"Invalid mask_rect: {mask_rect}")
            page.apply_redactions()

        doc.save(output_pdf_path, garbage=3, clean=True, deflate=True, deflate_images=True, deflate_fonts=True)
        logger.debug(f"マスキング処理が完了しました。")
    except Exception as e:
        logger.error(f"エラーが発生しました： {e}")
    finally:
        if 'doc' in locals() and not doc.is_closed:
            doc.close()

def masking_to_pdf_in_folder(input_folder_path, output_folder_path, rect_array, overwrite=True):
    """ 
       input_folder_pathで指定されたフォルダ内に存在する全てのPDFファイルに対し、rect_arrayで指定された矩形領域のマスキング処理を行い、output_folder_pathに出力します。
         引数：
            input_folder_path: 入力フォルダのパス
            output_folder_path: 出力フォルダのパス
            rect_array: マスキングする矩形領域のリスト
            overwrite: 出力ファイルが既に存在する場合に上書きするかどうか  
    """
    if(input_folder_path == output_folder_path):
        logger.fatal(f"入力フォルダと同一のフォルダが出力先に指定されています。{inspect.currentframe().f_code.co_name}を終了します。")
        return
    if(not os.path.exists(input_folder_path)):
        logger.fatal(f"フォルダ：{input_folder_path}は存在しません。{inspect.currentframe().f_code.co_name}を終了します。")
        return

    input_folder_abspath=os.path.abspath(input_folder_path)
    output_folder_abspath=os.path.abspath(output_folder_path)
           
    if(not os.path.exists(os.path.abspath(output_folder_abspath))):
        os.makedirs(os.path.abspath(output_folder_abspath))
        logger.info(f"フォルダ：{output_folder_abspath}を作成しました。")

    logger.info(f"フォルダ：{input_folder_abspath}中のPDFをマスキングし、フォルダ{output_folder_abspath}へ出力します。")

    for pdf_in_file in glob.glob(input_folder_abspath+"/*.pdf"):
            pdf_out_file = output_folder_abspath + "/" + os.path.basename(pdf_in_file)
            logger.info(f"PDFファイル{pdf_in_file}をマスキングし、{pdf_out_file}へ出力します。")
            masking_to_pdf(pdf_in_file,pdf_out_file,rect_array,overwrite=overwrite)
    
def grid_to_pdf(
        input_pdf_path, 
        output_pdf_path, 
        grid_interval=10.0, 
        grid_width=0.1,
        grid_thin_color=(0.502, 1, 0.553),
        grid_medium_color=(0.008, 0.859, 0.102),
        grid_thick_color=(0, 0.702, 0.075),
        overwrite=True
    ):
    """
        input_pdf_pathで指定されたPDFファイルの各ページに、grid_interalで指定された間隔で矩形グリッドを描画し、output_pdf_pathで指定されたPDFをファイルに保存します。
        引数：
            input_pdf_path: 入力PDFファイルのパス
            output_pdf_path: 出力PDFファイルのパス
            grid_interval: グリッド線の間隔
            grid_width: グリッド線の幅
            grid_thin_color: グリッド線の色 (細)
            grid_medium_color: グリッド線の色 (中)
            grid_thick_color: グリッド線の色 (太)
            overwrite: 出力ファイルが既に存在する場合に上書きするかどうか
    """
    try:
        if(not os.path.exists(input_pdf_path)):
            logger.fatal(f"ファイル：{input_pdf_path}は存在しません。{inspect.currentframe().f_code.co_name}を終了します。")
            return
        if(input_pdf_path == output_pdf_path):
            logger.fatal(f"入力ファイルと同一のファイルが出力先に指定されています。{inspect.currentframe().f_code.co_name}を終了します。")
            return
        if(not overwrite and os.path.exists(output_pdf_path)):
            logger.fatal(f"出力ファイル：{output_pdf_path}は既に存在します。上書きしません。{inspect.currentframe().f_code.co_name}を終了します。")
            return

        # PDFドキュメントを開く
        doc = pymupdf.open(input_pdf_path)

        # 各ページを処理
        for page_num in range(len(doc)):
            page = doc[page_num]
            if(page.rotation != 0):
                # ページの回転を解除。これにより、ページの回転が描画に影響しないようにする。
                page.remove_rotation()

            logger.debug(f"draw grid to page {page} / {page.rect}")

            # ページごとに新しいShapeオブジェクトを作成
            shape = page.new_shape()

            # ページの幅と高さを取得
            page_width = page.rect.width
            page_height = page.rect.height

            # 垂直方向にグリッド線に対応する矩形を描画
            # y座標を開始点として、grid_intervalずつ増加させながらループ
            for y_start in range(0, int(page_height), int(grid_interval)):
                # 水平方向にグリッド線に対応する矩形を描画
                # x座標を開始点として、grid_intervalずつ増加させながらループ
                for x_start in range(0, int(page_width), int(grid_interval)):
                    # 矩形の座標を定義
                    x0 = float(x_start)
                    y0 = float(y_start)
                    # x1, y1 はページの境界を超えないように調整
                    x1 = min(x_start + grid_interval, page_width)
                    y1 = min(y_start + grid_interval, page_height)

                    # 矩形が有効なサイズを持つ場合のみ描画
                    # (幅または高さが0以下の矩形は描画しない)
                    if x1 > x0 and y1 > y0:
                        rect_to_draw = pymupdf.Rect(x0, y0, x1, y1)
                        logger.debug(f"draw rect:{rect_to_draw}")
                        shape.draw_rect(rect_to_draw)
            
            # このページの全ての矩形に対するスタイルを定義し、描画を準備
            shape.finish(width=grid_width, color=grid_thin_color, fill=None)
            
            # ２回めのループ。矩形の描画間隔を５倍にして太線を描く。
            grid_interval_5th = grid_interval * 5
            # 垂直方向にグリッド線に対応する矩形を描画
            # y座標を開始点として、grid_intervalずつ増加させながらループ
            for y_start in range(0, int(page_height), int(grid_interval_5th)):
                # 水平方向にグリッド線に対応する矩形を描画
                # x座標を開始点として、grid_intervalずつ増加させながらループ
                for x_start in range(0, int(page_width), int(grid_interval_5th)):
                    # 矩形の座標を定義
                    x0 = float(x_start)
                    y0 = float(y_start)
                    # x1, y1 はページの境界を超えないように調整
                    x1 = min(x_start + grid_interval_5th, page_width)
                    y1 = min(y_start + grid_interval_5th, page_height)

                    # 矩形が有効なサイズを持つ場合のみ描画
                    # (幅または高さが0以下の矩形は描画しない)
                    if x1 > x0 and y1 > y0:
                        rect_to_draw = pymupdf.Rect(x0, y0, x1, y1)
                        shape.draw_rect(rect_to_draw)
            shape.finish(width=grid_width*5, color=grid_medium_color, fill=None)
            
            # 3回めのループ。矩形の描画間隔を10倍にして極太線を描く。
            grid_interval_10th = grid_interval * 10
            # 垂直方向にグリッド線に対応する矩形を描画
            # y座標を開始点として、grid_intervalずつ増加させながらループ
            for y_start in range(0, int(page_height), int(grid_interval_10th)):
                # 水平方向にグリッド線に対応する矩形を描画
                # x座標を開始点として、grid_intervalずつ増加させながらループ
                for x_start in range(0, int(page_width), int(grid_interval_10th)):
                    # 矩形の座標を定義
                    x0 = float(x_start)
                    y0 = float(y_start)
                    # x1, y1 はページの境界を超えないように調整
                    x1 = min(x_start + grid_interval_10th, page_width)
                    y1 = min(y_start + grid_interval_10th, page_height)

                    # 矩形が有効なサイズを持つ場合のみ描画
                    # (幅または高さが0以下の矩形は描画しない)
                    if x1 > x0 and y1 > y0:
                        rect_to_draw = pymupdf.Rect(x0, y0, x1, y1)
                        shape.draw_rect(rect_to_draw)
            
            shape.finish(width=grid_width*10, color=grid_thick_color, fill=None)
            shape.commit()
        # 変更を保存
        doc.save(output_pdf_path,garbage=3,clean=True,deflate=True,deflate_images=True,deflate_fonts=True)
        logger.debug(f"グリッドが描画されたPDFを {output_pdf_path} に保存しました。")
    except Exception as e:
        logger.debug(f"エラーが発生しました: {e}")
    finally:
        if 'doc' in locals() and doc.is_closed:
            doc.close()

class Delimiters(Enum):
    No = 0
    Space = 1
    Colon = 2
    Semicolon = 3
    Comma = 4
    Hyphen = 5
    Underscore = 6
    Slash = 7
    Backslash = 8
    At = 9
    Pipe = 10

def split_by_delimiter(text,delimiter=Delimiters.No):
    """
        指定された区切り文字でテキストを分割し、最初の部分を返します。
        引数:
            text: 分割するテキスト
            delimiter: 区切り文字を指定する列挙型
        戻り値:
            最初の部分のテキスト
    """
    if(delimiter==Delimiters.No):
        return text
    elif(delimiter==Delimiters.Space): 
        return text.split()[0]
    elif(delimiter==Delimiters.Colon): 
        return text.split(":")[0]
    elif(delimiter==Delimiters.Semicolon): 
        return text.split(";")[0]
    elif(delimiter==Delimiters.Comma):
        return text.split(",")[0]
    elif(delimiter==Delimiters.Hyphen):
        return text.split("-")[0]
    elif(delimiter==Delimiters.Underscore):
        return text.split("_")[0]
    elif(delimiter==Delimiters.Slash):
        return text.split("/")[0]
    elif(delimiter==Delimiters.Backslash):
        return text.split("\\")[0]
    elif(delimiter==Delimiters.At):
        return text.split("@")[0]
    elif(delimiter==Delimiters.Pipe):
        return text.split("|")[0]
    else:
        return text
    
def header_to_pdf(
        input_pdf_path,
        output_pdf_path,
        overwrite=True,
         
        font_name="japan", 
        font_size=20,
        text_color=(1, 0, 0), 
        
        header_height=70,
        padding=10,
        resize_original=False,
        
        delimiter=Delimiters.Space,
        
        stamp_only_firstpage=False,
        
        draw_header_line=False,
        header_line_color=(0, 0, 0),
        header_line_width=0.5,
        
        draw_content_border=False,
        content_border_width=0.5,
        content_border_color=(0,0,0) 
    ):
    """
    PDFの各ページにファイル名をヘッダーとして追加します。 
    引数:
        input_pdf_path: 入力PDFファイルのパス
        output_pdf_path: 出力PDFファイルのパス
        overwrite: 出力ファイルが既に存在する場合に上書きするかどうか

        font_name: ヘッダーフォント名
        font_size: ヘッダーフォントサイズ
        text_color: ヘッダーテキストカラー

        header_height: ヘッダーの高さ
        padding: ヘッダーのパディング
        resize_original: 元のページを縮小するかどうか

        delimiter: ファイル名を分割する区切り文字
        stamp_only_firstpage: 最初のページのみスタンプするかどうか
        draw_header_line: ヘッダー下に線を描画するかどうか
        header_line_color: ヘッダー線の色
        header_line_width: ヘッダー線の幅
        
        draw_content_border: コンテンツの境界線を描画するかどうか
        content_border_width: コンテンツ境界線の幅
        content_border_color: コンテンツ境界線の色
    """
    try:
        if(not os.path.exists(input_pdf_path)):
            logger.fatal(f"ファイル：{input_pdf_path}は存在しません。{inspect.currentframe().f_code.co_name}を終了します。")
            return
        if(input_pdf_path == output_pdf_path):
            logger.fatal(f"入力ファイルと同一のファイルが出力先に指定されています。{inspect.currentframe().f_code.co_name}を終了します。")
            return
        if(not overwrite and os.path.exists(output_pdf_path)):
            logger.fatal(f"出力ファイル：{output_pdf_path}は既に存在します。上書きしません。{inspect.currentframe().f_code.co_name}を終了します。")
    
        base_name = os.path.basename(input_pdf_path)
        header_text_content = split_by_delimiter(os.path.splitext(base_name)[0],delimiter=delimiter)
    
        source_doc = pymupdf.open(input_pdf_path)
        output_doc = pymupdf.open()

        for page_idx in range(source_doc.page_count):
            original_page = source_doc.load_page(page_idx)

            if(original_page.rotation != 0):
                original_page.remove_rotation()

            # 出力ドキュメントに新しいページを作成 (元のページと同じサイズ)
            new_page = output_doc.new_page(width=original_page.rect.width, height=original_page.rect.height)

            # ヘッダー領域を定義
            header_actual_rect = pymupdf.Rect(0, 0, new_page.rect.width, header_height)
        
            # 元のコンテンツが表示される領域を定義 (ヘッダー領域の下)
            if(resize_original): #オリジナルを縮小
                content_rect = pymupdf.Rect(0, header_height, new_page.rect.width, new_page.rect.height)
            else: #オリジナルを縮小しない
                content_rect = pymupdf.Rect(0, 0, new_page.rect.width, new_page.rect.height)

            # 元のページコンテンツを新しいページのcontent_rectに表示 (スケーリング)
            new_page.show_pdf_page(content_rect, source_doc, page_idx, keep_proportion=True, overlay=True)
            
            if(draw_content_border and resize_original):
                # 実際に描画されたコンテンツ領域を計算
                resize_ratio = content_rect.height / original_page.rect.height
                actual_shown_rect = create_rect(
                    (content_rect.width - original_page.rect.width * resize_ratio)/2, header_height, 
                    original_page.rect.width * resize_ratio, content_rect.height - content_border_width
                )
                shape = new_page.new_shape()
                shape.draw_rect(actual_shown_rect)
                shape.finish(width=content_border_width, color=content_border_color, fill=None)
                shape.commit(overlay=True)

            if(draw_header_line):
                line_start = pymupdf.Point(header_actual_rect.x0 + padding, header_actual_rect.y1 - padding)
                line_end = pymupdf.Point(header_actual_rect.x1 - padding, header_actual_rect.y1 - padding )
                new_page.draw_line(line_start, line_end, color=header_line_color, width=header_line_width)

            if(stamp_only_firstpage):
                continue
            
            # ヘッダーテキストを挿入するための矩形 (パディングを考慮)
            text_insertion_rect = pymupdf.Rect(
                header_actual_rect.x0 + padding, 
                header_actual_rect.y0 + padding * 2, # ヘッダーテキストの上パディングを大きめにとる。∵印刷範囲外になる
                header_actual_rect.x1 - padding * 2, # ヘッダーテキストの右パディングを大きめにとる。
                header_actual_rect.y1 # - padding
            )
        
            new_page.insert_textbox(
                text_insertion_rect,
                header_text_content,
                fontname=font_name,
                fontsize=font_size,
                color=text_color,
                align=pymupdf.TEXT_ALIGN_RIGHT
            )
        
        output_doc.save(output_pdf_path,garbage=3,clean=True,deflate=True,deflate_images=True,deflate_fonts=True)
        output_doc.close()
        
    except Exception as e:
        logger.error(f"エラーが発生しました： {e}")
    finally:
        if 'source_doc' in locals() and not source_doc.is_closed:
            source_doc.close()
        if 'output_doc' in locals() and not output_doc.is_closed:
            output_doc.close()

def header_to_pdf_in_folder(
        input_folder_path,
        output_folder_path,
        
        font_name="japan", 
        font_size=20,
        text_color=(1, 0, 0), 
        
        header_height=70,
        padding=10,
        
        resize_original=False,
        
        delimiter=Delimiters.Space,
        
        stamp_only_firstpage=False,
        
        draw_header_line=False,
        header_line_color=(0, 0, 0),
        header_line_width=0.5,
        
        draw_content_border=False,
        content_border_width=0.5,
        content_border_color=(0,0,0) 
    ):  
    """
    input_folder_pathで指定されたフォルダ内に存在する全てのPDFファイルに対し、ヘッダーを付与し、output_folder_pathに出力します。
    引数:
        input_folder_path: 入力フォルダのパス
        output_folder_path: 出力フォルダのパス
        font_name: ヘッダーフォント名
        font_size: ヘッダーフォントサイズ
        text_color: ヘッダーテキストカラー
        header_height: ヘッダーの高さ
        padding: ヘッダーのパディング
        resize_original: 元のページを縮小するかどうか
        delimiter: ファイル名を分割する区切り文字
        stamp_only_firstpage: 最初のページのみスタンプするかどうか
        draw_header_line: ヘッダー下に線を描画するかどうか
        header_line_color: ヘッダー線の色
        header_line_width: ヘッダー線の幅
        draw_content_border: 元PDFファイルのページコンテンツを囲む矩形を描画するかどうか
        content_border_width: 矩形を囲む線の幅
        content_border_color: 矩形を囲む線の色
    """
    if(input_folder_path == output_folder_path):
        logger.fatal(f"入力フォルダと同一のフォルダが出力先に指定されています。{inspect.currentframe().f_code.co_name}を終了します。")
        return
    if(not os.path.exists(input_folder_path)):
        logger.fatal(f"フォルダ：{input_folder_path}は存在しません。{inspect.currentframe().f_code.co_name}を終了します。")
        return
    
    input_folder_abspath = os.path.abspath(input_folder_path)
    output_folder_abspath = os.path.abspath(output_folder_path)
    logger.debug(f"フォルダ：{input_folder_abspath}中のPDFファイルにヘッダーを付与し、フォルダ{os.path.abspath(input_folder_abspath)}へ出力します。")
       
    if not(os.path.exists(os.path.abspath(output_folder_abspath))):
        os.makedirs(os.path.abspath(output_folder_abspath))
        logger.debug(f"フォルダ：{output_folder_abspath}を作成しました。")

    for pdf_in_file in glob.glob(input_folder_abspath+"/*.pdf"):
        pdf_out_file = output_folder_abspath + "/" + os.path.basename(pdf_in_file)
        logger.debug(f"PDFファイル{pdf_in_file}にヘッダーを付与し、{pdf_out_file}へ出力します。")
        header_to_pdf(
            input_pdf_path=pdf_in_file,
            output_pdf_path=pdf_out_file,
            font_name=font_name,
            font_size=font_size,
            text_color=text_color,
            header_height=header_height,
            padding=padding,
            resize_original=resize_original,
            delimiter=delimiter,
            stamp_only_firstpage=stamp_only_firstpage,
            draw_header_line=draw_header_line,
            header_line_color=header_line_color,
            header_line_width=header_line_width,
            draw_content_border=draw_content_border,
            content_border_color=content_border_color,
            content_border_width=content_border_width
        )

def pagenum_to_pdf(
        input_pdf_path, 
        output_pdf_path,
        overwrite=True,
        
        font_name="helv", 
        font_size=8, 
        text_color=(0, 0, 0), 
        
        footer_height=30,
        resize_original=False,
        padding=5,
        
        show_total_pages=False,
        
        draw_footer_line=False,
        footer_line_color=(0, 0, 0),
        footer_line_width=0.5
    ):
    """
    PDFの各ページにページ番号をフッターとして追加し、コンテンツをスケーリングします。
    引数:
        input_pdf_path: 入力PDFファイルのパス
        output_pdf_path: 出力PDFファイルのパス
        overwrite: 出力ファイルが既に存在する場合に上書きするかどうか
        font_name: フッターフォント名
        font_size: フッターフォントサイズ
        text_color: フッターテキストカラー
        footer_height: フッターの高さ
        resize_original: 元のページを縮小するかどうか
        padding: フッターのパディング
        show_total_pages: 総ページ数を表示するかどうか
        draw_footer_line: フッター上に線を描画するかどうか
        footer_line_color: フッター線の色
        footer_line_width: フッター線の幅
    """
    try:
        if(not os.path.exists(input_pdf_path)):
            logger.fatal(f"ファイル：{input_pdf_path}は存在しません。{inspect.currentframe().f_code.co_name}を終了します。")
            return
        if(input_pdf_path == output_pdf_path):
            logger.fatal(f"入力ファイルと同一のファイルが出力先に指定されています。{inspect.currentframe().f_code.co_name}を終了します。")
            return
        if(not overwrite and os.path.exists(output_pdf_path)):
            logger.fatal(f"出力ファイル：{output_pdf_path}は既に存在します。上書きしません。{inspect.currentframe().f_code.co_name}を終了します。")
    
        source_doc = pymupdf.open(input_pdf_path)
        output_doc = pymupdf.open()

        total_pages = source_doc.page_count

        for page_idx in range(total_pages):
            original_page = source_doc.load_page(page_idx)

            if(original_page.rotation != 0):
                original_page.remove_rotation()

            new_page = output_doc.new_page(width=original_page.rect.width, height=original_page.rect.height)

            # フッター領域を定義
            footer_actual_rect = pymupdf.Rect(0, new_page.rect.height - footer_height, new_page.rect.width, new_page.rect.height)
        
            # 元のコンテンツが表示される領域を定義 (フッター領域の上)
            content_rect = pymupdf.Rect(0, 0, new_page.rect.width, new_page.rect.height - footer_height)

            # オリジナルを縮小しない場合
            if not resize_original: 
                content_rect = pymupdf.Rect(0, 0, new_page.rect.width, new_page.rect.height)

            new_page.show_pdf_page(content_rect, source_doc, page_idx, keep_proportion=True, overlay=True)

            if(show_total_pages):
                page_number_text = f"{page_idx + 1} / {total_pages}"
            else:
                page_number_text = f"{page_idx + 1}"
        
            text_insertion_rect = pymupdf.Rect(
                footer_actual_rect.x0 + padding,
                footer_actual_rect.y0 + padding,
                footer_actual_rect.x1 - padding,
                footer_actual_rect.y1 - padding
            )
        
            new_page.insert_textbox(
                text_insertion_rect,
                page_number_text,
                fontname=font_name,
                fontsize=font_size,
                color=text_color,
                align=pymupdf.TEXT_ALIGN_CENTER
            )

            if(draw_footer_line):  
                line_start = pymupdf.Point(footer_actual_rect.x0 + padding, footer_actual_rect.y0 + padding)
                line_end = pymupdf.Point(footer_actual_rect.x1 - padding, footer_actual_rect.y0 + padding )
                new_page.draw_line(line_start, line_end, color=footer_line_color, width=footer_line_width)

        output_doc.save(output_pdf_path,garbage=3,clean=True,deflate=True,deflate_images=True,deflate_fonts=True)
        output_doc.close()
    except Exception as e:
        logger.error(f"エラーが発生しました： {e}")
    finally:
        if 'source_doc' in locals() and not source_doc.is_closed: 
            source_doc.close()
        if 'output_doc' in locals() and not output_doc.is_closed:
            output_doc.close()

def header_and_pagenum_to_pdf(
    # ヘッダー・フッターに共通する引数
    input_pdf_path, 
    output_pdf_path,
    overwrite=True,
    resize_original=True,

    # ヘッダー描画用引数
    header_font_name="japan", 
    header_font_size=20, 
    header_text_color=(1, 0, 0), 
    header_height=70,
    header_paddiing = 10,
    delimiter=Delimiters.Space,
    stamp_only_firstpage=False,
    draw_header_line=False,
    header_line_color=(0, 0, 0),
    header_line_width=0.5,

    # フッター描画用引数
    footer_font_name="helv", 
    footer_font_size=8, 
    footer_text_color=(0, 0, 0), 
    footer_height=30,
    footer_padding=3,
    write_pagenum = True,
    show_total_pages=False,
    draw_footer_line=False,
    footer_line_color=(0, 0, 0),
    footer_line_width=0.5,
    
    # 元のページコンテントを囲む矩形用引数
    draw_content_border=True,
    content_border_width=0.5,
    content_border_color=(0,0,0) 
):
    """
    PDFの各ページにファイル名をヘッダーとして、ページ番号をフッターとして追加し、コンテンツをスケーリングします。
    引数:  
        input_pdf_path: 入力PDFファイルのパス
        output_pdf_path: 出力PDFファイルのパス
        overwrite: 出力ファイルが既に存在する場合に上書きするかどうか
        resize_original: 元のページを縮小するかどうか
        
        header_font_name: ヘッダーフォント名
        header_font_size: ヘッダーフォントサイズ
        header_text_color: ヘッダーテキストカラー
        header_height: ヘッダーの高さ
        header_padding: ヘッダーのパディング
        delimiter: ファイル名を分割する区切り文字
        stamp_only_firstpage: 最初のページのみスタンプするかどうか
        draw_header_line: ヘッダー下に線を描画するかどうか
        header_line_color: ヘッダー線の色
        header_line_width: ヘッダー線の幅
        
        footer_font_name: フッターフォント名
        footer_font_size: フッターフォントサイズ
        footer_text_color: フッターテキストカラー
        footer_height: フッターの高さ
        footer_padding: フッターのパディング
        write_pagenum: ページ番号を描画するかどうか。
            ※これをFalseに設定し、resize_originalをTrueに設定し、フッターの高さを小さく設定することで、ベゼル的な表示ができる。
        show_total_pages: 総ページ数を表示するかどうか
        draw_footer_line: フッター上に線を描画するかどうか
        footer_line_color: フッター線の色
        footer_line_width: フッター線の幅
        
        draw_content_border: 元PDFファイルのページコンテンツを囲む矩形を描画するかどうか
        content_border_width: 矩形を囲む線の幅
        content_border_color: 矩形を囲む線の色
    """
    try:
        if(not os.path.exists(input_pdf_path)):
            logger.fatal(f"ファイル：{input_pdf_path}は存在しません。{inspect.currentframe().f_code.co_name}を終了します。")
            return
        if(input_pdf_path == output_pdf_path):
            logger.fatal(f"入力ファイルと同一のファイルが出力先に指定されています。{inspect.currentframe().f_code.co_name}を終了します。")
            return
        if(not overwrite and os.path.exists(output_pdf_path)):
            logger.fatal(f"出力ファイル：{output_pdf_path}は既に存在します。上書きしません。{inspect.currentframe().f_code.co_name}を終了します。")
   
        base_name = os.path.basename(input_pdf_path)
        header_text_content = split_by_delimiter(os.path.splitext(base_name)[0],delimiter=delimiter)
    
        source_doc = pymupdf.open(input_pdf_path)
        output_doc = pymupdf.open()

        total_pages = source_doc.page_count

        for page_idx in range(total_pages):
            original_page = source_doc.load_page(page_idx)

            if(original_page.rotation != 0):
                original_page.remove_rotation()

            new_page = output_doc.new_page(width=original_page.rect.width, height=original_page.rect.height)

            # ヘッダー領域とフッター領域を定義
            actual_header_rect = pymupdf.Rect(0, 0, new_page.rect.width, header_height)
            actual_footer_rect = pymupdf.Rect(0, new_page.rect.height - footer_height, new_page.rect.width, new_page.rect.height)
        
            # 元のコンテンツが表示される領域を定義
            if(resize_original):
                content_rect = pymupdf.Rect(0, header_height, new_page.rect.width, new_page.rect.height - footer_height)
            else:
                content_rect = pymupdf.Rect(0, 0, original_page.rect.width, original_page.rect.height)

            # コンテンツ領域の高さが負またはゼロにならないようにチェック
            if(content_rect.height <= 0):
                logger.warning(f"警告: ページ {page_idx + 1} でヘッダーとフッターの高さが大きすぎるため、コンテンツ領域がありません。スキップします。")
                # 元のページをそのままコピーするか、エラー処理を行う
                output_doc.insert_pdf(source_doc, from_page=page_idx, to_page=page_idx)
                continue

            new_page.show_pdf_page(content_rect, source_doc, page_idx, keep_proportion=True, overlay=True)

            if(write_pagenum):
                # フッターテキストを挿入
                if(show_total_pages):
                    page_number_text = f" {page_idx + 1} / {total_pages}"
                else:
                    page_number_text = f" {page_idx + 1}"
       
                footer_insertion_rect = pymupdf.Rect(
                    actual_footer_rect.x0 + footer_padding,
                    actual_footer_rect.y0 + footer_padding,
                    actual_footer_rect.x1 - footer_padding,
                    actual_footer_rect.y1 - footer_padding
                )

                new_page.insert_textbox(
                    footer_insertion_rect,
                    page_number_text,
                    fontname=footer_font_name,
                    fontsize=footer_font_size,
                    color=footer_text_color,
                    align=pymupdf.TEXT_ALIGN_CENTER
                )

            if(draw_footer_line):
                line_start = pymupdf.Point(actual_footer_rect.x0 + footer_padding, actual_footer_rect.y0 + footer_padding)
                line_end = pymupdf.Point(actual_footer_rect.x1 - footer_padding, actual_footer_rect.y0 + footer_padding)
                new_page.draw_line(line_start, line_end, color=footer_line_color, width=footer_line_width)

            if(draw_content_border and resize_original):
                # 実際に描画されたコンテンツ領域を計算
                resize_ratio = content_rect.height / original_page.rect.height
                actual_shown_rect = create_rect((content_rect.width - original_page.rect.width * resize_ratio)/2, header_height, original_page.rect.width * resize_ratio, content_rect.height)
                shape = new_page.new_shape()
                shape.draw_rect(actual_shown_rect)
                shape.finish(width=content_border_width, color=content_border_color, fill=None)
                shape.commit(overlay=True)
                
            if(draw_header_line):
                line_start = pymupdf.Point(actual_header_rect.x0 + header_paddiing, actual_header_rect.y1 - header_paddiing)
                line_end = pymupdf.Point(actual_header_rect.x1 - header_paddiing, actual_header_rect.y1 - header_paddiing )
                new_page.draw_line(line_start, line_end, color=header_line_color, width=header_line_width)
            
            if(stamp_only_firstpage and page_idx > 0 ):
                continue

            # ヘッダーテキストを挿入
            header_insertion_rect = pymupdf.Rect(
                actual_header_rect.x0 + header_paddiing,
                actual_header_rect.y0 + header_paddiing * 3, # ヘッダーテキストの上パディングを大きめにとる。 ∵印刷範囲外になる。
                actual_header_rect.x1 - header_paddiing * 3, # ヘッダーテキストの右パディングを大きめにとる。
                actual_header_rect.y1
            )

            new_page.insert_textbox(
                header_insertion_rect,
                header_text_content,
                fontname=header_font_name,
                fontsize=header_font_size,
                color=header_text_color,
                align=pymupdf.TEXT_ALIGN_RIGHT
            )

        output_doc.save(output_pdf_path,garbage=3,clean=True,deflate=True,deflate_images=True,deflate_fonts=True)
        output_doc.close()
        source_doc.close()
    except Exception as e:
        logger.error(f"エラーが発生しました： {e}")
    finally:
        if('source_doc' in locals() and not source_doc.is_closed):
            source_doc.close()
        if('output_doc' in locals() and not output_doc.is_closed):
            output_doc.close()

def header_and_pagenum_to_pdf_in_folder(
    input_folder_path,
    output_folder_path,
    overwrite=True,
    resize_original=True,

    # ヘッダー描画用引数
    header_font_name="japan", 
    header_font_size=20, 
    header_text_color=(1, 0, 0), 
    header_height=70,
    header_padding=10,
    delimiter=Delimiters.Space,
    stamp_only_firstpage=False,
    draw_header_line=False,
    header_line_color=(0, 0, 0),
    header_line_width=0.5,

    # フッター描画用引数
    footer_font_name="helv", 
    footer_font_size=8, 
    footer_text_color=(0, 0, 0), 
    footer_height=30,
    footer_padding=3,
    write_pagenum=True,
    show_total_pages=False,
    draw_footer_line=False,
    footer_line_color=(0, 0, 0),
    footer_line_width=0.5,
    
     # 矩形描画用引数
    draw_content_border=True,
    content_border_width=0.5,
    content_border_color=(0,0,0) 
): 
    """
    input_folder_pathで指定されたフォルダ内に存在する全てのPDFファイルに対し、ヘッダーとフッターを付与し、output_folder_pathに出力します。
    
    引数:
        input_folder_path: 入力フォルダのパス
        output_folder_path: 出力フォルダのパス 
        overwrite: 出力ファイルが既に存在する場合に上書きするかどうか
        padding: ヘッダー・フッターのパディング
        resize_original: 元のページを縮小するかどうか

        header_font_name: ヘッダーフォント名
        header_font_size: ヘッダーフォントサイズ
        header_text_color: ヘッダーテキストカラー
        header_height: ヘッダーの高さ
        header_padding: ヘッダーのパディング
        delimiter: ファイル名を分割する区切り文字
        stamp_only_firstpage: 最初のページのみスタンプするかどうか
        draw_header_line: ヘッダー下に線を描画するかどうか
        header_line_color: ヘッダー線の色
        header_line_width: ヘッダー線の幅
        
        footer_font_name: フッターフォント名
        footer_font_size: フッターフォントサイズ
        footer_text_color: フッターテキストカラー
        footer_height: フッターの高さ
        footer_padding: フッターのパディング
        write_pagenum: ページ番号を描画するかどうか
            ※これをFalseに設定し、resize_originalをTrueに設定し、フッターの高さを小さく設定することで、ベゼル的な表示ができる。
        show_total_pages: 総ページ数を表示するかどうか
        draw_footer_line: フッター上に線を描画するかどうか
        footer_line_color: フッター線の色
        footer_line_width: フッター線の幅

        draw_content_border: 元PDFファイルのページコンテンツを囲む矩形を描画するかどうか
        content_border_width: 矩形を囲む線の幅
        content_border_color: 矩形を囲む線の色
    """
            
    if(input_folder_path == output_folder_path):
        logger.fatal(f"入力フォルダと同一のフォルダが出力先に指定されています。{inspect.currentframe().f_code.co_name}を終了します。")
        return
    if(not os.path.exists(input_folder_path)):
        logger.fatal(f"フォルダ：{input_folder_path}は存在しません。{inspect.currentframe().f_code.co_name}を終了します。")
        return
    
    input_folder_abspath = os.path.abspath(input_folder_path)
    output_folder_abspath = os.path.abspath(output_folder_path)
    logger.info(f"フォルダ：{input_folder_abspath}中のPDFファイルにヘッダー及びフッターを付与し、フォルダ{os.path.abspath(input_folder_abspath)}へ出力します。")
       
    if not(os.path.exists(os.path.abspath(output_folder_abspath))):
        os.makedirs(os.path.abspath(output_folder_abspath))
        logger.info(f"フォルダ：{output_folder_abspath}を作成します。")

    for(pdf_in_file) in glob.glob(input_folder_abspath+"/*.pdf"):
        pdf_out_file = output_folder_abspath + "/" + os.path.basename(pdf_in_file)
        logger.info(f"PDFファイル{pdf_in_file}にヘッダーを付与し、{pdf_out_file}へ出力します。")

        header_and_pagenum_to_pdf(
            input_pdf_path=pdf_in_file,
            output_pdf_path=pdf_out_file,
            overwrite=overwrite,
            resize_original=resize_original,
            header_font_name=header_font_name,
            header_font_size=header_font_size,
            header_text_color=header_text_color,
            header_height=header_height,
            header_paddiing=header_padding,
            delimiter=delimiter,
            stamp_only_firstpage=stamp_only_firstpage,
            draw_header_line=draw_header_line,
            header_line_color=header_line_color,
            header_line_width=header_line_width,
            footer_font_name=footer_font_name,
            footer_font_size=footer_font_size,
            footer_text_color=footer_text_color,
            footer_height=footer_height,
            footer_padding=footer_padding,
            write_pagenum=write_pagenum,
            show_total_pages=show_total_pages,
            draw_footer_line=draw_footer_line,
            footer_line_color=footer_line_color,
            footer_line_width=footer_line_width,
            draw_content_border=draw_content_border,
            content_border_width=content_border_width,
            content_border_color=content_border_color
        )        
        
def header_and_frame_to_pdf(
    # ヘッダー・フッターに共通する引数
    input_pdf_path, 
    output_pdf_path,
    overwrite=True,

    # ヘッダー描画用引数
    header_font_name="japan", 
    header_font_size=20, 
    header_text_color=(1, 0, 0), 
    header_height=70,
    header_padding=10,
    delimiter=Delimiters.Space,
    stamp_only_firstpage=False,

    # フッター描画用引数
    footer_height=30,
    footer_padding=3,

    # フレーム（=content_border）描画用引数
    frame_width=0.5,
    frame_color=(0, 0, 0)
    ):
    """
    PDFの各ページにファイル名をヘッダーとして、コンテンツをスケーリングし、元コンテンツの周囲にフレームを描画します。
    引数:
        input_pdf_path: 入力PDFファイルのパス
        output_pdf_path: 出力PDFファイルのパス     
        overwrite: 出力ファイルが既に存在する場合に上書きするかどうか
        header_font_name: ヘッダーフォント名
        header_font_size: ヘッダーフォントサイズ
        header_text_color: ヘッダーテキストカラー
        header_height: ヘッダーの高さ
        header_padding: ヘッダーのパディング
        delimiter: ファイル名を分割する区切り文字
        stamp_only_firstpage: 最初のページのみスタンプするかどうか
        footer_height: フッターの高さ
        footer_padding: フッターのパディング
        frame_width: フレームの幅
        frame_color: フレームの色
    """
    header_and_pagenum_to_pdf(
        input_pdf_path=input_pdf_path,
        output_pdf_path=output_pdf_path,
        overwrite=overwrite,
        header_font_name=header_font_name,
        header_font_size=header_font_size,
        header_text_color=header_text_color,
        header_height=header_height,
        header_paddiing=header_padding,
        delimiter=delimiter,
        stamp_only_firstpage=stamp_only_firstpage,
        footer_height=footer_height,
        footer_padding=footer_padding,
        write_pagenum=False,
        draw_content_border=True,
        content_border_width=frame_width,
        content_border_color=frame_color
    )

def header_and_frame_to_pdf_in_folder(
    input_folder_path,
    output_folder_path,
    overwrite=True,

    # ヘッダー描画用引数
    header_font_name="japan", 
    header_font_size=20, 
    header_text_color=(1, 0, 0), 
    header_height=70,
    header_padding=10,
    delimiter=Delimiters.Space,
    stamp_only_firstpage=False,

    # フッター描画用引数
    footer_height=30,
    footer_padding=3,

    frame_width=0.5,
    frame_color=(0,0,0) 
) :
    """
    input_folder_pathで指定されたフォルダ内に存在する全てのPDFファイルに対し、ヘッダーを付与し、
    コンテンツをスケーリングし、元コンテンツの周囲をフレームで囲み、output_folder_pathに出力します。
    引数:
        input_folder_path: 入力フォルダのパス
        output_folder_path: 出力フォルダのパス
        header_font_name: ヘッダーフォント名
        header_font_size: ヘッダーフォントサイズ
        header_text_color: ヘッダーテキストカラー
        header_height: ヘッダーの高さ
        header_padding: ヘッダーのパディング
        delimiter: ファイル名を分割する区切り文字
        stamp_only_firstpage: 最初のページのみスタンプするかどうか
        footer_height: フッターの高さ
        footer_padding: フッターのパディング
        frame_width: フレームの幅
        frame_color: フレームの色
    """
    if(input_folder_path == output_folder_path):
        logger.fatal(f"入力フォルダと同一のフォルダが出力先に指定されています。{inspect.currentframe().f_code.co_name}を終了します。")
        return
    if(not os.path.exists(input_folder_path)):
        logger.fatal(f"フォルダ：{input_folder_path}は存在しません。{inspect.currentframe().f_code.co_name}を終了します。")
        return
    
    input_folder_abspath=os.path.abspath(input_folder_path)
    output_folder_abspath=os.path.abspath(output_folder_path)

    if not(os.path.exists(os.path.abspath(output_folder_abspath))):
        os.makedirs(os.path.abspath(output_folder_abspath))
        logger.debug(f"フォルダ：{output_folder_abspath}を作成します。")

    for pdf_in_file in glob.glob(input_folder_abspath+"/*.pdf"):
        pdf_out_file = output_folder_abspath + "/" + os.path.basename(pdf_in_file)
        logger.debug(f"PDFファイル{pdf_in_file}にヘッダー及びフレームを付与し、{pdf_out_file}へ出力します。")
        header_and_frame_to_pdf(
            input_pdf_path=pdf_in_file,
            output_pdf_path=pdf_out_file,
            overwrite=overwrite,

            header_font_name=header_font_name,
            header_font_size=header_font_size,
            header_text_color=header_text_color,
            header_height=header_height,
            header_padding=header_padding,
            delimiter=delimiter,
            stamp_only_firstpage=stamp_only_firstpage,

            footer_height=footer_height,
            footer_padding=footer_padding,

            frame_width=frame_width,
            frame_color=frame_color
        )

def concat_pdf(input_folder_path,output_pdf_path,overwrite=True):
    """
    input_folder_pathで指定されたフォルダ内に存在する全てのPDFファイルを結合し、output_pdf_pathに出力します。
    引数:
        input_folder_path: 入力フォルダのパス
        output_pdf_path: 出力PDFファイルのパス
        overwrite: 出力ファイルが既に存在する場合に上書きするかどうか
    """
    if(not os.path.exists(input_folder_path)):
        logger.fatal(f"フォルダ：{input_folder_path}は存在しません。{inspect.currentframe().f_code.co_name}を終了します。")
        return
    if(not overwrite and os.path.exists(output_pdf_path)):
        logger.fatal(f"出力ファイル：{output_pdf_path}は既に存在します。上書きしません。{inspect.currentframe().f_code.co_name}を終了します。")
    
    input_folder_abspath = os.path.abspath(input_folder_path)

    try:
        output_doc = output_doc = pymupdf.open()
        output_doc_paths = []

        for(pdf_in_file) in glob.glob(input_folder_abspath+"/*.pdf"):
            output_doc_paths.append(pdf_in_file)
        
        # 名前に数字が含まれるファイルは、配列のsort()を用いてもうまくいかないので、natsortを用いて自然順にソートする。
        # 例）"10.pdf", "1.pdf", "2.pdf" → "1.pdf", "2.pdf", "10.pdf"
        output_doc_paths = natsort.natsorted(output_doc_paths)

        logger.info(f"PDFファイルを結合し、ファイル{output_doc}へ出力します。")

        for pdf_in_file in output_doc_paths:
            tmp_doc = pymupdf.open(pdf_in_file)
            logger.info(f"ファイル:{tmp_doc}を結合します。")
            output_doc.insert_file(tmp_doc)
            tmp_doc.close()

        output_doc.save(output_pdf_path)
        output_doc.close()
    except Exception as e:
        logger.error(f"エラーが発生しました： {e}")
    finally:
        if('tmp_doc' in locals() and not tmp_doc.is_closed):
            tmp_doc.close()
        if('output_doc' in locals() and not output_doc.is_closed):
            output_doc.close()

