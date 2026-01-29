from pptx .util import Inches
import streamlit as st
from pptx import Presentation
from pptx .util import Pt, Inches
from pptx .enum .text import PP_ALIGN, MSO_ANCHOR
from pptx .dml .color import RGBColor
from docx import Document
import random
import os
import io
import re
from lxml import etree
import requests

KEYWORDS = {
    "Batak": {
        "1_marende": ["MARENDE", "B.E", "BE.", "HKBP NO"],
        "2_votum": ["VOTUM"],
        "3_patik": ["PATIK", "P A T I K"],
        "4_dosa": ["MANOPOTI DOSA"],
        "5_epistel": ["E P I S T E L", "EPISTEL"],
        "6_iman": ["MANGHATINDANGHON HAPORSEAON"],
        "7_koor": ["K O O R", "KOOR"],
        "8_tingting": ["TINGTING"],
        "9_pelean": ["PAPUNGU PELEAN", "PELEAN"],
        "10_jamita": ["JAMITA", "J A M I T A", "TURPUK"],
        "11_tutup": ["TANGIANG PANGUJUNGI"],
        "stand": ["JONGJONG", "[Jongjong]"]
    },
    "Indonesia": {
        "1_marende": ["BERNYANYI"],
        "2_votum": ["VOTUM"],
        "3_patik": ["HUKUM TAURAT", "HUKUM TUHAN"],
        "4_dosa": ["DOA PENGAKUAN"],
        "5_epistel": ["E P I S T E L", "EPISTEL"],
        "6_iman": ["PENGAKUAN IMAN"],
        "7_koor": ["K O O R", "KOOR"],
        "8_tingting": ["WARTA JEMAAT"],
        "9_pelean": ["PAPUNGU PELEAN", "PELEAN"],
        "10_jamita": ["KHOTBAH", "TURPUK"],
        "11_tutup": ["DOA PENUTUP"],
        "stand": ["BERDIRI", "JONGJONG", "~~~", "~~~ JEMAAT BERDIRI ~~~", "~~ Jemaat Berdiri ~~"]
    }
}

w_pt, h_pt = 960, 720
MARGIN = Pt(30)
THEMES = [(0, 32, 96), (20, 60, 20), (60, 20, 60), (30, 30, 30)]


def extract_metadata(file):
    if file is None:
        return "", "", "Indonesia", True

    doc = Document(io .BytesIO(file .getvalue()))
    paragraphs = [p .text .strip()for p in doc .paragraphs if p .text .strip()]
    full_text = " ".join(paragraphs).upper()

    if "MINGGU"not in full_text:
        return "", "", "Indonesia", False

    minggu_val, topik_val, auto_lang, is_sure = "", "", "Indonesia", False

    for i, line in enumerate(paragraphs):
        line_upper = line .upper()

        if not minggu_val:
            if "PARTURENA" in line_upper or "TATA IBADAH" in line_upper:
                auto_lang = "Batak"if "PARTURENA" in line_upper else "Indonesia"
                is_sure = True
                if i + 1 < len(paragraphs):
                    minggu_val = paragraphs[i + 1].strip()
            elif "MINGGU" in line_upper:
                minggu_val = line .strip()

        if "TOPIK" in line_upper:
            if ":" in line and len(line .split(":", 1)[1].strip()) > 5:
                topik_val = line .split(":", 1)[1].strip()
            elif i + 1 < len(paragraphs):
                topik_val = paragraphs[i + 1].strip()

    return minggu_val, topik_val, (auto_lang if is_sure else "Lainnya"), True


def check_tata_ibadah_validation(file):
    if file is None:
        return True, ""

    doc = Document(io .BytesIO(file .getvalue()))

    paragraphs = [p .text .strip()
                  for p in doc .paragraphs if p .text .strip()][:25]
    full_text = " ".join(paragraphs)
    full_text_upper = full_text .upper()

    if "WARTA" in full_text_upper and "PELAYANAN JEMAAT" in full_text_upper:
        return False, "❌ Dokumen tidak sesuai! Warta Jemaat di kolom Warta Jemaat"

    nama_ibadah = "Tata Ibadah"
    for line in paragraphs:
        if "TATA IBADAH" in line .upper():

            nama_ibadah = line .strip()
            break
        elif "PARTURENA" in line .upper():
            nama_ibadah = line .strip()
            break

    date_match = re .search(
        r'(\d{1,2}[\s\-\/](?:Januari|Februari|Pebruari|Maret|April|Mei|Juni|Juli|Agustus|September|Oktober|November|Desember|[0-9]{1,2})[\s\-\/]\d{4})',
        full_text,
        re .IGNORECASE)
    tgl_info = f" ({date_match .group(1)})"if date_match else ""

    valid_keywords = ["TATA IBADAH", "PARTURENA PARMINGGUON", "PARTURENA"]
    if any(k in full_text_upper for k in valid_keywords):
        return True, f"✅ {nama_ibadah} {tgl_info}."

    return True, f"⚠️ Dokumen terdeteksi, format tidak diketahui{tgl_info}."


def check_warta_validation(file):
    if file is None:
        return True, ""

    doc = Document(io .BytesIO(file .getvalue()))

    paragraphs = [p .text .strip()
                  for p in doc .paragraphs if p .text .strip()][:25]
    full_text = " ".join(paragraphs)
    full_text_upper = full_text .upper()

    wrong_keywords = ["TATA IBADAH", "PARTURENA PARMINGGUON", "PARTURENA"]
    if any(k in full_text_upper for k in wrong_keywords):
        return False, "❌ Dokumen tidak sesuai! Tata Ibadah di kolom Tata Ibadah"

    date_match = re .search(
        r'(\d{1,2}[\s\-\/](?:Januari|Februari|Maret|April|Mei|Juni|Juli|Agustus|September|Oktober|November|Desember|[0-9]{1,2})[\s\-\/]\d{4})',
        full_text,
        re .IGNORECASE)
    tgl_info = f" ({date_match .group(1)})"if date_match else ""

    valid_keywords = ["WARTA", "TINGTING", "PENGUMUMAN", "BERITA JEMAAT"]
    if any(k in full_text_upper for k in valid_keywords):
        return True, f"✅ Warta Jemaat {tgl_info}."

    return True, f"⚠️ Dokumen terdeteksi{tgl_info}, pastikan formatnya benar."


def add_image_to_slide(prs, image_stream):

    try:
        slide_layout = prs .slide_layouts[6]
    except BaseException:
        slide_layout = prs .slide_layouts[5]

    slide = prs .slides .add_slide(slide_layout)

    slide_width = prs .slide_width
    slide_height = prs .slide_height

    from pptx .util import Inches
    margin = Inches(0.1)

    target_width = slide_width - (2 * margin)

    pic = slide .shapes .add_picture(
        image_stream, margin, 0, width=target_width)

    if pic .height < slide_height:
        pic .top = int((slide_height - pic .height)/2)
    else:

        pic .height = slide_height - (2 * margin)
        pic .width = int(
            pic .width * ((slide_height - 2 * margin)/pic .height))
        pic .left = int((slide_width - pic .width)/2)
        pic .top = margin


def process_warta_with_images(prs, w_file):

    if not w_file:
        return

    w_file .seek(0)
    doc = Document(w_file)

    for shape in doc .inline_shapes:
        try:

            rId = shape ._inline .graphic .graphicData .pic .blipFill .blip .embed
            image_part = doc .part .related_parts[rId]

            image_stream = io .BytesIO(image_part .blob)

            add_image_to_slide(prs, image_stream, "LAPORAN KEUANGAN")
        except Exception as e:

            continue


def split_text_by_punctuation(text, limit=20):

    sentences = re .split(r'(?<=[.!?])\s+', text)

    chunks = []
    current_chunk = []
    current_word_count = 0

    for sentence in sentences:

        words_only = re .sub(r'^\d+[\s\.\)]+', '', sentence).split()
        num_words = len(words_only)

        if current_word_count + num_words > limit and current_chunk:
            chunks .append(" ".join(current_chunk))
            current_chunk = [sentence]
            current_word_count = num_words
        else:
            current_chunk .append(sentence)
            current_word_count += num_words

    if current_chunk:
        chunks .append(" ".join(current_chunk))

    return chunks


def add_smart_slide(
        prs,
        content,
        theme_idx,
        footer=None,
        font_size=72,
        is_title=False,
        slide_format="Projector",
        is_warta=False,
        current_bg=1):
    if not content .strip():
        return

    words_to_check = re .sub(r'^\d+[\s\.\)]+', '', content).split()

    if not is_title and len(words_to_check) > 20:
        parts = split_text_by_punctuation(content, limit=20)

        if len(parts) > 1:
            for part in parts:
                add_smart_slide(prs, part, theme_idx, footer, font_size,
                                is_title, slide_format, is_warta, current_bg)
            return

    if slide_format == "Streaming":
        prs .slide_width, prs .slide_height = Pt(960), Pt(540)
        box_h, box_top, box_w = Pt(125), Pt(415), prs .slide_width - Pt(80)
        start_font = Pt(font_size)if font_size < 50 else Pt(45)
    else:
        prs .slide_width, prs .slide_height = Pt(960), Pt(720)
        box_h, box_top, box_w = Pt(570), Pt(50), prs .slide_width - Pt(100)
        start_font = Pt(font_size)

    if slide_format == "Streaming":

        if is_title:
            adjusted_font = Pt(40)
        else:

            char_count = len(content)
            if char_count < 100:
                size = 45
            elif char_count < 200:
                size = 40
            else:
                size = 36
            adjusted_font = Pt(size)
    else:

        text_clean = content .strip()
        char_count = len(text_clean)

        if is_title:

            size = 68
        else:

            if char_count <= 80:
                size = 72
            else:

                reduction = (char_count - 80)//4
                size = 72 - reduction

            if size < 44:
                size = 44

        adjusted_font = Pt(size)

    slide = prs .slides .add_slide(prs .slide_layouts[6])

    if slide_format == "Streaming":

        slide .background .fill .solid()
        slide .background .fill .fore_color .rgb = RGBColor(255, 255, 255)

        if is_warta:

            txBox = slide .shapes .add_textbox(Pt(50), Pt(
                50), prs .slide_width - Pt(100), prs .slide_height - Pt(100))
            font_color_main = RGBColor(0, 0, 0)

        else:

            rect = slide .shapes .add_shape(
                1, 0, box_top, prs .slide_width, box_h)
            rect .fill .solid()
            rect .fill .fore_color .rgb = RGBColor(0, 255, 0)
            rect .line .fill .background()

            txBox = slide .shapes .add_textbox(
                Pt(40), box_top + Pt(10), box_w, box_h - Pt(20))
            fill = txBox .fill
            fill .solid()
            fill .fore_color .rgb = RGBColor(0, 0, 0)
            try:
                from pptx .oxml import parse_xml

                fill ._xPr .solidFill .srgbClr .append(parse_xml(
                    '<a:alpha val="100000" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>'))
            except BaseException:
                pass
            font_color_main = RGBColor(255, 255, 255)

    elif is_warta:

        slide .background .fill .solid()
        slide .background .fill .fore_color .rgb = RGBColor(255, 255, 255)
        txBox = slide .shapes .add_textbox(Pt(50), box_top, box_w, box_h)
        font_color_main = RGBColor(0, 0, 0)

    else:

        bg_path = os .path .join("pics", f"{current_bg}.jpg")
        if os .path .exists(bg_path):
            pic = slide .shapes .add_picture(
                bg_path, 0, 0, width=prs .slide_width, height=prs .slide_height)
            slide .shapes ._spTree .remove(pic ._element)
            slide .shapes ._spTree .insert(2, pic ._element)

            overlay = slide .shapes .add_shape(
                1, 0, 0, prs .slide_width, prs .slide_height)
            overlay .fill .solid()
            overlay .fill .fore_color .rgb = RGBColor(0, 0, 0)
            overlay .line .fill .background()
            try:
                from pptx .oxml import parse_xml
                overlay .fill ._xPr .solidFill .srgbClr .append(parse_xml(
                    '<a:alpha val="35000" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>'))
            except BaseException:
                pass
            slide .shapes ._spTree .remove(overlay ._element)
            slide .shapes ._spTree .insert(3, overlay ._element)
        else:
            bg_color = RGBColor(*THEMES[theme_idx % len(THEMES)])
            slide .background .fill .solid()
            slide .background .fill .fore_color .rgb = bg_color

        txBox = slide .shapes .add_textbox(Pt(50), box_top, box_w, box_h)
        font_color_main = RGBColor(255, 255, 255)

    tf = txBox .text_frame
    tf .word_wrap = True
    tf .vertical_anchor = MSO_ANCHOR .MIDDLE

    p = tf .paragraphs[0]
    p .alignment = PP_ALIGN .CENTER

    run = p .add_run()
    run .text = content .strip()

    run .font .name = 'Amasis MT Pro Black'if slide_format == "Streaming"else 'Verdana'
    run .font .bold = True
    run .font .size = adjusted_font

    run .font .color .rgb = font_color_main

    if footer and slide_format == "Projector":
        f_box = slide .shapes .add_textbox(
            0, Pt(620), prs .slide_width, Pt(80))
        pf = f_box .text_frame .paragraphs[0]
        pf .text = f"--- {footer} ---"
        pf .font .size = Pt(32)
        pf .font .color .rgb = RGBColor(255, 255, 50)
        pf .alignment = PP_ALIGN .CENTER


def format_be_title(line, slide_format="Projector"):

    line_clean = re .sub(r'♫|♪|\v|\n|\r|\t', ' ', line)
    line_clean = " ".join(line_clean .split())

    is_song = re .search(r'(B\.?E\.?|HKBP|NO\.?)\s*\d+',
                         line_clean, re .IGNORECASE)

    if is_song:

        line_song = re .sub(r'^(MARENDE|BERNYANYI|PAPUNGU PELEAN)\s+',
                            '', line_clean, flags=re .IGNORECASE).strip()

        match = re .search(
            r"(B\.?E\.?\s*(?:HKBP\.?)?)\s*(?:No\.?\s*)?([\d\s\.\-\:–,]+)(.*)",
            line_song,
            re .IGNORECASE
        )

        if match:
            tag = "B.E. HKBP"
            nomor = re .sub(r'[,\.\s]+$', '', match .group(2).strip()).strip()
            judul_raw = match .group(3).strip()

            judul_clean = re .sub(
                r'[\(\[\{]?(?:BL|B\.?L|BE|B\.?E|BN|B\.?N)\.?\s*\d+.*[\)\]\}]?',
                '',
                judul_raw,
                flags=re .IGNORECASE).strip()

            judul_clean = re .sub(
                r'\b(?:BL|B\.?L|BE|B\.?E|BN|B\.?N)\.?\s*\d+.*$',
                '',
                judul_clean,
                flags=re .IGNORECASE).strip()

            judul_final = re .sub(
                r'[“”"\'‘’\(\)\[\]]', '', judul_clean).strip()

            judul_final = re .sub(r'^[:\-–\s,]+', '', judul_final).strip()

            formatted_judul = f"“{judul_final .upper()}”"if judul_final else ""

            if slide_format == "Streaming":
                res = f"{tag} No. {nomor}"
                return f"{res}\n{formatted_judul}".strip()
            else:
                return f"{tag}\nNo. {nomor}\n{formatted_judul}".strip()

    return line_clean .upper()


def get_clean_content(file):
    if file is None:
        return []

    doc = Document(io .BytesIO(file .read()))
    content = []

    for p in doc .paragraphs:

        txt = re .sub(r'♫|♪', '', p .text).strip()
        if txt:
            content .append(txt)

    for table in doc .tables:
        for row in table .rows:

            row_data = []
            for cell in row .cells:
                clean_cell = re .sub(r'♫|♪', '', cell .text).strip()
                if clean_cell and (not row_data or clean_cell != row_data[-1]):
                    row_data .append(clean_cell)

            if row_data:

                content .append(" | ".join(row_data))

    return content


def check_match(text, keyword_list):
    text_up = text .upper()
    return any(k .upper() in text_up for k in keyword_list)


def generate_logic(
        tata_lines,
        warta_lines,
        minggu,
        topik,
        bahasa,
        slide_format,
        w_file):
    if slide_format == "Streaming":

        return generate_streaming_logic(
            tata_lines, warta_lines, minggu, topik, bahasa, w_file)
    else:

        return generate_projector_logic(
            tata_lines, warta_lines, minggu, topik, bahasa, w_file)


def generate_projector_logic(
        tata_lines,
        warta_lines,
        minggu,
        topik,
        bahasa,
        w_file):

    prs = Presentation()

    prs .slide_width, prs .slide_height = Pt(960), Pt(720)
    slide_format = "Projector"
    active_bg = random .randint(1, 214)

    add_smart_slide(
        prs, f"{minggu .upper()}\n\nTOPIK:\n“{topik .upper()}”",
        0, is_title=True, slide_format=slide_format, current_bg=active_bg
    )

    key = KEYWORDS[bahasa]
    ALL_STOPS = []
    for category in key:
        if category != "stand":
            ALL_STOPS .extend(key[category])

    done_sections = set()
    theme_idx = 0
    i = 0

    while i < len(tata_lines):
        line = tata_lines[i].strip()
        if not line:
            i += 1
            continue

        line_clean = re .sub(r'^\d+[\s\.]+', '', line).strip()

        if check_match(
                line_clean,
                key["1_marende"]) or check_match(
                line_clean,
                key["9_pelean"]):
            theme_idx += 1

            active_bg = random .randint(1, 214)

            header_text = format_be_title(line_clean)if bahasa == "Batak" and check_match(
                line_clean,
                key["1_marende"])else re .sub(
                r'BERNYANYI|MARENDE|PAPUNGU PELEAN',
                '',
                line_clean,
                flags=re .IGNORECASE).strip().upper()

            add_smart_slide(prs, header_text, theme_idx,
                            slide_format=slide_format, current_bg=active_bg)

            i += 1
            all_lines = []

            while i < len(tata_lines):
                txt = tata_lines[i].strip()

                if check_match(
                        re .sub(
                            r'^\d+[\s\.]+',
                            '',
                            txt).strip(),
                        ALL_STOPS):
                    break

                is_stand = re .search(
                    r'BERDIRI|JONGJONG|~~', txt, re .IGNORECASE)

                if is_stand:

                    lyric_part = re .sub(
                        r'\[?JONGJONG\]?|\[?BERDIRI\]?|~~',
                        '',
                        txt,
                        flags=re .IGNORECASE).strip()
                    lyric_part = re .sub(r'♫|♪', '', lyric_part).strip()

                    if lyric_part:
                        all_lines .append(lyric_part)

                    if all_lines:
                        for n in range(0, len(all_lines), 2):
                            add_smart_slide(prs,
                                            "\n".join(all_lines[n:n + 2]),
                                            theme_idx,
                                            slide_format=slide_format,
                                            current_bg=active_bg)
                        all_lines = []

                    stand_text = "JONGJONG"if bahasa == 'Batak'else 'BERDIRI'
                    add_smart_slide(
                        prs,
                        f"*** {stand_text} ***",
                        theme_idx,
                        slide_format=slide_format,
                        current_bg=active_bg)

                elif not txt:

                    if all_lines:
                        for n in range(0, len(all_lines), 2):
                            add_smart_slide(prs,
                                            "\n".join(all_lines[n:n + 2]),
                                            theme_idx,
                                            slide_format=slide_format,
                                            current_bg=active_bg)
                        all_lines = []
                else:

                    sub_lines = txt .split('\n')
                    for sl in sub_lines:
                        clean_sl = re .sub(r'♫|♪', '', sl).strip()

                        is_book_code = re .match (
                            r'^(BL|B\.L|BE|B\.E|BN|B\.N)\.?\s*\d+\s*$', clean_sl, re .IGNORECASE)
                        if is_book_code:
                            continue

                        if clean_sl:
                            all_lines .append(clean_sl)

                i += 1

            if all_lines:
                for n in range(0, len(all_lines), 2):
                    add_smart_slide(prs,
                                    "\n".join(all_lines[n:n + 2]),
                                    theme_idx,
                                    slide_format=slide_format,
                                    current_bg=active_bg)
            continue

        elif any(check_match(line_clean, key[k])for k in ["3_patik"]):
            active_bg = random .randint(1, 214)

            found_ayat = ""
            for offset in range(1, 6):
                if i + offset < len(tata_lines):
                    next_line = tata_lines[i + offset].strip()
                    if not next_line or next_line .startswith("("):
                        continue

                    clean_line = re .sub(
                        r'^[PU]\s*[:：]\s*',
                        '',
                        next_line,
                        flags=re .IGNORECASE).strip()

                    match_hukum = re .search(
                        r'((HUKUM|PATIK)\s+(YANG\s+)?(PERTAMA|I|KEDUA|II|KE[\w]+).*)',
                        clean_line,
                        re .IGNORECASE)

                    match_ref = re .search(
                        r'([A-Za-z0-9\s]+\d+\s*[:：]\s*[\d\-\s]+)', clean_line)

                    if match_hukum:
                        found_ayat = match_hukum .group(1).strip().upper()

                        if ":" in found_ayat and len(
                                found_ayat .split(":")[1]) > 50:
                            found_ayat = found_ayat .split(":")[0].strip()
                        break
                    elif match_ref:
                        found_ayat = match_ref .group(1).strip().upper()

                        found_ayat = re .sub(
                            r'^I\s+MA\s+', '', found_ayat).strip()
                        break

                    if "HUKUM" in clean_line .upper() or "PATIK" in clean_line .upper():

                        parts = re .split(r'yaitu|yakni|i ma',
                                          clean_line, flags=re .IGNORECASE)
                        if len(parts) > 1:
                            found_ayat = parts[1].strip().upper()
                            break

            nama_kategori = key["3_patik"][0].upper()

            found_ayat = found_ayat .rstrip(".:")

            judul_gabungan = f"{nama_kategori}\n{found_ayat}"if found_ayat else nama_kategori
            add_smart_slide(prs, judul_gabungan, 0, font_size=72,
                            slide_format=slide_format, current_bg=active_bg)

            i += 1
            while i < len(tata_lines):
                target_line = tata_lines[i].strip().upper()
                if not target_line:
                    i += 1
                    continue

                if any(
                    stop in target_line for stop in [
                        "MARENDE",
                        "BE.",
                        "BN.",
                        "KJ.",
                        "PKJ.",
                        "MANOPOTI",
                        "DOSA",
                        "4.",
                        "5."]):
                    i -= 1
                    break
                i += 1

            i += 1
            continue

        elif check_match(line_clean, key["3_patik"]) and "3_patik"not in done_sections:
            theme_idx += 1
            add_smart_slide(
                prs,
                line_clean .rstrip(':').upper(),
                theme_idx,
                slide_format=slide_format,
                current_bg=active_bg)
            done_sections .add("3_patik")

            i += 1
            while i < len(tata_lines):
                l_res = tata_lines[i].strip()
                if not l_res:
                    i += 1
                    continue
                if check_match(
                    re .sub(
                        r'^\d+[\s\.]+',
                        '',
                        l_res).strip(),
                    key["1_marende"]) or re .match (
                    r'^(4|5|6)[\s\.]+',
                        l_res):
                    break

                if not re .search(
                    r'JONGJONG|HUNDUL|~~~',
                    l_res,
                        re .IGNORECASE):
                    add_smart_slide(
                        prs,
                        l_res,
                        theme_idx,
                        font_size=48,
                        slide_format=slide_format,
                        current_bg=active_bg)
                i += 1
            continue

        elif check_match(line_clean, key["4_dosa"]):
            theme_idx += 1
            judul_dosa = "DOA PENGAKUAN DOSA DAN\nJANJI KESELAMATAN"if bahasa == "Indonesia"else "MANOPOTI DOSA DOHOT\nBAGABAGA HASESAAN NI DOSA"
            add_smart_slide(
                prs,
                judul_dosa,
                theme_idx,
                font_size=72,
                slide_format=slide_format,
                current_bg=active_bg)
            done_sections .add("4_dosa")

            i += 1
            isi_dosa_full = []
            while i < len(tata_lines):
                l_res = tata_lines[i].strip()
                if not l_res:
                    i += 1
                    continue
                if any(s in l_res .upper()
                       for s in ["BERNYANYI", "MARENDE", "EPISTEL", ""]):
                    break
                if not re .search(
                    r'BERDIRI|JONGJONG|HUNDUL|~~~|\[.*?\]',
                    l_res,
                        re .IGNORECASE):
                    isi_dosa_full .append(l_res)
                i += 1

            if isi_dosa_full:
                full_text = " ".join(isi_dosa_full)
                words = full_text .split()
                n = len(words)

                parts = [" ".join(words[:n // 3]),
                         " ".join(words[n // 3:2 * n // 3]),
                         " ".join(words[2 * n // 3:])]
                for part in parts:
                    if part .strip():
                        add_smart_slide(
                            prs,
                            part .strip(),
                            theme_idx,
                            font_size=40,
                            slide_format=slide_format,
                            current_bg=active_bg)
            continue

        elif check_match(line_clean, key["5_epistel"]) and "5_epistel"not in done_sections:
            theme_idx += 1
            add_smart_slide(
                prs,
                line_clean .rstrip(':').upper(),
                theme_idx,
                slide_format=slide_format,
                current_bg=active_bg)
            done_sections .add("5_epistel")
            i += 1
            while i < len(tata_lines):
                l_res = tata_lines[i].strip()
                if not l_res:
                    i += 1
                    continue
                l_check = re .sub(r'^\d+[\s\.]+', '', l_res).strip().upper()
                if check_match(
                    l_check,
                    key["1_marende"]) or check_match(
                    l_check,
                    key["6_iman"]) or re .match (
                    r'^\d+[\s\.]+(BERNYANYI|MARENDE|IMAN|MANGHATINDANGHON)',
                        l_res .upper()):
                    break
                if not re .search(
                    r'JONGJONG|HUNDUL|~~~',
                    l_res,
                        re .IGNORECASE):
                    add_smart_slide(
                        prs,
                        l_res,
                        theme_idx,
                        slide_format=slide_format,
                        current_bg=active_bg)
                i += 1
            continue

        elif check_match(line_clean, key["8_tingting"]):
            theme_idx += 1
            add_smart_slide(
                prs,
                "WARTA JEMAAT",
                theme_idx,
                font_size=72,
                slide_format=slide_format,
                is_warta=True)

            if warta_lines:
                for w in warta_lines:
                    w_strip = w .strip()
                    if not w_strip:
                        continue
                    w_upper = w_strip .upper()
                    if "PELAYAN MINGGU INI" in w_upper or "PENGKHOTBAH" in w_upper:
                        break
                    clean_text = re .sub(r'\s+', ' ', w_strip)
                    if len(clean_text) > 2:
                        add_smart_slide(
                            prs,
                            clean_text,
                            theme_idx,
                            font_size=40,
                            slide_format=slide_format,
                            is_warta=True)

            if w_file is not None:
                try:
                    w_file .seek(0)
                    doc_w = Document(w_file)

                    image_parts = doc_w .part .related_parts
                    found_any = False

                    for rel_id in image_parts:
                        part = image_parts[rel_id]
                        if "image" in part .content_type:
                            image_stream = io .BytesIO(part .blob)

                            add_image_to_slide(prs, image_stream)

                            found_any = True

                    if not found_any:
                        for shape in doc_w .inline_shapes:
                            rId = shape ._inline .graphic .graphicData .pic .blipFill .blip .embed
                            image_stream = io .BytesIO(
                                doc_w .part .related_parts[rId].blob)
                            add_image_to_slide(
                                prs, image_stream, "LAPORAN KEUANGAN (BACKUP)")

                except Exception as e:
                    st .error(f"⚠️ Sistem gagal membongkar gambar: {e}")

            i += 1
            while i < len(tata_lines) and not check_match(
                    re .sub(r'^\d+[\s\.]+', '', tata_lines[i]).strip(), ALL_STOPS):
                i += 1
            continue

        elif check_match(line_clean, key .get("7_koor", ["KOOR", "K O O R", "PADUAN SUARA"])):
            active_bg = random .randint(1, 214)

            first_line = re .sub(r'^\d+[\s\.]+', '', line).strip()
            koor_list = [first_line]

            i += 1

            while i < len(tata_lines):
                next_l = tata_lines[i].strip()
                if not next_l:
                    i += 1
                    continue

                if check_match(
                        re .sub(
                            r'^\d+[\s\.]+',
                            '',
                            next_l).strip(),
                        ALL_STOPS):
                    break

                koor_list .append(next_l)
                i += 1

            full_koor_text = "\n".join(koor_list)

            add_smart_slide(
                prs,
                full_koor_text,
                0,
                font_size=45,
                slide_format=slide_format,
                current_bg=active_bg
            )
            continue

        elif check_match(line_clean, key .get("7_jamita", ["JAMITA", "KHOTBAH"])):
            theme_idx += 1
            active_bg = random .randint(1, 214)

            add_smart_slide(
                prs,
                line_clean .upper(),
                theme_idx,
                font_size=72,
                slide_format=slide_format,
                current_bg=active_bg)

            i += 1
            isi_jamita = []
            while i < len(tata_lines):
                l_res = tata_lines[i].strip()
                if not l_res:
                    i += 1
                    continue

                l_check = re .sub(r'^\d+[\s\.]+', '', l_res).strip().upper()

                if check_match(
                        l_check, ALL_STOPS) or check_match(
                        l_check, key["1_marende"]):
                    break

                if not re .search(
                    r'JONGJONG|HUNDUL|~~~',
                    l_res,
                        re .IGNORECASE):
                    isi_jamita .append(l_res)
                i += 1

            if isi_jamita:
                full_nats = " ".join(isi_jamita)

                parts = split_text_by_punctuation(full_nats, limit=20)
                for part in parts:
                    add_smart_slide(
                        prs,
                        part,
                        theme_idx,
                        font_size=40,
                        slide_format=slide_format,
                        current_bg=active_bg)
            continue

        elif check_match(line_clean, key["11_tutup"]) and "11_tutup"not in done_sections:
            theme_idx += 1
            active_bg = random .randint(1, 214)

            add_smart_slide(
                prs,
                line_clean .rstrip(':').upper(),
                theme_idx,
                slide_format=slide_format,
                current_bg=active_bg)
            done_sections .add("11_tutup")
            i += 1
            while i < len(tata_lines):
                l_res = tata_lines[i].strip()
                if not l_res:
                    i += 1
                    continue
                l_check = re .sub(r'^\d+[\s\.]+', '', l_res).strip().upper()
                if check_match(
                    l_check,
                    key["1_marende"]) or check_match(
                    l_check,
                    key["6_iman"]) or re .match (
                    r'^\d+[\s\.]+(BERNYANYI|MARENDE|IMAN|MANGHATINDANGHON)',
                        l_res .upper()):
                    break
                if not re .search(
                    r'JONGJONG|HUNDUL|~~~',
                    l_res,
                        re .IGNORECASE):
                    add_smart_slide(
                        prs,
                        l_res,
                        theme_idx,
                        slide_format=slide_format,
                        current_bg=active_bg)
                i += 1
            continue

        i += 1

    add_smart_slide(prs, "HAPPY SUNDAY", 0, is_title=True,
                    slide_format=slide_format, current_bg=active_bg)

    buf = io .BytesIO()
    prs .save(buf)
    buf .seek(0)
    return buf .getvalue()


def generate_streaming_logic(
        tata_lines,
        warta_lines,
        minggu,
        topik,
        bahasa,
        w_file):
    theme_idx = 0

    prs = Presentation()
    prs .slide_width, prs .slide_height = Pt(960), Pt(540)
    slide_format = "Streaming"

    add_smart_slide(
        prs, f"{minggu .upper()}\n{topik .upper()}",
        0, is_title=True, slide_format=slide_format
    )

    key = KEYWORDS[bahasa]
    ALL_STOPS = []
    for category in key:
        if category != "stand":
            ALL_STOPS .extend(key[category])

    i = 0
    while i < len(tata_lines):
        line = tata_lines[i].strip()
        if not line:
            i += 1
            continue

        line_clean = re .sub(r'^\d+[\s\.]+', '', line).strip()

        if check_match(
                line_clean,
                key["1_marende"]) or check_match(
                line_clean,
                key["9_pelean"]):
            header_text = format_be_title(
                line_clean,
                slide_format=slide_format)if bahasa == "Batak"else line_clean .upper()
            if header_text:
                add_smart_slide(prs, header_text, 0,
                                font_size=40, slide_format=slide_format)
            i += 1
            current_bait = []
            while i < len(tata_lines):
                txt = tata_lines[i].strip()
                if check_match(
                        re .sub(
                            r'^\d+[\s\.]+',
                            '',
                            txt).strip(),
                        ALL_STOPS):
                    break
                if re .search(r'BERDIRI|JONGJONG|~~', txt, re .IGNORECASE):
                    if current_bait:
                        for line_lirik in current_bait:
                            add_smart_slide(
                                prs, line_lirik, 0, font_size=40, slide_format=slide_format)
                        current_bait = []

                    stand_text = "*** JONGJONG ***"if bahasa == 'Batak'else '*** BERDIRI ***'
                    add_smart_slide(prs, stand_text, 0,
                                    font_size=40, slide_format=slide_format)

                elif not txt:
                    if current_bait:
                        for line_lirik in current_bait:
                            add_smart_slide(
                                prs, line_lirik, 0, font_size=40, slide_format=slide_format)
                        current_bait = []
                else:
                    clean_txt = re .sub(r'♫|♪', '', txt).strip()
                    if clean_txt:
                        current_bait .append(clean_txt)
                i += 1
            if current_bait:
                for line_lirik in current_bait:
                    add_smart_slide(prs, line_lirik, 0,
                                    font_size=40, slide_format=slide_format)
            continue

        target_keys = ["2_votum", "3_patik", "4_dosa",
                       "5_epistel", "6_iman", "11_tutup"]

        is_matched = False
        judul_bagian = ""

        for k in target_keys:
            if check_match(line_clean, key[k]):
                is_matched = True

                if k == "11_tutup":

                    main_title = re .split(r'[:–\-]', line_clean)[0].strip()
                    judul_bagian = main_title .upper()
                elif k == "2_votum":
                    judul_bagian = "VOTUM - INTROITUS - DOA"if bahasa == "Indonesia"else "VOTUM - INTROITUS - TANGIANG"
                elif k == "6_iman":
                    judul_bagian = "PENGAKUAN IMAN RASULI"if bahasa == "Indonesia"else "MANGHATINDANGHON HAPORSEAON"
                else:

                    judul_bagian = re .split(
                        r'[:–\-]', line_clean)[0].strip().upper()
                break

        if is_matched:

            add_smart_slide(prs, judul_bagian, 0, font_size=40,
                            is_title=True, slide_format=slide_format)

            i += 1
            while i < len(tata_lines):
                next_line = tata_lines[i].strip()
                if not next_line:
                    i += 1
                    continue

                next_clean = re .sub(
                    r'^\d+[\s\.]+', '', next_line).strip().upper()

                if check_match(
                    next_clean,
                    key["1_marende"]) or check_match(
                    next_clean,
                    key["9_pelean"]) or re .match (
                    r'^\d+[\s\.]+',
                        next_line):
                    break
                i += 1
            continue

        elif check_match(line_clean, key["8_tingting"]):
            theme_idx += 1
            add_smart_slide(
                prs,
                "WARTA JEMAAT",
                theme_idx,
                font_size=72,
                slide_format=slide_format,
                is_warta=True)

            if warta_lines:
                for w in warta_lines:
                    w_strip = w .strip()
                    if not w_strip:
                        continue
                    w_upper = w_strip .upper()
                    if "PELAYAN MINGGU INI" in w_upper or "PENGKHOTBAH" in w_upper:
                        break
                    clean_text = re .sub(r'\s+', ' ', w_strip)
                    if len(clean_text) > 2:
                        add_smart_slide(
                            prs,
                            clean_text,
                            theme_idx,
                            font_size=40,
                            slide_format=slide_format,
                            is_warta=True)

            if w_file is not None:
                try:
                    w_file .seek(0)
                    doc_w = Document(w_file)

                    image_parts = doc_w .part .related_parts
                    found_any = False

                    for rel_id in image_parts:
                        part = image_parts[rel_id]
                        if "image" in part .content_type:
                            image_stream = io .BytesIO(part .blob)

                            add_image_to_slide(prs, image_stream)

                            found_any = True

                    if not found_any:
                        for shape in doc_w .inline_shapes:
                            rId = shape ._inline .graphic .graphicData .pic .blipFill .blip .embed
                            image_stream = io .BytesIO(
                                doc_w .part .related_parts[rId].blob)
                            add_image_to_slide(
                                prs, image_stream, "LAPORAN KEUANGAN (BACKUP)")

                except Exception as e:
                    st .error(f"⚠️ Sistem gagal membongkar gambar: {e}")

            i += 1
            while i < len(tata_lines) and not check_match(
                    re .sub(r'^\d+[\s\.]+', '', tata_lines[i]).strip(), ALL_STOPS):
                i += 1
            continue

        i += 1

    add_smart_slide(prs, "HAPPY SUNDAY", 0,
                    font_size=40, slide_format=slide_format)
    buf = io .BytesIO()
    prs .save(buf)
    buf .seek(0)
    return buf .getvalue()


st .set_page_config(page_title="Slidenauli", layout="centered")

st .markdown("""
    <style>
    header {visibility: hidden;}
    footer {visibility: hidden;}

    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&family=Archivo+Black&display=swap');

    html, body, [class*="st-"] {
        font-family: 'Inter', sans-serif !important;
    }

    h1 {
        font-family: 'Archivo Black', sans-serif !important;
        text-transform: uppercase;
        letter-spacing: -2px;
        color: var(--text-color);
        margin-bottom: 0px;
    }

    /* Tombol Utama & Download (Disamakan Ukurannya) */
    div.stButton > button, div.stDownloadButton > button {
        width: 100% !important;
        border-radius: 8px !important;
        height: 3.5rem !important; /* Ukuran tinggi sama */
        background-color: #ff0055 !important;
        color: white !important;
        border: none !important;
        font-weight: 700 !important;
        font-size: 16px !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        transition: transform 0.1s ease;
    }

    div.stButton > button:hover, div.stDownloadButton > button:hover {
        opacity: 0.9;
        border: none !important;
    }

    /* Input Styling */
    .stTextInput > div > div > input, .stTextArea > div > div > textarea {
        border-radius: 8px !important;
    }

    /* Notifikasi Mungil */
    .small-notif {
        padding: 10px 15px;
        border-radius: 8px;
        background-color: rgba(255, 0, 85, 0.1);
        border: 1px solid #ff0055;
        color: var(--text-color);
        font-size: 14px;
        font-weight: 500;
        margin-bottom: 15px;
    }
    </style>
    """, unsafe_allow_html=True)


st .markdown("<h1>SLIDE<span style='color:#ff0055;'>NAULI</span> 🚀</h1>",
             unsafe_allow_html=True)
st .caption("Document converter by Multimedia HKBP Perum 2 Bekasi")
st .write("---")


st .markdown("###  1. Dokumen")
col1, col2 = st .columns(2)
with col1:
    t_file = st .file_uploader("📂 Upload Tata Ibadah",
                               type=["docx"], key="t_up")

    if t_file:

        is_valid_tata, message_tata = check_tata_ibadah_validation(t_file)

        if not is_valid_tata:

            st .error(message_tata)
        elif "✅" in message_tata:
            st .success(message_tata)
        else:
            st .warning(message_tata)

with col2:
    w_file = st .file_uploader(
        "📂 Buka file warta simpan sebagai docx lalu upload di sini",
        type=["docx"],
        key="w_up")

    if w_file:

        is_valid, message = check_warta_validation(w_file)

        if not is_valid:

            st .error(message)
        elif "✅" in message:

            st .success(message)
        else:

            st .warning(message)


st .write("")
st .markdown("### 2. Pengaturan")

default_minggu = ""
default_topik = ""
default_lang_idx = 1

if t_file:
    t_id = f"{t_file .name}_{t_file .size}"
    if f"meta_{t_id}"not in st .session_state:
        m, t, l, v = extract_metadata(t_file)
        st .session_state[f"meta_{t_id}"] = (m, t, l, v)

    m_val, t_val, l_val, _ = st .session_state[f"meta_{t_id}"]
    default_minggu = m_val
    default_topik = t_val
    default_lang_idx = 0 if l_val == "Batak"else 1

c1, c2 = st .columns(2)
with c1:
    final_lang = st .radio("Bahasa (Otomatis)", [
        "Batak", "Indonesia"], index=default_lang_idx, horizontal=True)
with c2:
    final_format = st .selectbox("Format", ["Projector", "Streaming"])

final_minggu = st .text_input("Nama Minggu / Acara", value=default_minggu)
final_topik = st .text_area("Tema Utama / Topik", value=default_topik)


st .write("---")
if st .button("🚀 Proses & Buat Slide"):
    if not t_file:
        st .warning("⚠️ Silakan unggah file Tata Ibadah dulu ya.")
    else:
        with st .spinner("⏳ Lagi diproses..."):
            try:
                t_data = get_clean_content(t_file)
                w_data = get_clean_content(w_file)if w_file else []

                ppt_bytes = generate_logic(
                    t_data,
                    w_data,
                    final_minggu,
                    final_topik,
                    final_lang,
                    final_format,
                    w_file)

                st .markdown(
                    f"<div class='small-notif'>✅ Selesai! Slide <b>{final_format}</b> siap diunduh.</div>",
                    unsafe_allow_html=True)

                st .download_button(
                    label=f"📥 Download {final_format} PPTX",
                    data=ppt_bytes,
                    file_name=f"{final_format}_{final_minggu}.pptx",
                    use_container_width=True
                )
            except Exception as e:
                st .error(f"❌ Gagal: {e}")

st .write("")
st .markdown(
    "<p style='text-align: center; opacity: 0.5; font-size: 11px;'>Made with ❤️ by Mulmed Team</p>",
    unsafe_allow_html=True)
