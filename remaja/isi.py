# isi.py
import streamlit as st
from docx import Document
import re


def extract_isi(doc):
    paragraphs = [re.sub(r'\s+', ' ', p.text).strip()
                  for p in doc.paragraphs if p.text.strip()]

    keywords_acara = [
        "BERNYANYI", "VOTUM", "HUKUM", "DOA", "EPISTEL", "E P I S T E L",
        "PENGAKUAN IMAN", "WARTA", "K O O R", "KOOR", "KHOTBAH", "K H O T B A H",
        "PENGAKUAN DOSA", "JANJI KESELAMATAN", "DOA PENUTUP", "ACARA PANDIDION",
        "MENGUMPULKAN PERSEMBAHAN", "PELEAN"
    ]

    re_nomor = re.compile(r"^\s*(\d{1,2})[\.\s:]+(.*)", re.IGNORECASE)
    re_koor = re.compile(r"K\s*O\s*O\s*R", re.IGNORECASE)
    re_limit_koor = re.compile(
        r"BERNYANYI|KHOTBAH|K\s*H\s*O\s*T\s*B\s*A\s*H|WARTA|DOA|P\s*:|U\s*:|H\s*:|MENGUMPULKAN", re.IGNORECASE)

    sections = []
    current_section = None
    manual_counter = 1

    for text in paragraphs:
        match_num = re_nomor.match(text)
        normalized_text = text.upper().replace(" ", "")

        is_keyword = any(normalized_text.startswith(
            k.replace(" ", "")) for k in keywords_acara)

        if is_keyword and len(text) > 150 and not match_num:
            is_keyword = False

        if match_num or is_keyword:
            if current_section:
                sections.append(current_section)

            if match_num:
                num = int(match_num.group(1))
                head = match_num.group(2).strip()
                if not head:
                    head = text
                manual_counter = num + 1
            else:
                num = manual_counter
                head = text
                manual_counter += 1

            # CARA RADIKAL: Jika judul mengandung "BERNYANYI", potong semua kata sebelumnya
            if "BERNYANYI" in head.upper():
                head = re.sub(r'.*?(?=(?:BERNYANYI))', '',
                              head, flags=re.IGNORECASE).strip()

            current_section = {
                "nomor": num,
                "header_lines": [head],
                "content_lines": [],
                "is_koor": bool(re_koor.search(text))
            }
        else:
            if current_section:
                if current_section["is_koor"]:
                    if re_limit_koor.search(text) and len(text) < 100:
                        current_section["is_koor"] = False
                        current_section["content_lines"].append(text)
                    elif "-" in text or "PKL" in text.upper() or "WIB" in text.upper():
                        current_section["header_lines"].append(text)
                    else:
                        current_section["content_lines"].append(text)
                else:
                    current_section["content_lines"].append(text)

    if current_section:
        sections.append(current_section)

    sections.sort(key=lambda x: x['nomor'])

    formatted_sections = []
    for s in sections:
        formatted_sections.append({
            "nomor": s["nomor"],
            "judul": re.sub(r'\s+', ' ', " ".join(s["header_lines"])),
            "isi": s["content_lines"]
        })

    return formatted_sections


if __name__ == "__main__":
    st.title("Ekstraksi Detail Acara")
    uploaded_file = st.file_uploader("Upload file DOCX", type=["docx"])

    if uploaded_file:
        doc = Document(uploaded_file)
        result = extract_isi(doc)

        for section in result:
            with st.expander(f"Acara {section['nomor']}: {section['judul']}"):
                if section['isi']:
                    for line in section['isi']:
                        if re.match(r"^[PLJSUHM]\s*[:\-]", line) or "BERNYANYI" in line.upper() or \
                           "KHOTBAH" in line.upper().replace(" ", "") or "MENGUMPULKAN" in line.upper():
                            st.markdown(f"**{line}**")
                        elif line.startswith("[") or "---" in line:
                            st.caption(line)
                        else:
                            st.text(line)
                else:
                    st.write("*Detail sudah masuk ke judul.*")
