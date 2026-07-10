import streamlit as st
from google import genai
from google.genai import types
from PIL import Image, ImageOps
import docx
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io
import base64

# 1. NASTAVENIE STRÁNKY
st.set_page_config(page_title="Znalec - Fotodokumentácia", layout="wide")
st.title("📸 Profesionálna Fotodokumentácia pre posudky")

# 2. VSTUPNÉ ÚDAJE PRE HLAVIČKU A PÄTU
st.subheader("📝 Základné údaje posudku")
col1, col2, col3, col4 = st.columns(4)
with col1:
    znalec = st.text_input("Meno znalca:", value="Ing. Jozef HRČKA")
with col2:
    datum = st.text_input("Dátum obhliadky:", value="08.07.2026")
with col3:
    objekt = st.text_input("Objekt / Byt (Zápätie):", value="Byt č.7 ,3.p., o.č.6, Javorová 643/6, k.ú. Chrenová, mesto Nitra, okres Nitra")
with col4:
    gemini_key = st.text_input("🔑 Bezplatný Gemini API kľúč:", type="password")

# Inicializácia AI
client = None
if gemini_key:
    try:
        client = genai.Client(api_key=gemini_key)
    except Exception as e:
        st.error(f"Chyba inicializácie AI: {e}")

# 3. HROMADNÉ NAHRÁVANIE FOTIEK
st.subheader("🖼️ Nahrať fotky pre projekt")
uploaded_files = st.file_uploader("Vyberte alebo potiahnite fotky naraz:", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

if 'catalog' not in st.session_state:
    st.session_state.catalog = {}
if 'table_rows' not in st.session_state:
    st.session_state.table_rows = []

if uploaded_files:
    for f in uploaded_files:
        if f.name not in st.session_state.catalog:
            raw_img = Image.open(f)
            corrected_img = ImageOps.exif_transpose(raw_img)
            
            # Konverzia na optimalizovanú miniatúru pre web tabuľku (Base64)
            web_thumb = corrected_img.copy()
            web_thumb.thumbnail((120, 90))
            thumb_buffer = io.BytesIO()
            web_thumb.save(thumb_buffer, format="JPEG")
            encoded_thumb = f"data:image/jpeg;base64,{base64.b64encode(thumb_buffer.getvalue()).decode()}"
            
            # Originálne bajty pre Word
            img_byte_arr = io.BytesIO()
            corrected_img.save(img_byte_arr, format="JPEG")
            final_bytes = img_byte_arr.getvalue()
            
            ai_desc = ""
            if client:
                try:
                    prompt = "Hovoríš po slovensky. Analyzuj tento obrázok zo znaleckej obhliadky nehnuteľnosti. Napíš stručný, profesionálny názov toho, čo vidíš (napr. 'Pohľad na Kúpeľňu', 'Pohľad na Predsieň', 'Pohľad na Rozvádzač RS'). Max 4 slová. Odpavedz LEN týmto popisom."
                    response = client.models.generate_content(
                        model='gemini-2.5-flash',
                        contents=[corrected_img, prompt]
                    )
                    ai_desc = response.text.strip()
                except Exception:
                    ai_desc = f"Pohľad na objekt"
            
            if not ai_desc:
                ai_desc = f"Pohľad na objekt"
                
            st.session_state.catalog[f.name] = final_bytes
            st.session_state.table_rows.append({
                "Názov súboru": f.name,
                "Náhľad (2x3 cm)": encoded_thumb,
                "Popis obrázku (Kliknite a upravte)": ai_desc
            })

    # Odstránenie zmazaných fotiek
    current_names = [f.name for f in uploaded_files]
    st.session_state.table_rows = [r for r in st.session_state.table_rows if r["Názov súboru"] in current_names]

    st.markdown("---")
    st.subheader("📄 Usporiadanie poradia MYŠOU a úprava textu")
    st.info("💡 TIP: Chyťte riadok myšou na ľavom okraji a potiahnite ho hore/dole pre zmenu poradia vo Worde. Text popisu prepíšete dvojklikom.")

    # INTERAKTÍVNA TABUĽKA S OPRAVENÝM VYKRESLENÍM OBRÁZKOV
    edited_df = st.data_editor(
        st.session_state.table_rows,
        column_config={
            "Náhľad (2x3 cm)": st.column_config.ImageColumn(width="small"),
            "Názov súboru": st.column_config.TextColumn(disabled=True),
            "Popis obrázku (Kliknite a upravte)": st.column_config.TextColumn(width="large")
        },
        disabled=["Názov súboru"],
        num_rows="dynamic",
        use_container_width=True,
        key="drag_drop_grid"
    )
    
    # Bezpečné priradenie hodnôt z tabuľky
    if isinstance(edited_df, list):
        rows_to_process = edited_df
    elif isinstance(edited_df, dict) and "edited_rows" in edited_df:
        rows_to_process = st.session_state.table_rows
    else:
        rows_to_process = st.session_state.table_rows

    # Pomocná funkcia na pridanie čiernych okrajov do tabuľky Wordu
    def set_cell_borders(cell):
        tcPr = cell._tc.get_or_add_tcPr()
        tcBorders = OxmlElement('w:tcBorders')
        for border_name in ['top', 'left', 'bottom', 'right']:
            border = OxmlElement(f'w:{border_name}')
            border.set(qn('w:val'), 'single')
            border.set(qn('w:sz'), '4')  # Hrúbka čiary
            border.set(qn('w:space'), '0')
            border.set(qn('w:color'), '000000')  # Čierna
            tcBorders.append(border)
        tcPr.append(tcBorders)

    # 5. GENERÁTOR WORDU (.DOCX) - STOPERCENTNE IMÚNNY VOČI CHYBE S DICT/LIST
    def generate_docx():
        doc = docx.Document()
        
        # Nastavenie úzkych okrajov pre A4
        for section in doc.sections:
            section.top_margin = Inches(0.5)
            section.bottom_margin = Inches(0.5)
            section.left_margin = Inches(0.5)
            section.right_margin = Inches(0.5)
            
            # Hlavička
            header = section.header
            hp = header.paragraphs if header.paragraphs else header.add_paragraph()
            hp.text = f"Znalec: {znalec}\t\tFOTODOKUMENTÁCIA\t\t{datum}"
            hp.alignment = WD_ALIGN_PARAGRAPH.LEFT
            
            # Päta
            footer = section.footer
            fp = footer.paragraphs if footer.paragraphs else footer.add_paragraph()
            fp.text = objekt
            fp.alignment = WD_ALIGN_PARAGRAPH.CENTER

        table = doc.add_table(rows=0, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        for i in range(0, len(rows_to_process), 2):
            cells_img = table.add_row().cells
            cells_txt = table.add_row().cells
            
            set_cell_borders(cells_img[0])
            set_cell_borders(cells_img[1])
            set_cell_borders(cells_txt[0])
            set_cell_borders(cells_txt[1])
            
            # --- ĽAVÁ STRANA ---
            item1 = rows_to_process[i]
            bytes1 = st.session_state.catalog[item1["Názov súboru"]]
            p1 = cells_img[0].paragraphs if cells_img[0].paragraphs else cells_img[0].add_paragraph()
            p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p1.add_run().add_picture(io.BytesIO(bytes1), width=Inches(3.2))
            
            # Vytiahnutie čistého textu popisu z tabuľky
            raw_text1 = item1.get("Popis obrázku (Kliknite a upravte)", "Pohľad na objekt")
            t1 = cells_txt[0].paragraphs if cells_txt[0].paragraphs else cells_txt[0].add_paragraph()
            t1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            t1.add_run(str(raw_text1)).font.size = Pt(9)
            
            # --- PRAVÁ STRANA ---
            if i + 1 < len(rows_to_process):
                item2 = rows_to_process[i+1]
                bytes2 = st.session_state.catalog[item2["Názov súboru"]]
                p2 = cells_img[1].paragraphs if cells_img[1].paragraphs else cells_img[1].add_paragraph()
                p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p2.add_run().add_picture(io.BytesIO(bytes2), width=Inches(3.2))
                
                raw_text2 = item2.get("Popis obrázku (Kliknite a upravte)", "Pohľad na objekt")
                t2 = cells_txt[1].paragraphs if cells_txt[1].paragraphs else cells_txt[1].add_paragraph()
                t2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                t2.add_run(str(raw_text2)).font.size = Pt(9)
                
        bio = io.BytesIO()
        doc.save(bio)
        return bio.getvalue()

    # 6. EXPORT
    st.markdown("---")
    st.subheader("💾 Export a uloženie")
    
    try:
        docx_data = generate_docx()
        st.download_button(
            label="📥 Stiahnuť hotovú prílohu vo WORD (.docx)",
            data=docx_data,
            file_name="Fotodokumentacia.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as error:
        st.error(f"Chyba pri generovaní dokumentu: {error}")
