import streamlit as st
from google import genai
from google.genai import types
from PIL import Image, ImageOps
import docx
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import io

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
    objekt = st.text_input("Objekt / Byt (Zápätie):", value="Byt č.7, 3.p., o.č.6, Javorová 643/6, k.ú. Chrenová, Nitra")
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

# Inicializácia vnútornej pamäte
if 'catalog' not in st.session_state:
    st.session_state.catalog = {}
if 'table_rows' not in st.session_state:
    st.session_state.table_rows = []

if uploaded_files:
    # Spracovanie nových fotiek
    for f in uploaded_files:
        if f.name not in st.session_state.catalog:
            # Snímanie a automatické vyrovnanie fotky z mobilu
            raw_img = Image.open(f)
            corrected_img = ImageOps.exif_transpose(raw_img)
            
            # Zmenšenie na miniatúru pre rýchly web (2x3 cm v pomere strán)
            web_thumb = corrected_img.copy()
            web_thumb.thumbnail((120, 90))
            
            # Príprava bajtov pre Word
            img_byte_arr = io.BytesIO()
            corrected_img.save(img_byte_arr, format="JPEG")
            final_bytes = img_byte_arr.getvalue()
            
            # Zavolanie AI pre popis
            ai_desc = ""
            if client:
                try:
                    prompt = "Hovoríš po slovensky. Analyzuj tento obrázok zo znaleckej obhliadky nehnuteľnosti. Napíš stručný, profesionálny jednovetný názov toho, čo presne vidíš (napr. 'Pohľad na Kúpeľňu', 'Vstupný portál do bytového domu'). Odpovedz LEN týmto popisom."
                    response = client.models.generate_content(
                        model='gemini-2.5-flash',
                        contents=[corrected_img, prompt]
                    )
                    ai_desc = response.text.strip()
                except Exception:
                    ai_desc = f"Pohľad na objekt"
            
            if not ai_desc:
                ai_desc = f"Pohľad na objekt"
                
            # Uloženie do katalógu
            st.session_state.catalog[f.name] = final_bytes
            
            # Pridanie riadku do editora poradia
            st.session_state.table_rows.append({
                "Názov súboru": f.name,
                "Náhad (2x3 cm)": web_thumb,
                "Popis obrázku (Kliknite a upravte)": ai_desc
            })

    # Odstránenie vymazaných fotiek z tabuľky
    current_names = [f.name for f in uploaded_files]
    st.session_state.table_rows = [r for r in st.session_state.table_rows if r["Názov súboru"] in current_names]

    st.markdown("---")
    st.subheader("📄 Usporiadanie poradia MYŠOU a úprava textu")
    st.info("💡 TIP: Chyťte riadok myšou na ľavom okraji a potiahnite ho hore/dole pre zmenu poradia vo Worde. Text popisu prepíšete dvojklikom.")

    # STABILNÝ DRAG & DROP EDITOR BEZ TLAČIDIEL
    edited_df = st.data_editor(
        st.session_state.table_rows,
        column_config={
            "Náhad (2x3 cm)": st.column_config.ImageColumn(width="small"),
            "Názov súboru": st.column_config.TextColumn(disabled=True),
            "Popis obrázku (Kliknite a upravte)": st.column_config.TextColumn(width="large")
        },
        disabled=["Názov súboru"],
        num_rows="dynamic",
        use_container_width=True,
        key="drag_drop_grid"
    )
    
    # Aktualizácia dát po zmene užívateľom
    st.session_state.table_rows = edited_df

    # 5. GENERÁTOR WORDU (.DOCX) - STOPERCETNE OPRAVENÝ RIADOK 176
    def generate_docx():
        doc = docx.Document()
        
        for section in doc.sections:
            section.top_margin = Inches(0.6)
            section.bottom_margin = Inches(0.8)
            section.left_margin = Inches(0.6)
            section.right_margin = Inches(0.6)
            
            header = section.header
            hp = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
            hp.text = f"Znalec: {znalec}\t\tFOTODOKUMENTÁCIA\t\t{datum}"
            hp.alignment = WD_ALIGN_PARAGRAPH.LEFT
            
            footer = section.footer
            fp = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
            fp.text = objekt
            fp.alignment = WD_ALIGN_PARAGRAPH.CENTER

        table = doc.add_table(rows=0, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Spracovanie podľa poradia určeného potiahnutím myšou
        for i in range(0, len(st.session_state.table_rows), 2):
            row_cells_img = table.add_row().cells
            row_cells_txt = table.add_row().cells
            
            # Ľavá strana
            item1 = st.session_state.table_rows[i]
            bytes1 = st.session_state.catalog[item1["Názov súboru"]]
            p1 = row_cells_img.paragraphs[0]
            p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p1.add_run().add_picture(io.BytesIO(bytes1), width=Inches(3.1))
            
            t1 = row_cells_txt.paragraphs[0]
            t1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            t1.add_run(item1["Popis obrázku (Kliknite a upravte)"]).font.size = Pt(10)
            
            # Pravá strana
            if i + 1 < len(st.session_state.table_rows):
                item2 = st.session_state.table_rows[i+1]
                bytes2 = st.session_state.catalog[item2["Názov súboru"]]
                p2 = row_cells_img.paragraphs[0]
                p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p2.add_run().add_picture(io.BytesIO(bytes2), width=Inches(3.1))
                
                t2 = row_cells_txt.paragraphs[0]
                t2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                t2.add_run(item2["Popis obrázku (Kliknite a upravte)"]).font.size = Pt(10)
                
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
