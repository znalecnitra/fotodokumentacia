import streamlit as st
import google.generativeai as genai
from PIL import Image
import docx
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import io

# 1. NASTAVENIE STRÁNKY
st.set_page_config(page_title="Znalec - Fotodokumentácia", layout="wide")
st.title("📸 Profesionálna Fotodokumentácia pre posudky")

# 2. BEZPLATNÉ AI NASTAVENIE (Google Gemini)
gemini_key = st.sidebar.text_input("Vložte bezplatný Gemini API kľúč:", type="password")

if gemini_key:
    genai.configure(api_key=gemini_key)
    model = genai.GenerativeModel('gemini-2.5-flash')
else:
    st.warning("🔑 Pre funkciu automatických popisov vložte do bočného panelu váš bezplatný Gemini API kľúč.")

# 3. VSTUPNÉ ÚDAJE PRE HLAVIČKU A PÄTU
st.subheader("📝 Základné údaje posudku")
col1, col2, col3 = st.columns(3)
with col1:
    znalec = st.text_input("Meno znalca:", value="Ing. Jozef HRČKA")
with col2:
    datum = st.text_input("Dátum obhliadky:", value="08.07.2026")
with col3:
    objekt = st.text_area("Objekt / Byt (Päta):", value="Byt č.7 ,3.p., o.č.6, Javorová 643/6, k.ú. Chrenová, mesto Nitra, okres Nitra")

# 4. HROMADNÉ NAHRÁVANIE FOTIEK
st.subheader("🖼️ Nahrať fotky pre projekt")
uploaded_files = st.file_uploader("Vyberte alebo potiahnite fotky naraz:", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

# Inicializácia pamäte pre popisy fotiek
if 'photo_data' not in st.session_state:
    st.session_state.photo_data = {}

if uploaded_files:
    st.subheader("🔍 Generovanie popisov a úprava textu")
    
    # Zoradenie súborov abecedne/podľa času (názvu súboru)
    sorted_files = sorted(uploaded_files, key=lambda x: x.name)
    
    # Načítanie dát
    for f in sorted_files:
        if f.name not in st.session_state.photo_data:
            img = Image.open(f)
            
            if gemini_key:
                try:
                    prompt = "Analyzuj tento obrázok zo znaleckej obhliadky nehnuteľnosti. Napíš stručný, profesionálny názov alebo popis toho, čo je na fotke v slovenčine (napr. 'Pohľad na Kúpeľňu', 'Pohľad na Kuchyňu', 'Rozvody v byte', 'Pohľad severo-východný na Bytový dom'). Odpovedz LEN týmto popisom, nič iné nepíš."
                    response = model.generate_content([prompt, img])
                    ai_desc = response.text.strip()
                except Exception as e:
                    ai_desc = f"Pohľad na {f.name}"
            else:
                ai_desc = f"Pohľad na {f.name}"
                
            st.session_state.photo_data[f.name] = {"image": img, "desc": ai_desc, "bytes": f.getvalue()}

    # Zobrazenie v stabilnej mriežke 2 stĺpce
    cols = st.columns(2)
    for idx, f_name in enumerate(list(st.session_state.photo_data.keys())):
        data = st.session_state.photo_data[f_name]
        current_col = cols[idx % 2]
        with current_col:
            st.image(data["image"], use_container_width=True)
            # Každý popis je okamžite upraviteľný
            new_desc = st.text_input(f"Popis ({f_name}):", value=data["desc"], key=f"input_{f_name}")
            st.session_state.photo_data[f_name]["desc"] = new_desc
            st.markdown("---")

    # 5. GENERÁTOR WORDU (.DOCX)
    def generate_docx():
        doc = docx.Document()
        
        # Nastavenie okrajov strany (cca 1.5 cm)
        for section in doc.sections:
            section.top_margin = Inches(0.6)
            section.bottom_margin = Inches(0.8)
            section.left_margin = Inches(0.6)
            section.right_margin = Inches(0.6)
            
            # Štýl pre hlavičku vo Worde
            header = section.header
            hp = header.paragraphs[0]
            hp.text = f"Znalec: {znalec}\t\tFOTODOKUMENTÁCIA\t\t{datum}"
            hp.alignment = WD_ALIGN_PARAGRAPH.LEFT
            
            # Štýl pre pätu vo Worde
            footer = section.footer
            fp = footer.paragraphs[0]
            fp.text = objekt
            fp.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Vytvorenie pevnej mriežky tabuľky (2 stĺpce)
        table = doc.add_table(rows=0, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        items = list(st.session_state.photo_data.values())
        
        for i in range(0, len(items), 2):
            row_cells_img = table.add_row().cells
            row_cells_txt = table.add_row().cells
            
            # Ľavá bunka
            img_stream1 = io.BytesIO(items[i]["bytes"])
            p1 = row_cells_img.paragraphs[0]
            p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p1.add_run().add_picture(img_stream1, width=Inches(3.2))
            
            t1 = row_cells_txt.paragraphs[0]
            t1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            t1.add_run(items[i]["desc"]).font.size = Pt(10)
            
            # Pravá bunka
            if i + 1 < len(items):
                img_stream2 = io.BytesIO(items[i+1]["bytes"])
                p2 = row_cells_img.paragraphs[0]
                p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p2.add_run().add_picture(img_stream2, width=Inches(3.2))
                
                t2 = row_cells_txt.paragraphs[0]
                t2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                t2.add_run(items[i+1]["desc"]).font.size = Pt(10)
                
        bio = io.BytesIO()
        doc.save(bio)
        return bio.getvalue()

    # TLAČIDLO NA STIAHNUTIE WORDU
    st.subheader("💾 Stiahnuť fotodokumentáciu")
    docx_data = generate_docx()
    st.download_button(
        label="📥 Stiahnuť hotový WORD (.docx) s Hlavičkou a Pätou",
        data=docx_data,
        file_name="Fotodokumentacia.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
