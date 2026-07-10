import streamlit as st
import google.generativeai as genai
from PIL import Image, ImageOps
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
    try:
        genai.configure(api_key=gemini_key)
        model = genai.GenerativeModel('gemini-2.5-flash')
    except Exception as e:
        st.sidebar.error(f"Chyba konfigurácie AI: {e}")
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

# Inicializácia relácie a poradia v pamäti
if 'photo_order' not in st.session_state:
    st.session_state.photo_order = []
if 'photo_catalog' not in st.session_state:
    st.session_state.photo_catalog = {}

if uploaded_files:
    # Spracovanie nových nahraných súborov
    for f in uploaded_files:
        if f.name not in st.session_state.photo_catalog:
            # Načítanie obrázka a AUTOMATICKÉ OTOČENIE podľa EXIF dát z mobilu
            raw_img = Image.open(f)
            corrected_img = ImageOps.exif_transpose(raw_img)
            
            # Príprava bajtov pre export
            img_byte_arr = io.BytesIO()
            corrected_img.save(img_byte_arr, format=raw_img.format if raw_img.format else "JPEG")
            final_bytes = img_byte_arr.getvalue()
            
            # Volanie AI pre skutočný popis (nie iba názov súboru)
            ai_desc = ""
            if gemini_key:
                try:
                    prompt = "Hovoríš po slovensky. Analyzuj tento obrázok zo znaleckej obhliadky nehnuteľnosti. Napíš stručný, jednovetný, profesionálny názov toho, čo presne vidíš (napr. 'Pohľad na bytový dom z exteriéru', 'Vstupný portál a schodisko do bytového domu', 'Pohľad na interiér obývacej izby'). Odpovedz LEN týmto popisom."
                    response = model.generate_content([prompt, corrected_img])
                    ai_desc = response.text.strip()
                except Exception:
                    ai_desc = f"Pohľad na objekt ({f.name})"
            
            if not ai_desc:
                ai_desc = f"Pohľad na objekt ({f.name})"
                
            # Uloženie do katalógu
            st.session_state.photo_catalog[f.name] = {
                "image": corrected_img,
                "desc": ai_desc,
                "bytes": final_bytes
            }
            if f.name not in st.session_state.photo_order:
                st.session_state.photo_order.append(f.name)

    # Vyčistenie starých súborov, ktoré už užívateľ odobral z uploaderu
    current_names = [f.name for f in uploaded_files]
    st.session_state.photo_order = [name for name in st.session_state.photo_order if name in current_names]

    st.subheader("🔍 Usporiadanie poradia a úprava textu")
    
    # Vykreslenie mriežky s kontrolou poradia (tlačidlá hore/dole)
    for idx, f_name in enumerate(st.session_state.photo_order):
        data = st.session_state.photo_catalog[f_name]
        
        # Vytvorenie riadku: Fotka | Tlačidlá poradia | Textové pole
        r_col1, r_col2, r_col3 = st.columns([2, 1, 5])
        
        with r_col1:
            # Kompaktný náhľad na webe (šírka len 300px)
            st.image(data["image"], width=300)
            
        with r_col2:
            st.write("") # Odsadenie
            # Tlačidlá na zmenu poradia v zozname
            if idx > 0:
                if st.button("🔼 Hore", key=f"up_{f_name}_{idx}"):
                    st.session_state.photo_order[idx], st.session_state.photo_order[idx-1] = st.session_state.photo_order[idx-1], st.session_state.photo_order[idx]
                    st.rerun()
            if idx < len(st.session_state.photo_order) - 1:
                if st.button("🔽 Dole", key=f"down_{f_name}_{idx}"):
                    st.session_state.photo_order[idx], st.session_state.photo_order[idx+1] = st.session_state.photo_order[idx+1], st.session_state.photo_order[idx]
                    st.rerun()
                    
        with r_col3:
            # Editovateľné textové pole pre úpravu popisu
            new_desc = st.text_area(f"Popis dokumentu ({f_name}):", value=data["desc"], key=f"txt_{f_name}_{idx}")
            st.session_state.photo_catalog[f_name]["desc"] = new_desc
            
        st.markdown("---")

    # 5. GENERÁTOR WORDU (.DOCX) - Kompletne opravený riadok 111
    def generate_docx():
        doc = docx.Document()
        
        # Nastavenie okrajov strany (cca 1.5 cm)
        for section in doc.sections:
            section.top_margin = Inches(0.6)
            section.bottom_margin = Inches(0.8)
            section.left_margin = Inches(0.6)
            section.right_margin = Inches(0.6)
            
            # Fixná hlavička
            header = section.header
            hp = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
            hp.text = f"Znalec: {znalec}\t\tFOTODOKUMENTÁCIA\t\t{datum}"
            hp.alignment = WD_ALIGN_PARAGRAPH.LEFT
            
            # Fixná päta
            footer = section.footer
            fp = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
            fp.text = objekt
            fp.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Mriežka tabuľky vo Worde (2 stĺpce)
        table = doc.add_table(rows=0, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Čítanie dát presne podľa usporiadaného poradia od užívateľa
        order = st.session_state.photo_order
        
        for i in range(0, len(order), 2):
            row_cells_img = table.add_row().cells
            row_cells_txt = table.add_row().cells
            
            # Ľavá strana tabuľky
            name1 = order[i]
            item1 = st.session_state.photo_catalog[name1]
            img_stream1 = io.BytesIO(item1["bytes"])
            p1 = row_cells_img.paragraphs[0]
            p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p1.add_run().add_picture(img_stream1, width=Inches(3.2))
            
            t1 = row_cells_txt.paragraphs[0]
            t1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            t1.add_run(item1["desc"]).font.size = Pt(10)
            
            # Pravá strana tabuľky (ak existuje druhá fotka do páru)
            if i + 1 < len(order):
                name2 = order[i+1]
                item2 = st.session_state.photo_catalog[name2]
                img_stream2 = io.BytesIO(item2["bytes"])
                p2 = row_cells_img.paragraphs[0]
                p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p2.add_run().add_picture(img_stream2, width=Inches(3.2))
                
                t2 = row_cells_txt.paragraphs[0]
                t2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                t2.add_run(item2["desc"]).font.size = Pt(10)
                
        bio = io.BytesIO()
        doc.save(bio)
        return bio.getvalue()

    # TLAČIDLO NA STIAHNUTIE WORDU
    st.subheader("💾 Stiahnuť fotodokumentáciu")
    try:
        docx_data = generate_docx()
        st.download_button(
            label="📥 Stiahnuť hotový WORD (.docx) s Hlavičkou a Pätou",
            data=docx_data,
            file_name="Fotodokumentacia.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as error:
        st.error(f"Chyba pri príprave dokumentu na stiahnutie: {error}")
