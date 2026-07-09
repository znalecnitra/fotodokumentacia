import streamlit as st
import google.generativeai as genai
from PIL import Image
import docx
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import io
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfgen import canvas

# 1. NASTAVENIE STRÁNKY
st.set_page_config(page_title="Znalec - Fotodokumentácia", layout="wide")
st.title("📸 Automatická Fotodokumentácia pre posudky")

# 2. BEZPLATNÉ AI NASTAVENIE (Google Gemini)
# API kľúč si vygenerujete zadarmo na: https://aistudio.google.com/
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
uploaded_files = st.file_uploader("Vyberte alebo potiahnite fotky (aj viacero naraz):", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

# Inicializácia pamäte pre popisy fotiek
if 'photo_data' not in st.session_state:
    st.session_state.photo_data = {}

if uploaded_files:
    st.subheader("🔍 Kontrola a úprava popisov")
    
    # Paralelné spracovanie fotiek pomocou AI
    for f in uploaded_files:
        if f.name not in st.session_state.photo_data:
            img = Image.open(f)
            
            # Ak je vložený kľúč, spýtame sa AI na popis
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

    # Zobrazenie v mriežke 2 fotky vedľa seba (ako na výstupe)
    cols = st.columns(2)
    for idx, (f_name, data) in enumerate(st.session_state.photo_data.items()):
        current_col = cols[idx % 2]
        with current_col:
            st.image(data["image"], use_container_width=True)
            # Upraviteľný text pre používateľa
            new_desc = st.text_input(f"Popis pre {f_name}:", value=data["desc"], key=f"input_{f_name}")
            st.session_state.photo_data[f_name]["desc"] = new_desc
            st.markdown("---")

    # 5. GENERATOR WORDU (.DOCX)
    def generate_docx():
        doc = docx.Document()
        
        # Nastavenie okrajov na 1,5 cm pre viac miesta
        for section in doc.sections:
            section.top_margin = Inches(0.6)
            section.bottom_margin = Inches(0.8)
            section.left_margin = Inches(0.6)
            section.right_margin = Inches(0.6)
            
            # Nastavenie fixnej hlavičky a päty do Word vlastností
            header = section.header
            hp = header.paragraphs[0]
            hp.text = f"Znalec: {znalec}\t\tFOTODOKUMENTÁCIA\t\t{datum}"
            hp.alignment = WD_ALIGN_PARAGRAPH.LEFT
            
            footer = section.footer
            fp = footer.paragraphs[0]
            fp.text = objekt
            fp.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Hlavná tabuľka pre fotky (2 stĺpce)
        items = list(st.session_state.photo_data.values())
        table = doc.add_table(rows=0, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        for i in range(0, len(items), 2):
            row_cells_img = table.add_row().cells
            row_cells_txt = table.add_row().cells
            
            # Ľavá bunka
            img_stream1 = io.BytesIO(items[i]["bytes"])
            p1 = row_cells_img[0].paragraphs[0]
            p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p1.add_run().add_picture(img_stream1, width=Inches(3.2))
            
            t1 = row_cells_txt[0].paragraphs[0]
            t1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            t1.add_run(items[i]["desc"]).font.size = Pt(10)
            
            # Pravá bunka (ak existuje párna fotka)
            if i + 1 < len(items):
                img_stream2 = io.BytesIO(items[i+1]["bytes"])
                p2 = row_cells_img[1].paragraphs[0]
                p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p2.add_run().add_picture(img_stream2, width=Inches(3.2))
                
                t2 = row_cells_txt[1].paragraphs[0]
                t2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                t2.add_run(items[i+1]["desc"]).font.size = Pt(10)
                
        bio = io.BytesIO()
        doc.save(bio)
        return bio.getvalue()

    # 6. GENERÁTOR PDF (ReportLab s Canvas pre fixné záhlavie/zápätie)
    class NumberedCanvas(canvas.Canvas):
        def __init__(self, *args, **kwargs):
            canvas.Canvas.__init__(self, *args, **kwargs)
            self._saved_page_states = []

        def showPage(self):
            self._saved_page_states.append(dict(self.__dict__))
            self._startPage()

        def save(self):
            num_pages = len(self._saved_page_states)
            for state in self._saved_page_states:
                self.__dict__.update(state)
                self.draw_page_elements(num_pages)
                canvas.Canvas.showPage(self)
            canvas.Canvas.save(self)

        def draw_page_elements(self, page_count):
            self.setFont("Helvetica", 9)
            # Hlavička
            self.drawString(40, 750, f"Znalec: {znalec}")
            self.drawCentredString(300, 750, "FOTODOKUMENTÁCIA")
            self.drawRightString(560, 750, datum)
            self.line(40, 742, 560, 742)
            
            # Päta
            self.line(40, 50, 560, 50)
            self.drawCentredString(300, 38, objekt)

    def generate_pdf():
        pdf_buffer = io.BytesIO()
        doc = SimpleDocTemplate(pdf_buffer, pagesize=letter, leftMargin=40, rightMargin=40, topMargin=70, bottomMargin=70)
        story = []
        
        styles = getSampleStyleSheet()
        desc_style = ParagraphStyle('DescStyle', parent=styles['Normal'], fontName='Helvetica', fontSize=9, alignment=1)
        
        items = list(st.session_state.photo_data.values())
        table_data = []
        
        for i in range(0, len(items), 2):
            row_img = []
            row_txt = []
            
            # Ľavá strana
            img1 = Image.open(io.BytesIO(items[i]["bytes"]))
            img1.thumbnail((240, 180))
            img_w1, img_h1 = img1.size
            rl_img1 = RLImage(io.BytesIO(items[i]["bytes"]), width=img_w1, height=img_h1)
            row_img.append(rl_img1)
            row_txt.append(Paragraph(items[i]["desc"], desc_style))
            
            # Pravá strana
            if i + 1 < len(items):
                img2 = Image.open(io.BytesIO(items[i+1]["bytes"]))
                img2.thumbnail((240, 180))
                img_w2, img_h2 = img2.size
                rl_img2 = RLImage(io.BytesIO(items[i+1]["bytes"]), width=img_w2, height=img_h2)
                row_img.append(rl_img2)
                row_txt.append(Paragraph(items[i+1]["desc"], desc_style))
            else:
                row_img.append("")
                row_txt.append("")
                
            table_data.append(row_img)
            table_data.append(row_txt)
            # Pridáme malú medzeru medzi riadkami fotiek
            table_data.append(["", ""])
            
        t = Table(table_data, colWidths=[260, 260])
        t.setStyle(TableStyle([
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('BOTTOMPADDING', (0,0), (-1,-1), 10),
        ]))
        story.append(t)
        
        doc.build(story, canvasmaker=NumberedCanvas)
        return pdf_buffer.getvalue()

    # 7. TLAČIDLÁ NA STIAHNUTIE
    st.subheader("💾 Stiahnuť hotové dokumenty")
    c1, c2 = st.columns(2)
    with c1:
        docx_data = generate_docx()
        st.download_button(label="📥 Stiahnuť upraviteľný WORD (.docx)", data=docx_data, file_name="Fotodokumentacia.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    with c2:
        pdf_data = generate_pdf()
        st.download_button(label="📥 Stiahnuť hotové PDF (.pdf)", data=pdf_data, file_name="Fotodokumentacia.pdf", mime="application/pdf")
