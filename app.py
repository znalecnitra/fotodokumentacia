import streamlit as st
import google.generativeai as genai
from PIL import Image, ImageOps
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io
import os

# Pokročilá knižnica pre Drag & Drop rozhranie
from streamlit_elements import elements, mui, html

# 1. NASTAVENIE STRÁNKY
st.set_page_config(page_title="Znalec - Fotodokumentácia PRO", layout="wide")
st.title("📸 Profesionálna Fotodokumentácia pre posudky")

# Pomocná funkcia na pridanie čierneho tenkého okraja bunkám vo Worde
def set_cell_border(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)
    for edge in ('top', 'left', 'bottom', 'right'):
        edge_data = OxmlElement(f'w:{edge}')
        edge_data.set(qn('w:val'), 'single')
        edge_data.set(qn('w:sz'), '4')  # 4 = 0.5 bodu (tenká čiara)
        edge_data.set(qn('w:space'), '0')
        edge_data.set(qn('w:color'), '000000')  # Čierna farba
        tcBorders.append(edge_data)

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

# Pamäť pre uchovanie stavu aplikácie
if 'order_list' not in st.session_state:
    st.session_state.order_list = []
if 'photos_dict' not in st.session_state:
    st.session_state.photos_dict = {}

if uploaded_files:
    # Spracovanie nových nahraných súborov
    for f in uploaded_files:
        if f.name not in st.session_state.photos_dict:
            # 3. POŽIADAVKA: Automatické otočenie obrázka podľa EXIF metadát z mobilu
            raw_img = Image.open(f)
            processed_img = ImageOps.exif_transpose(raw_img)
            
            # Optimalizácia pre rýchlosť webu (zmenšenie rozlíšenia pre náhľady)
            preview_img = processed_img.copy()
            preview_img.thumbnail((300, 300))
            
            # Volanie AI pre automatický popis
            if gemini_key:
                try:
                    prompt = "Analyzuj tento obrázok zo znaleckej obhliadky nehnuteľnosti. Napíš stručný, profesionálny názov alebo popis toho, čo je na fotke v slovenčine (napr. 'Pohľad na Kúpeľňu', 'Pohľad na Kuchyňu', 'Rozvody v byte', 'Pohľad severo-východný na Bytový dom'). Odpovedz LEN týmto popisom, maximálne 5 slov, nič iné nepíš."
                    response = model.generate_content([prompt, processed_img])
                    ai_desc = response.text.strip()
                except Exception:
                    ai_desc = f"Pohľad na {f.name}"
            else:
                ai_desc = f"Pohľad na {f.name}"
            
            # Uloženie do pamäte
            st.session_state.photos_dict[f.name] = {
                "full_img": processed_img,
                "preview_img": preview_img,
                "desc": ai_desc,
                "width": processed_img.width,
                "height": processed_img.height
            }
            st.session_state.order_list.append(f.name)

    # 1. a 5. POŽIADAVKA: Kompaktná mriežka s malými náhľadmi a Drag & Drop prehadzovaním
    st.subheader("↔️ Usporiadanie poradia (Potiahnite myšou) a úprava textu")
    st.caption("Kliknutím a potiahnutím sivého obdĺžnika zmeníte poradie. V textovom poli môžete popis okamžite prepísať.")

    # Vykreslenie Drag & Drop rozhrania pomocou Material UI (MUI) komponentov
    with elements("drag_and_drop_grid"):
        # Vytvorenie flexibilného kontajnera pre mriežku náhľadov
        items_layout = []
        for i, f_name in enumerate(st.session_state.order_list):
            # Definovanie pozície každej karty v mriežke (4 karty na riadok)
            x_pos = (i % 4) * 3
            y_pos = (i // 4) * 4
            items_layout.append({"i": f_name, "x": x_pos, "y": y_pos, "w": 3, "h": 4})

        def handle_layout_change(new_layout):
            # Aktualizácia poradia podľa toho, ako používateľ presunul karty
            sorted_layout = sorted(new_layout, key=lambda k: (k['y'], k['x']))
            st.session_state.order_list = [item['i'] for item in sorted_layout]

        with mui.react_grid_layout(layout=items_layout, cols=12, rowHeight=100, onLayoutChange=handle_layout_change, isResizable=False):
            for f_name in st.session_state.order_list:
                p_data = st.session_state.photos_dict[f_name]
                
                # Prevedenie náhľadu na formát, ktorý vieme zobraziť vo webe bez spomaľovania
                buffered = io.BytesIO()
                p_data["preview_img"].save(buffered, format="JPEG")
                import base64
                img_str = base64.b64encode(buffered.getvalue()).decode()
                
                # Karta pre každú fotku
                with mui.Card(key=f_name, variant="outlined", style={"display": "flex", "flexDirection": "column", "padding": "8px", "backgroundColor": "#fafafa"}):
                    with mui.Box(style={"display": "flex", "justifyContent": "center", "alignItems": "center", "height": "140px", "overflow": "hidden", "cursor": "move"}):
                        html.img(src=f"data:image/jpeg;base64,{img_str}", style={"maxHeight": "100%", "maxWidth": "100%", "objectFit": "contain"})
                    
                    # Interaktívne textové pole pre zmenu popisu
                    def make_change_handler(name):
                        return lambda event: st.session_state.photos_dict[name].update({"desc": event.target.value})
                        
                    mui.TextField(
                        label=f_name[:15] + "...",
                        defaultValue=p_data["desc"],
                        variant="standard",
                        size="small",
                        fullWidth=True,
                        onChange=make_change_handler(f_name),
                        style={"marginTop": "8px"}
                    )

    # 5. GENERÁTOR WORDU (.DOCX) PODĽA VŠETKÝCH POŽIADAVIEK
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
            for run in hp.runs:
                run.font.name = 'Arial'
                run.font.size = Pt(10)
                run.font.color.rgb = RGBColor(0, 0, 0)
            
            # Štýl pre pätu vo Worde
            footer = section.footer
            fp = footer.paragraphs[0]
            fp.text = objekt
            fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in fp.runs:
                run.font.name = 'Arial'
                run.font.size = Pt(10)
                run.font.color.rgb = RGBColor(0, 0, 0)

        # 6. POŽIADAVKA: Vytvorenie pevnej mriežky tabuľky (2 stĺpce)
        table = doc.add_table(rows=0, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Získanie fotiek v usporiadanom poradí od používateľa
        ordered_items = [st.session_state.photos_dict[name] for name in st.session_state.order_list]
        
        for i in range(0, len(ordered_items), 2):
            row_cells_img = table.add_row().cells
            row_cells_txt = table.add_row().cells
            
            # 4. POŽIADAVKA: Výpočty pre dodržanie pevnej výšky / šírky mriežky
            # Ak je obrázok na šírku, nastavíme mu šírku 3.2 palca. Ak je otočený na výšku, prispôsobíme šírku na 2.4 palca, aby na výšku nepresiahol pevnú mriežku.
            
            # Spracovanie ľavej bunky
            cell_left_img = row_cells_img[0]
            set_cell_border(cell_left_img)
            p1 = cell_left_img.paragraphs[0]
            p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cell_left_img.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            
            img_stream1 = io.BytesIO()
            ordered_items[i]["full_img"].save(img_stream1, format="JPEG")
            img_stream1.seek(0)
            
            # Rozhodnutie o rozmere na základe orientácie fotky
            if ordered_items[i]["width"] > ordered_items[i]["height"]:
                p1.add_run().add_picture(img_stream1, width=Inches(3.2))
            else:
                p1.add_run().add_picture(img_stream1, width=Inches(2.4))
                
            cell_left_txt = row_cells_txt[0]
            set_cell_border(cell_left_txt)
            t1 = cell_left_txt.paragraphs[0]
            t1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cell_left_txt.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
