import streamlit as st
from docx import Document
from docx.shared import Pt, Inches, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from datetime import datetime
import io
import re

# ==================== KONFIGURACJA STRONY ====================
st.set_page_config(
    page_title="Kontrakt D/s | BDSM Contract Builder",
    page_icon="🔗",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==================== CUSTOM CSS – ciemny elegancki motyw ====================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@500;700&family=Inter:wght@400;500;600&display=swap');
    
    .stApp {
        background: linear-gradient(160deg, #0f0f0f 0%, #1a0a0a 50%, #0d0d0d 100%);
        color: #e8e0e0;
    }
    
    h1, h2, h3 {
        font-family: 'Playfair Display', serif !important;
        color: #f5e6e6 !important;
    }
    
    .main-title {
        font-family: 'Playfair Display', serif;
        font-size: 2.6rem;
        font-weight: 700;
        background: linear-gradient(90deg, #c9a0a0, #e8c4c4, #c9a0a0);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        text-align: center;
        margin-bottom: 0.3rem;
        letter-spacing: 1px;
    }
    
    .subtitle {
        text-align: center;
        color: #a89090;
        font-size: 1.05rem;
        margin-bottom: 2rem;
        font-weight: 400;
    }
    
    .section-card {
        background: rgba(30, 15, 15, 0.7);
        border: 1px solid #3d2020;
        border-radius: 12px;
        padding: 1.5rem 1.8rem;
        margin-bottom: 1.4rem;
        box-shadow: 0 4px 20px rgba(0,0,0,0.4);
    }
    
    .stTextInput > div > div > input,
    .stTextArea > div > div > textarea,
    .stSelectbox > div > div {
        background-color: #1f1212 !important;
        color: #f0e8e8 !important;
        border: 1px solid #4a2c2c !important;
        border-radius: 8px !important;
    }
    
    .stTextArea textarea {
        font-size: 0.95rem !important;
    }
    
    .stButton > button {
        background: linear-gradient(135deg, #6b2d2d, #8b3a3a) !important;
        color: #fff !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 0.6rem 1.8rem !important;
        font-weight: 600 !important;
        letter-spacing: 0.5px !important;
        transition: all 0.25s ease !important;
    }
    
    .stButton > button:hover {
        background: linear-gradient(135deg, #8b3a3a, #a84a4a) !important;
        box-shadow: 0 0 18px rgba(160, 60, 60, 0.45) !important;
        transform: translateY(-1px);
    }
    
    .stDownloadButton > button {
        background: linear-gradient(135deg, #2d4a3e, #3d6b55) !important;
        color: #fff !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
    }
    
    .stDownloadButton > button:hover {
        background: linear-gradient(135deg, #3d6b55, #4d8b6a) !important;
        box-shadow: 0 0 18px rgba(60, 140, 100, 0.4) !important;
    }
    
    .info-box {
        background: rgba(60, 30, 30, 0.5);
        border-left: 4px solid #a84a4a;
        padding: 1rem 1.2rem;
        border-radius: 0 8px 8px 0;
        margin: 1rem 0;
        font-size: 0.92rem;
        color: #d8c8c8;
    }
    
    .success-box {
        background: rgba(30, 50, 40, 0.5);
        border-left: 4px solid #4a8b6a;
        padding: 1rem 1.2rem;
        border-radius: 0 8px 8px 0;
        margin: 1rem 0;
        color: #c8e0d0;
    }
    
    .role-badge {
        display: inline-block;
        padding: 0.25rem 0.8rem;
        border-radius: 20px;
        font-size: 0.85rem;
        font-weight: 600;
        letter-spacing: 0.5px;
    }
    
    .badge-dom {
        background: #4a2020;
        color: #e8b0b0;
        border: 1px solid #6b3030;
    }
    
    .badge-sub {
        background: #202a4a;
        color: #b0c0e8;
        border: 1px solid #30406b;
    }
    
    div[data-testid="stSidebar"] {
        background: #120a0a;
        border-right: 1px solid #2a1515;
    }
    
    .stRadio > div {
        gap: 1rem;
    }
    
    label {
        color: #d0c0c0 !important;
        font-weight: 500 !important;
    }
    
    .stMarkdown p {
        color: #d8d0d0;
    }
    
    hr {
        border-color: #3a2020 !important;
    }
</style>
""", unsafe_allow_html=True)

# ==================== FUNKCJE DOCX ====================

def set_run_font(run, name="Georgia", size=11, bold=False, color=None, italic=False):
    run.font.name = name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), name)
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color:
        run.font.color.rgb = RGBColor(*color)


def add_horizontal_line(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '12')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), '8B3A3A')
    pBdr.append(bottom)
    pPr.append(pBdr)


def create_contract_docx(data: dict, filled_by: str) -> bytes:
    """
    Tworzy profesjonalny dokument Word.
    filled_by: 'dom' | 'sub' | 'both'
    """
    doc = Document()

    # Marginesy
    for section in doc.sections:
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(2.3)
        section.right_margin = Cm(2.3)

    # Style
    styles = doc.styles
    style_normal = styles['Normal']
    style_normal.font.name = 'Georgia'
    style_normal.font.size = Pt(11)
    style_normal._element.rPr.rFonts.set(qn('w:eastAsia'), 'Georgia')

    # ===== TYTUŁ =====
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("KONTRAKT")
    set_run_font(run, size=22, bold=True, color=(90, 30, 30))
    title.paragraph_format.space_after = Pt(2)

    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subtitle.add_run("Dominujący / Dominująca  —  Uległy / Uległa")
    set_run_font(run, size=14, italic=True, color=(120, 60, 60))
    subtitle.paragraph_format.space_after = Pt(4)

    date_p = doc.add_paragraph()
    date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = date_p.add_run(f"Data utworzenia: {datetime.now().strftime('%d.%m.%Y')}")
    set_run_font(run, size=9, color=(100, 100, 100))
    date_p.paragraph_format.space_after = Pt(12)

    add_horizontal_line(doc)

    # ===== PREAMBUŁA =====
    h = doc.add_paragraph()
    run = h.add_run("1. Preambuła i zasady ogólne")
    set_run_font(run, size=13, bold=True, color=(90, 30, 30))
    h.paragraph_format.space_before = Pt(10)
    h.paragraph_format.space_after = Pt(6)

    preamble = (
        "Niniejszy dokument jest dobrowolnym, świadomym i wzajemnym porozumieniem pomiędzy stronami. "
        "Nie stanowi umowy prawnie wiążącej i nie zastępuje obowiązującego prawa. "
        "Każda ze stron ma prawo w każdej chwili, bez podawania przyczyny, wypowiedzieć ten kontrakt "
        "lub zawiesić jego obowiązywanie. Zasady SSC (Safe, Sane, Consensual) oraz RACK (Risk-Aware Consensual Kink) "
        "stanowią fundament tej relacji. Bezpieczeństwo fizyczne i psychiczne obu stron ma absolutny priorytet."
    )
    p = doc.add_paragraph()
    run = p.add_run(preamble)
    set_run_font(run, size=10)
    p.paragraph_format.space_after = Pt(10)
    p.paragraph_format.line_spacing = 1.15

    # ===== STRONY =====
    h = doc.add_paragraph()
    run = h.add_run("2. Strony kontraktu")
    set_run_font(run, size=13, bold=True, color=(90, 30, 30))
    h.paragraph_format.space_before = Pt(8)
    h.paragraph_format.space_after = Pt(6)

    p = doc.add_paragraph()
    run = p.add_run("Osoba Dominująca: ")
    set_run_font(run, size=11, bold=True)
    run = p.add_run(data.get("dom_name") or "……………………………………")
    set_run_font(run, size=11)

    p = doc.add_paragraph()
    run = p.add_run("Osoba Uległa: ")
    set_run_font(run, size=11, bold=True)
    run = p.add_run(data.get("sub_name") or "……………………………………")
    set_run_font(run, size=11)

    p = doc.add_paragraph()
    run = p.add_run("Zakres dynamiki: ")
    set_run_font(run, size=11, bold=True)
    run = p.add_run(data.get("scope") or "nie określono")
    set_run_font(run, size=11)
    p.paragraph_format.space_after = Pt(8)

    add_horizontal_line(doc)

    # ===== FUNKCJA POMOCNICZA DO SEKCJI =====
    def add_section(title_text: str, content: str, is_placeholder: bool = False):
        h = doc.add_paragraph()
        run = h.add_run(title_text)
        set_run_font(run, size=12, bold=True, color=(90, 30, 30))
        h.paragraph_format.space_before = Pt(10)
        h.paragraph_format.space_after = Pt(4)

        p = doc.add_paragraph()
        if content and content.strip():
            run = p.add_run(content.strip())
            set_run_font(run, size=10)
        else:
            run = p.add_run("[DO UZUPEŁNIENIA PRZEZ PARTNERA]")
            set_run_font(run, size=10, italic=True, color=(150, 80, 80))
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.line_spacing = 1.15

    # ===== SEKCJE DLA DOMINUJĄCEGO =====
    h = doc.add_paragraph()
    run = h.add_run("3. Część Osoby Dominującej")
    set_run_font(run, size=13, bold=True, color=(90, 30, 30))
    h.paragraph_format.space_before = Pt(6)
    h.paragraph_format.space_after = Pt(4)

    add_section("3.1  Czego potrzebuję od osoby uległej", data.get("dom_needs", ""))
    add_section("3.2  Czego oczekuję", data.get("dom_expectations", ""))
    add_section("3.3  Hard Limits (kategorycznie nie)", data.get("dom_hard_limits", ""))
    add_section("3.4  Soft Limits (do omówienia / z warunkami)", data.get("dom_soft_limits", ""))
    add_section("3.5  Preferencje dotyczące aftercare i komunikacji", data.get("dom_aftercare", ""))

    add_horizontal_line(doc)

    # ===== SEKCJE DLA ULEGŁEGO =====
    h = doc.add_paragraph()
    run = h.add_run("4. Część Osoby Uległej")
    set_run_font(run, size=13, bold=True, color=(90, 30, 30))
    h.paragraph_format.space_before = Pt(6)
    h.paragraph_format.space_after = Pt(4)

    add_section("4.1  Czego potrzebuję od osoby dominującej", data.get("sub_needs", ""))
    add_section("4.2  Czego oczekuję", data.get("sub_expectations", ""))
    add_section("4.3  Hard Limits (kategorycznie nie)", data.get("sub_hard_limits", ""))
    add_section("4.4  Soft Limits (do omówienia / z warunkami)", data.get("sub_soft_limits", ""))
    add_section("4.5  Preferencje dotyczące aftercare i komunikacji", data.get("sub_aftercare", ""))

    add_horizontal_line(doc)

    # ===== SAFEWORDS =====
    h = doc.add_paragraph()
    run = h.add_run("5. Słowa bezpieczeństwa (Safewords)")
    set_run_font(run, size=13, bold=True, color=(90, 30, 30))
    h.paragraph_format.space_before = Pt(8)
    h.paragraph_format.space_after = Pt(4)

    p = doc.add_paragraph()
    run = p.add_run("System słów bezpieczeństwa: ")
    set_run_font(run, size=10, bold=True)
    run = p.add_run(data.get("safewords") or "RED – natychmiastowy stop  |  YELLOW – zwolnij / sprawdź  |  GREEN – wszystko w porządku")
    set_run_font(run, size=10)
    p.paragraph_format.space_after = Pt(4)

    p = doc.add_paragraph()
    run = p.add_run(
        "Użycie słowa bezpieczeństwa (szczególnie RED) nigdy nie jest powodem do kary ani dyskusji w momencie użycia. "
        "Scena zostaje natychmiast przerwana lub dostosowana."
    )
    set_run_font(run, size=9, italic=True, color=(80, 80, 80))
    p.paragraph_format.space_after = Pt(8)

    # ===== CZAS TRWANIA =====
    h = doc.add_paragraph()
    run = h.add_run("6. Czas trwania i przeglądy")
    set_run_font(run, size=13, bold=True, color=(90, 30, 30))
    h.paragraph_format.space_before = Pt(6)
    h.paragraph_format.space_after = Pt(4)

    p = doc.add_paragraph()
    run = p.add_run(data.get("duration") or "Do ustalenia przez obie strony. Rekomendowany regularny przegląd (np. co 1–3 miesiące).")
    set_run_font(run, size=10)
    p.paragraph_format.space_after = Pt(10)

    add_horizontal_line(doc)

    # ===== PODPISY =====
    h = doc.add_paragraph()
    run = h.add_run("7. Potwierdzenie")
    set_run_font(run, size=13, bold=True, color=(90, 30, 30))
    h.paragraph_format.space_before = Pt(8)
    h.paragraph_format.space_after = Pt(8)

    p = doc.add_paragraph()
    run = p.add_run(
        "Podpisując (lub potwierdzając) ten dokument, obie strony oświadczają, że zapoznały się z jego treścią, "
        "rozumieją go i dobrowolnie wyrażają zgodę na zawarte w nim ustalenia. "
        "Zgoda może zostać wycofana w dowolnym momencie."
    )
    set_run_font(run, size=9, italic=True)
    p.paragraph_format.space_after = Pt(16)

    # Tabelka na podpisy
    table = doc.add_table(rows=2, cols=2)
    table.autofit = True

    cell = table.rows[0].cells[0]
    p = cell.paragraphs[0]
    run = p.add_run("Osoba Dominująca")
    set_run_font(run, size=10, bold=True)
    p = cell.add_paragraph()
    run = p.add_run(data.get("dom_name") or "________________________")
    set_run_font(run, size=10)
    p = cell.add_paragraph()
    run = p.add_run("Data: _______________")
    set_run_font(run, size=9, color=(100, 100, 100))

    cell = table.rows[0].cells[1]
    p = cell.paragraphs[0]
    run = p.add_run("Osoba Uległa")
    set_run_font(run, size=10, bold=True)
    p = cell.add_paragraph()
    run = p.add_run(data.get("sub_name") or "________________________")
    set_run_font(run, size=10)
    p = cell.add_paragraph()
    run = p.add_run("Data: _______________")
    set_run_font(run, size=9, color=(100, 100, 100))

    # Stopka
    doc.add_paragraph()
    footer = doc.add_paragraph()
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = footer.add_run("Dokument prywatny • Tylko do użytku stron • Nie stanowi umowy cywilnoprawnej")
    set_run_font(run, size=8, color=(130, 130, 130), italic=True)

    # Zapis do bytes
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def extract_text_from_docx(file) -> str:
    """Proste wyciągnięcie tekstu z wgranego pliku (do podglądu)."""
    doc = Document(file)
    full = []
    for para in doc.paragraphs:
        if para.text.strip():
            full.append(para.text.strip())
    return "\n".join(full)


# ==================== UI ====================

st.markdown('<div class="main-title">Kontrakt D/s</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Generator kontraktu Dominujący/a — Uległy/a<br>Narzędzie komunikacji i bezpieczeństwa</div>', unsafe_allow_html=True)

# Sidebar
with st.sidebar:
    st.markdown("### Tryb pracy")
    mode = st.radio(
        "Wybierz:",
        ["Tworzę nowy kontrakt", "Uzupełniam otrzymany plik"],
        label_visibility="collapsed"
    )
    
    st.markdown("---")
    st.markdown("### Wskazówki")
    st.markdown("""
    <div class="info-box">
    <b>Hard Limits</b> – rzeczy, na które się kategorycznie nie zgadzasz. Nigdy nie powinny być negocjowane pod presją.<br><br>
    <b>Soft Limits</b> – obszary, które wymagają dodatkowej rozmowy, warunków lub wolniejszego wprowadzania.<br><br>
    Bądź konkretny/a. Im precyzyjniej opiszesz potrzeby i granice, tym bezpieczniejsza i przyjemniejsza będzie dynamika.
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("---")
    st.caption("Aplikacja działa lokalnie. Żadne dane nie są wysyłane na serwer.")

# ==================== TRYB 1: NOWY KONTRAKT ====================
if mode == "Tworzę nowy kontrakt":
    st.markdown("### Krok 1 — Kto wypełnia teraz?")
    filler = st.radio(
        "Wypełniam jako:",
        ["Osoba Dominująca", "Osoba Uległa"],
        horizontal=True
    )
    is_dom = filler == "Osoba Dominująca"
    
    st.markdown("---")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("#### Dane podstawowe")
        dom_name = st.text_input("Imię / pseudonim Osoby Dominującej", placeholder="np. Alex, Mistress K.")
        sub_name = st.text_input("Imię / pseudonim Osoby Uległej", placeholder="np. Sam, pet")
        scope = st.selectbox(
            "Zakres dynamiki",
            [
                "Tylko sceny (scene-only)",
                "Częściowy (part-time / weekendowy)",
                "24/7 z wyjątkami",
                "Total Power Exchange (TPE)",
                "Inny / do ustalenia"
            ]
        )
        duration = st.text_input(
            "Czas trwania / przeglądy",
            placeholder="np. 3 miesiące z przeglądem co miesiąc"
        )
        safewords = st.text_input(
            "Safewords",
            value="RED – stop  |  YELLOW – zwolnij  |  GREEN – ok"
        )
    
    with col2:
        st.markdown("#### Twoja część")
        if is_dom:
            st.markdown('<span class="role-badge badge-dom">DOMINUJĄCY / DOMINUJĄCA</span>', unsafe_allow_html=True)
            needs = st.text_area(
                "Czego potrzebuję od osoby uległej",
                height=110,
                placeholder="Np. szczerości, regularnej komunikacji, konkretnego rodzaju służby, zaufania..."
            )
            expectations = st.text_area(
                "Czego oczekuję",
                height=110,
                placeholder="Np. przestrzegania ustalonych reguł, inicjatywy w określonych obszarach, konkretnego podejścia do protokołu..."
            )
            hard = st.text_area(
                "Hard Limits (kategorycznie nie)",
                height=100,
                placeholder="Np. blood play, publiczne ujawnianie tożsamości, konkretne praktyki..."
            )
            soft = st.text_area(
                "Soft Limits (do omówienia)",
                height=80,
                placeholder="Opcjonalnie..."
            )
            aftercare = st.text_area(
                "Aftercare i komunikacja – moje preferencje",
                height=80,
                placeholder="Np. potrzebuję ciszy po scenie / rozmowy / kontaktu fizycznego..."
            )
        else:
            st.markdown('<span class="role-badge badge-sub">ULEGŁY / ULEGŁA</span>', unsafe_allow_html=True)
            needs = st.text_area(
                "Czego potrzebuję od osoby dominującej",
                height=110,
                placeholder="Np. jasnych oczekiwań, poczucia bezpieczeństwa, konsekwencji, opieki po scenie..."
            )
            expectations = st.text_area(
                "Czego oczekuję",
                height=110,
                placeholder="Np. respektowania limitów, regularnego feedbacku, konkretnego stylu prowadzenia..."
            )
            hard = st.text_area(
                "Hard Limits (kategorycznie nie)",
                height=100,
                placeholder="Np. konkretne praktyki, słowa, sytuacje, które są absolutnie wykluczone..."
            )
            soft = st.text_area(
                "Soft Limits (do omówienia)",
                height=80,
                placeholder="Opcjonalnie..."
            )
            aftercare = st.text_area(
                "Aftercare i komunikacja – moje preferencje",
                height=80,
                placeholder="Np. potrzebuję przytulenia, rozmowy, wody, czasu w ciszy..."
            )
    
    st.markdown("---")
    
    if st.button("Generuj plik Word (.docx)", use_container_width=True):
        data = {
            "dom_name": dom_name,
            "sub_name": sub_name,
            "scope": scope,
            "duration": duration,
            "safewords": safewords,
        }
        
        if is_dom:
            data.update({
                "dom_needs": needs,
                "dom_expectations": expectations,
                "dom_hard_limits": hard,
                "dom_soft_limits": soft,
                "dom_aftercare": aftercare,
                # druga strona pusta
                "sub_needs": "",
                "sub_expectations": "",
                "sub_hard_limits": "",
                "sub_soft_limits": "",
                "sub_aftercare": "",
            })
            filled_by = "dom"
        else:
            data.update({
                "sub_needs": needs,
                "sub_expectations": expectations,
                "sub_hard_limits": hard,
                "sub_soft_limits": soft,
                "sub_aftercare": aftercare,
                "dom_needs": "",
                "dom_expectations": "",
                "dom_hard_limits": "",
                "dom_soft_limits": "",
                "dom_aftercare": "",
            })
            filled_by = "sub"
        
        docx_bytes = create_contract_docx(data, filled_by)
        
        filename = f"kontrakt_Ds_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
        
        st.markdown("""
        <div class="success-box">
        Plik wygenerowany. Pobierz go i prześlij partnerowi/partnerce. 
        On/ona wgra plik w trybie „Uzupełniam otrzymany plik” i wypełni swoją część.
        </div>
        """, unsafe_allow_html=True)
        
        st.download_button(
            label="Pobierz kontrakt (.docx)",
            data=docx_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

# ==================== TRYB 2: UZUPEŁNIANIE ====================
else:
    st.markdown("### Wgraj otrzymany plik kontraktu")
    
    uploaded = st.file_uploader(
        "Wybierz plik .docx wygenerowany wcześniej",
        type=["docx"],
        help="Wgraj plik, który dostałeś/aś od partnera/partnerki"
    )
    
    if uploaded:
        st.markdown("""
        <div class="info-box">
        Plik wczytany. Poniżej możesz zobaczyć jego treść (podgląd) i uzupełnić brakującą część.
        Aplikacja nie edytuje automatycznie istniejącego pliku – generuje nową, kompletną wersję na podstawie tego, co wpiszesz.
        </div>
        """, unsafe_allow_html=True)
        
        with st.expander("Podgląd treści wgranego pliku"):
            try:
                text = extract_text_from_docx(uploaded)
                st.text(text[:4000] + ("..." if len(text) > 4000 else ""))
            except Exception as e:
                st.error(f"Nie udało się odczytać pliku: {e}")
        
        st.markdown("---")
        st.markdown("### Uzupełnij swoją część")
        
        filler2 = st.radio(
            "Wypełniam jako:",
            ["Osoba Dominująca", "Osoba Uległa"],
            horizontal=True,
            key="filler2"
        )
        is_dom2 = filler2 == "Osoba Dominująca"
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("#### Dane podstawowe (uzupełnij / popraw)")
            dom_name2 = st.text_input("Imię / pseudonim Osoby Dominującej", key="dn2")
            sub_name2 = st.text_input("Imię / pseudonim Osoby Uległej", key="sn2")
            scope2 = st.selectbox(
                "Zakres dynamiki",
                [
                    "Tylko sceny (scene-only)",
                    "Częściowy (part-time / weekendowy)",
                    "24/7 z wyjątkami",
                    "Total Power Exchange (TPE)",
                    "Inny / do ustalenia"
                ],
                key="sc2"
            )
            duration2 = st.text_input("Czas trwania / przeglądy", key="du2")
            safewords2 = st.text_input(
                "Safewords",
                value="RED – stop  |  YELLOW – zwolnij  |  GREEN – ok",
                key="sw2"
            )
        
        with col2:
            st.markdown("#### Twoja część")
            if is_dom2:
                st.markdown('<span class="role-badge badge-dom">DOMINUJĄCY / DOMINUJĄCA</span>', unsafe_allow_html=True)
                needs2 = st.text_area("Czego potrzebuję od osoby uległej", height=110, key="n2")
                expectations2 = st.text_area("Czego oczekuję", height=110, key="e2")
                hard2 = st.text_area("Hard Limits (kategorycznie nie)", height=100, key="h2")
                soft2 = st.text_area("Soft Limits (do omówienia)", height=80, key="s2")
                aftercare2 = st.text_area("Aftercare i komunikacja – moje preferencje", height=80, key="a2")
            else:
                st.markdown('<span class="role-badge badge-sub">ULEGŁY / ULEGŁA</span>', unsafe_allow_html=True)
                needs2 = st.text_area("Czego potrzebuję od osoby dominującej", height=110, key="n2b")
                expectations2 = st.text_area("Czego oczekuję", height=110, key="e2b")
                hard2 = st.text_area("Hard Limits (kategorycznie nie)", height=100, key="h2b")
                soft2 = st.text_area("Soft Limits (do omówienia)", height=80, key="s2b")
                aftercare2 = st.text_area("Aftercare i komunikacja – moje preferencje", height=80, key="a2b")
        
        st.markdown("---")
        st.markdown("#### Opcjonalnie – wklej to, co partner już wypełnił (jeśli chcesz mieć pełny dokument)")
        st.caption("Możesz skopiować z podglądu powyżej odpowiednie fragmenty, żeby finalny plik zawierał obie części.")
        
        if is_dom2:
            other_needs = st.text_area("Czego potrzebuje osoba uległa (z pliku)", height=70, key="on")
            other_exp = st.text_area("Czego oczekuje osoba uległa", height=70, key="oe")
            other_hard = st.text_area("Hard Limits osoby uległej", height=70, key="oh")
            other_soft = st.text_area("Soft Limits osoby uległej", height=60, key="os")
            other_after = st.text_area("Aftercare osoby uległej", height=60, key="oa")
        else:
            other_needs = st.text_area("Czego potrzebuje osoba dominująca (z pliku)", height=70, key="on")
            other_exp = st.text_area("Czego oczekuje osoba dominująca", height=70, key="oe")
            other_hard = st.text_area("Hard Limits osoby dominującej", height=70, key="oh")
            other_soft = st.text_area("Soft Limits osoby dominującej", height=60, key="os")
            other_after = st.text_area("Aftercare osoby dominującej", height=60, key="oa")
        
        if st.button("Generuj kompletny kontrakt (.docx)", use_container_width=True, key="gen2"):
            data = {
                "dom_name": dom_name2,
                "sub_name": sub_name2,
                "scope": scope2,
                "duration": duration2,
                "safewords": safewords2,
            }
            
            if is_dom2:
                data.update({
                    "dom_needs": needs2,
                    "dom_expectations": expectations2,
                    "dom_hard_limits": hard2,
                    "dom_soft_limits": soft2,
                    "dom_aftercare": aftercare2,
                    "sub_needs": other_needs,
                    "sub_expectations": other_exp,
                    "sub_hard_limits": other_hard,
                    "sub_soft_limits": other_soft,
                    "sub_aftercare": other_after,
                })
            else:
                data.update({
                    "sub_needs": needs2,
                    "sub_expectations": expectations2,
                    "sub_hard_limits": hard2,
                    "sub_soft_limits": soft2,
                    "sub_aftercare": aftercare2,
                    "dom_needs": other_needs,
                    "dom_expectations": other_exp,
                    "dom_hard_limits": other_hard,
                    "dom_soft_limits": other_soft,
                    "dom_aftercare": other_after,
                })
            
            docx_bytes = create_contract_docx(data, "both")
            filename = f"kontrakt_Ds_kompletny_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
            
            st.markdown("""
            <div class="success-box">
            Kompletny kontrakt wygenerowany. Obie strony powinny go przeczytać i potwierdzić.
            </div>
            """, unsafe_allow_html=True)
            
            st.download_button(
                label="Pobierz kompletny kontrakt (.docx)",
                data=docx_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                key="dl2"
            )

# Stopka
st.markdown("---")
st.markdown(
    "<div style='text-align:center; color:#666; font-size:0.8rem; padding:1rem 0;'>"
    "Narzędzie edukacyjne i komunikacyjne • Nie stanowi porady prawnej • Zawsze stawiaj bezpieczeństwo i zgodę na pierwszym miejscu"
    "</div>",
    unsafe_allow_html=True
)
