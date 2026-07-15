import streamlit as st
import re
from io import BytesIO
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# ==========================================================
# CONFIGURATION DE L'INTERFACE WEB
# ==========================================================
st.set_page_config(page_title="Générateur de Cours LPA", page_icon="📘", layout="centered")

st.title("📘 Générateur de Cours Word - LPA")
st.write("Collez votre brouillon formaté avec des emojis ci-dessous pour obtenir instantanément votre document Word avec toutes les couleurs, le gras et l'italique.")

# ==========================================================
# VARIABLES DE COULEURS ET EMOJIS
# ==========================================================
COLORS = {
    "primary": RGBColor(24, 64, 130),
    "secondary": RGBColor(201, 146, 37),
}

EMOJI_MAP = {
    '📌': 'F2F2F2', '🌍': 'FCE4D6', '▶': 'DDEBF7', '📖': 'E2EFDA',
    '📘': 'E2F0D9', '⚠': 'FFF2CC', '📊': 'F2F2F2', '◆': 'D9EAF7',
    '✎': 'FFF2CC', '💡': 'EAF2F8', '◉': 'EAF2F8', '◎': 'FFFFFF',
    '◈': 'FFF9EE', '✔': 'E2F0D9'
}

# ==========================================================
# FONCTIONS DU MOTEUR WORD
# ==========================================================
def add_formatted_text(paragraph, text):
    """Analyse les balises **gras** et *italique* pour le Word"""
    if text.startswith('* '): text = text[2:]
    elif text.startswith('- '): text = text[2:]
    elif text.startswith('· '): text = text[2:]
        
    parts = re.split(r'(\*\*.*?\*\*|\*.*?\*)', text)
    for part in parts:
        if part.startswith('**') and part.endswith('**'):
            run = paragraph.add_run(part[2:-2])
            run.bold = True
        elif part.startswith('*') and part.endswith('*'):
            run = paragraph.add_run(part[1:-1])
            run.italic = True
        else:
            paragraph.add_run(part)

def split_into_chunks(texte):
    lignes = texte.split('\n')
    chunks, current_chunk = [], []
    triggers = ['Section ', 'SECTION ', 'Phase ', 'PHASE ', 'Sommaire', 'SOMMAIRE'] + list(EMOJI_MAP.keys())
    
    for ligne in lignes:
        ligne_strip = ligne.strip("┌─└│┐┘━_ \t")
        if not ligne_strip:
            if current_chunk:
                chunks.append(current_chunk)
                current_chunk = []
            continue
        is_trigger = any(ligne_strip.startswith(t) for t in triggers)
        if is_trigger and current_chunk:
            chunks.append(current_chunk)
            current_chunk = []
        current_chunk.append(ligne_strip)
    if current_chunk:
        chunks.append(current_chunk)
    return chunks

def create_colored_box(doc, lignes, couleur):
    if couleur == 'FFFFFF':
        doc.add_heading(lignes[0].replace('**', '').replace('*', ''), level=3)
        for ligne in lignes[1:]:
            p = doc.add_paragraph(style='List Bullet') if ligne.startswith(('*', '-', '·')) else doc.add_paragraph()
            add_formatted_text(p, ligne)
        return

    table = doc.add_table(rows=1, cols=1)
    table.style = "Table Grid"
    cell = table.cell(0, 0)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), couleur)
    cell._tc.get_or_add_tcPr().append(shd)
    
    p = cell.paragraphs[0]
    run = p.add_run(lignes[0].replace('**', '').replace('*', ''))
    run.bold = True
    
    for ligne in lignes[1:]:
        p = cell.add_paragraph()
        if ligne.startswith(('* ', '- ', '· ')):
            add_formatted_text(p, "• " + ligne[2:])
        else:
            add_formatted_text(p, ligne)
    doc.add_paragraph()

def render_chunk(doc, chunk_lines):
    if not chunk_lines: return
    first_line = chunk_lines[0]
    first_line_upper = first_line.upper()

    if "LPA" in first_line or "Lycée Personnel" in first_line:
        for i, line in enumerate(chunk_lines):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(line.replace('**', '').replace('*', ''))
            if i == 0 or line.isupper():
                run.bold, run.font.size, run.font.color.rgb = True, Pt(16), COLORS["primary"]
            elif "·" in line and ("Leçon" in line or "Classe" in line):
                run.bold, run.font.size = True, Pt(12)
        doc.add_paragraph()
        return

    for emoji, color in EMOJI_MAP.items():
        if first_line.startswith(emoji):
            create_colored_box(doc, chunk_lines, color)
            return

    if first_line_upper.startswith("SECTION"):
        doc.add_heading(first_line.replace('**', '').replace('*', ''), level=1)
        for line in chunk_lines[1:]:
            p = doc.add_paragraph(style='List Bullet') if line.startswith(('* ', '- ', '· ')) else doc.add_paragraph()
            add_formatted_text(p, line)
        return

    if first_line_upper.startswith("PHASE") or first_line_upper == "SOMMAIRE":
        doc.add_heading(first_line.replace('**', '').replace('*', ''), level=2)
        for line in chunk_lines[1:]:
            p = doc.add_paragraph(style='List Bullet') if line.startswith(('* ', '- ', '· ')) else doc.add_paragraph()
            add_formatted_text(p, line)
        return

    if first_line.startswith('"') or first_line.startswith('«'):
        doc.add_paragraph()
        for line in chunk_lines:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_formatted_text(p, line)
        return

    for line in chunk_lines:
        p = doc.add_paragraph(style='List Bullet') if line.startswith(('* ', '- ', '· ')) else doc.add_paragraph()
        add_formatted_text(p, line)

def generate_word(cours_texte):
    doc = Document()
    section = doc.sections[0]
    section.top_margin, section.bottom_margin = Cm(1.5), Cm(1.5)
    section.left_margin, section.right_margin = Cm(1.5), Cm(1.5)

    styles = doc.styles
    styles["Heading 1"].font.size = Pt(16)
    styles["Heading 1"].font.color.rgb = COLORS["primary"]
    styles["Heading 1"].font.bold = True
    styles["Heading 2"].font.size = Pt(14)
    styles["Heading 2"].font.color.rgb = COLORS["secondary"]
    styles["Heading 2"].font.bold = True
    styles["Heading 3"].font.size = Pt(12)
    styles["Heading 3"].font.color.rgb = COLORS["primary"]
    styles["Heading 3"].font.bold = True

    chunks = split_into_chunks(cours_texte)
    for chunk in chunks:
        render_chunk(doc, chunk)

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# ==========================================================
# INTERFACE UTILISATEUR
# ==========================================================
texte_input = st.text_area("✍️ Collez le texte complet de votre cours ici :", height=400)

if st.button("🚀 Générer le fichier Word"):
    if texte_input.strip():
        with st.spinner("Fabrication du document avec mise en forme en cours..."):
            docx_file = generate_word(texte_input)
            
            # Récupérer automatiquement un nom de fichier logique
            filename = "Cours_Généré.docx"
            for line in texte_input.split('\n')[:10]:
                if "·" in line and ("Leçon" in line or "Classe" in line):
                    parts = line.split("·")
                    for p in parts:
                        if "Leçon" in p:
                            filename = p.strip().replace(" ", "_") + ".docx"
                            break
                            
            st.success("✅ Document prêt ! Le gras, l'italique et les couleurs ont été appliqués.")
            st.download_button(
                label="📥 TÉLÉCHARGER LE FICHIER WORD",
                data=docx_file,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    else:
        st.warning("⚠️ Veuillez coller du texte avant de cliquer sur le bouton.")
