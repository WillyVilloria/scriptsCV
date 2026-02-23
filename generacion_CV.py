import json
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from datetime import datetime
from docx.shared import Cm

# Crear documento

class Documento:
    def __init__(self):
        self.doc = Document()
        self.fecha = datetime.now()

        with open('texto_comun.json', 'r', encoding='utf-8') as f:
            self.textos_json = json.load(f)

    def encabezado(self):
        # margenes
        section = self.doc.sections[0]
        section.header_distance = Cm(1)
        self.__margenes(section)
        # Acceder al encabezado de la primera sección
        header = self.doc.sections[0].header
        # Limpiar contenido anterior (opcional)
        header.is_linked_to_previous = False
        header_para = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
        header_run = header_para.add_run("")
        header_run.font.size = Pt(10)         # tamaño de fuente
        header_run.font.bold = True
        header_run.font.color.rgb = RGBColor(0, 51, 102)   

    def __margenes(self, section):
        # Ajustar márgenes
        section.top_margin = Cm(2.0)       # margen superior
        section.bottom_margin = Cm(2.0)    # margen inferior
        section.left_margin = Cm(2.5)        # margen izquierdo
        section.right_margin = Cm(2.0)       # margen derecho

    def pie(self):
        # Acceder al pie de página de la primera sección
        footer = self.doc.sections[0].footer
        footer.is_linked_to_previous = False
        footer_para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        footer_run = footer_para.add_run("CV Miguel A. Lorenzo Villoria \t")
        footer_run.font.size = Pt(8)         # tamaño de fuente
        footer_run.font.italic = True
        footer_run.font.color.rgb = RGBColor(128, 128, 255)
        #Cambiar altura del pie de página
        section = self.doc.sections[0]
        section.footer_distance = Cm(1)

        self.__numero_pagina(footer_para)

    def __numero_pagina(self, footer_para):
        # Insertar "Página X de Y"
        footer_para.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT 
        footer_para.add_run("Página ")
        # Campo PAGE
        run = footer_para.add_run()
        run.font.size = Pt(10)
        fldChar1 = OxmlElement("w:fldChar")
        fldChar1.set(qn("w:fldCharType"), "begin")

        instrText = OxmlElement("w:instrText")
        instrText.text = "PAGE"

        fldChar2 = OxmlElement("w:fldChar")
        fldChar2.set(qn("w:fldCharType"), "end")

        run._r.append(fldChar1)
        run._r.append(instrText)
        run._r.append(fldChar2)

        footer_para.add_run(" de ")

        # Campo NUMPAGES
        run2 = footer_para.add_run()
        fldChar1 = OxmlElement("w:fldChar")
        fldChar1.set(qn("w:fldCharType"), "begin")

        instrText = OxmlElement("w:instrText")
        instrText.text = "NUMPAGES"

        fldChar2 = OxmlElement("w:fldChar")
        fldChar2.set(qn("w:fldCharType"), "end")

        run2._r.append(fldChar1)
        run2._r.append(instrText)
        run2._r.append(fldChar2)

    def estilo(self):
        # Acceder al estilo Normal
        normal_style = self.doc.styles['Normal']

        # Cambiar fuente
        font = normal_style.font # pyright: ignore[reportAttributeAccessIssue]
        font.name = 'Cambria'
        font.size = Pt(11)
        font.bold = False
        font.color.rgb = RGBColor(0, 0, 0)  # negro

        # Cambiar párrafo base (alineación, espaciado, etc.)
        paragraph_format = normal_style.paragraph_format # pyright: ignore[reportAttributeAccessIssue]
        paragraph_format.space_after = Pt(6)
        paragraph_format.space_before = Pt(6)
        paragraph_format.line_spacing = 1.15

    # Función para añadir un título con color y estilo
    def add_colored_heading(self, text, level=1, color=RGBColor(0, 51, 102)):
        # Crear un heading vacío → esto devuelve un Paragraph
        paragraph = self.doc.add_heading("", level=level)

        # Acceder al estilo correspondiente
        style_name = f"Heading {level}"
        heading_style = self.doc.styles[style_name]

        # Modificar formato de párrafo del estilo (afecta a todos los títulos de ese nivel)
        para_format = heading_style.paragraph_format # pyright: ignore[reportAttributeAccessIssue]
        para_format.space_before = Pt(16)   # Espaciado anterior
        para_format.space_after = Pt(8)    # Espaciado posterior

        # Agregar el texto como Run dentro del párrafo
        run = paragraph.add_run(text)
        run.font.color.rgb = color
        run.font.bold = True

        return paragraph

    def cabecera(self):
        # ===== CABECERA =====
        name = self.doc.add_heading("Miguel Ángel Lorenzo Villoria", 0)
        name.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        contact = self.doc.add_paragraph(
            "Gijón, Asturias   |   miguelcarter@gmail.com   |    635563966   |   Nacionalidad: Española"
        )
        contact.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

        #self.doc.add_paragraph("")

    def Perfil_prof(self):
        # ===== PERFIL PROFESIONAL =====
        self.add_colored_heading("Perfil Profesional", 1)
        perfil1 = self.doc.add_paragraph(
            self.textos_json["perfil_profesional"]["perfil1"]
        )
        perfil1.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        perfil2 = self.doc.add_paragraph(
            self.textos_json["perfil_profesional"]["perfil2"]
        )
        perfil2.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        perfil3 = self.doc.add_paragraph(
            self.textos_json["perfil_profesional"]["perfil3"]
        )
        perfil3.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        perfil4 = self.doc.add_paragraph(
            self.textos_json["perfil_profesional"]["perfil4"]
            
        )
        perfil4.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        perfil5 = self.doc.add_paragraph(
            self.textos_json["perfil_profesional"]["perfil5"]
        )
        perfil5.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    def experiencia_prof(self):
        # ===== EXPERIENCIA PROFESIONAL =====
        self.add_colored_heading("Experiencia Profesional", 1)

        exp1 = self.doc.add_paragraph()
        exp1.add_run(self.textos_json["experiencia_profesional"]["experiencia1"]["titulo"]).bold = True
        experiencia = self.textos_json["experiencia_profesional"]["experiencia1"]["desarrollo"]
        for item in experiencia:
            bullet = self.doc.add_paragraph(f"•\t{item}")
            bullet.paragraph_format.left_indent = Pt(18)         # sangría para alinear con el texto
            bullet.paragraph_format.first_line_indent = Pt(-18)  # primera línea “sale” la viñeta
            bullet.paragraph_format.space_after = Pt(0)
            bullet.paragraph_format.line_spacing = 1.0 

        exp2 = self.doc.add_paragraph()
        exp2.add_run(self.textos_json["experiencia_profesional"]["experiencia2"]["titulo"]).bold = True
        operaciones = self.textos_json["experiencia_profesional"]["experiencia2"]["desarrollo"]
        for item in operaciones:
            bullet = self.doc.add_paragraph(f"•\t{item}")
            bullet.paragraph_format.left_indent = Pt(18)         # sangría para alinear con el texto
            bullet.paragraph_format.first_line_indent = Pt(-18)  # primera línea “sale” la viñeta
            bullet.paragraph_format.space_after = Pt(0)
            bullet.paragraph_format.line_spacing = 1.0 
        
        exp3 = self.doc.add_paragraph()
        exp3.add_run(self.textos_json["experiencia_profesional"]["experiencia3"]["titulo"]).bold = True
        exp3.add_run(self.textos_json["experiencia_profesional"]["experiencia3"]["desarrollo"])

        exp4 = self.doc.add_paragraph()
        exp4.add_run(self.textos_json["experiencia_profesional"]["experiencia4"]["titulo"]).bold = True
        exp5 = self.doc.add_paragraph()
        exp5.add_run(self.textos_json["experiencia_profesional"]["experiencia5"]["titulo"]).bold = True

        exp6 = self.doc.add_paragraph()
        exp6.add_run(self.textos_json["experiencia_profesional"]["experiencia6"]["titulo"]).bold = True

        exp7 = self.doc.add_paragraph()
        exp7.add_run(self.textos_json["experiencia_profesional"]["experiencia7"]["titulo"]).bold = True

    def logros(self):
        # ===== LOGROS DESTACADOS =====
        self.add_colored_heading("Logros Destacados", 1)
        logros = self.textos_json["logros_destacados"]["logros"]
        
        for item in logros:
            bullet = self.doc.add_paragraph(f"•\t{item}")
            bullet.paragraph_format.left_indent = Pt(20)         # sangría para alinear con el texto
            bullet.paragraph_format.first_line_indent = Pt(-18)  # primera línea “sale” la viñeta
            bullet.paragraph_format.space_after = Pt(0)
            bullet.paragraph_format.line_spacing = 1.0 

    def formacion(self):
        # ===== FORMACIÓN =====
        self.add_colored_heading("Formación Académica", 1)
        formacion = self.textos_json["formacion_academica"]["formacion"]
        for item in formacion:
            bullet = self.doc.add_paragraph()
            run = bullet.add_run(f"•\t{item}")
            run.bold = True
            bullet.paragraph_format.left_indent = Pt(20)         # sangría para alinear con el texto
            bullet.paragraph_format.first_line_indent = Pt(-18)  # primera línea “sale” la viñeta
            bullet.paragraph_format.space_after = Pt(0)
            bullet.paragraph_format.line_spacing = 1.0 
        otros_cursos = "Otros cursos: Máster en Gestión de Calidad y Medio Ambiente, Técnico de Prevención de Riesgos Laborales."
        self.doc.add_paragraph(otros_cursos)

    def habilidades(self):
        # ===== HABILIDADES TÉCNICAS =====
        self.add_colored_heading("Habilidades Técnicas", 1)
        habilidades = self.textos_json["habilidades_tecnicas"]["habilidades"]
        for item in habilidades:
            bullet = self.doc.add_paragraph(f"•\t{item}")
            bullet.paragraph_format.left_indent = Pt(20)         # sangría para alinear con el texto
            bullet.paragraph_format.first_line_indent = Pt(-18)  # primera línea “sale” la viñeta
            bullet.paragraph_format.space_after = Pt(0)
            bullet.paragraph_format.line_spacing = 1.0 

    def idiomas(self):
        # ===== IDIOMAS =====
        self.add_colored_heading("Idiomas", 1)
        self.doc.add_paragraph("• " + self.textos_json["idiomas"]["idioma1"][0] + "\n• " + self.textos_json["idiomas"]["idioma2"][0])

    def guardar(self):
        # Guardar documento
        output_path_visual = f"../../CV/CV_Miguel_Angel_Lorenzo_Villoria_Data_Visual_{self.fecha.month}_{self.fecha.day}.docx"
        self.doc.save(output_path_visual)

        #output_path_visual

if __name__ == "__main__":
    document = Documento()
    document.encabezado()
    document.pie()
    document.estilo()
    document.cabecera()
    document.Perfil_prof()
    document.experiencia_prof()
    document.logros()
    document.formacion()
    document.habilidades()
    document.idiomas()
    document.guardar()