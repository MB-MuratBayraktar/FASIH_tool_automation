import streamlit as st
import pandas as pd
import pdfkit
from fpdf import FPDF
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import letter
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer,PageBreak
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
import reportlab
from reportlab.pdfbase.ttfonts import TTFont
import arabic_reshaper
from bidi.algorithm import get_display
styles = getSampleStyleSheet()

pdfmetrics.registerFont(TTFont("Khayal", "Khayal-Font-Demo.ttf"))


def generate_arabic_pdf(data, i):
    styleN = styles['Normal']
    styleH_1 = styles['Heading1']

    arabic_text_style = ParagraphStyle(
        'border',
        parent=styleN,
        borderColor='#333333',
        borderWidth=1,
        borderPadding=2,
        fontName="Khayal"
    )

    # Extract Arabic content from the DataFrame
    arabic_text_name = data['اسم الطالب']
    arabic_text_status = data['الحالة']

    # Reshape and apply bidi algorithm
    rehaped_text_name = arabic_reshaper.reshape(arabic_text_name)
    rehaped_text_status = arabic_reshaper.reshape(arabic_text_status)

    bidi_text_name = get_display(rehaped_text_name)
    bidi_text_status = get_display(rehaped_text_status)

    story = []
    story.append(Paragraph("Perfect arabic text ",styleH_1))
    story.append(Spacer(1,8))
    story.append(Paragraph(bidi_text_name,arabic_text_style))

    story.append(Paragraph("Perfect arabic text ",styleH_1))
    story.append(Spacer(1,8))
    story.append(Paragraph(bidi_text_status,arabic_text_style))

    
    doc = SimpleDocTemplate(f'student_{i}.pdf', pagesize=letter)
    doc.build(story)

    
def generate_pdf(data,i):
    # Create an FPDF object
    pdf = FPDF()
    pdf.add_page()

    # Add your PDF template design here, using PDF's cell and multi_cell methods
    # Example:
    pdf.set_font("Arial", size=12)
    pdf.cell(200, 10, txt="Student : " + str(data['اسم الطالب']), ln=1, align='L')
    pdf.cell(200, 10, txt="Student status : " + str(data['الحالة']), ln=1, align='L')
    

    # Save the PDF
    pdf.output(f"student_{i}.pdf")

def main():
    st.title("PDF Generator from Excel")

    # File Uploader
    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

    if uploaded_file is not None:
        i=0
        df = pd.read_excel(uploaded_file)

        # Display the DataFrame (optional)
        st.dataframe(df)

        # Process and Generate PDFs
        for index, row in df.iterrows():
            i += 1
            generate_arabic_pdf(row, i)


        st.success("PDFs generated successfully!")

if __name__ == '__main__':
    main()