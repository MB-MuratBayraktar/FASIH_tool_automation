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
import os
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
    arabic_text_name = str(data['اسم الطالب'])
    arabic_text_status = str(data['الحالة'])
    class_presence_week_1 = str(data[3])
    class_presence_week_2 = str(data[4])
    class_presence_week_3 = str(data[5])
    class_presence_week_4 = str(data[6])




    # Reshape and apply bidi algorithm
    rehaped_text_name = arabic_reshaper.reshape(arabic_text_name)
    rehaped_text_status = arabic_reshaper.reshape(arabic_text_status)
    rehaped_text_presence_week_1 = arabic_reshaper.reshape(class_presence_week_1)
    rehaped_text_presence_week_2 = arabic_reshaper.reshape(class_presence_week_2)
    rehaped_text_presence_week_3 = arabic_reshaper.reshape(class_presence_week_3)
    rehaped_text_presence_week_4 = arabic_reshaper.reshape(class_presence_week_4)

    bidi_text_name = get_display(rehaped_text_name)
    bidi_text_status = get_display(rehaped_text_status)
    bidi_text_presence_week_1 = get_display(rehaped_text_presence_week_1)
    bidi_text_presence_week_2 = get_display(rehaped_text_presence_week_2)
    bidi_text_presence_week_3 = get_display(rehaped_text_presence_week_3)
    bidi_text_presence_week_4 = get_display(rehaped_text_presence_week_4)

    story = []
    story.append(Paragraph("Student name",styleH_1))
    story.append(Spacer(1,8))
    story.append(Paragraph(bidi_text_name,arabic_text_style))

    story.append(Paragraph("student status ",styleH_1))
    story.append(Spacer(1,8))
    story.append(Paragraph(bidi_text_status,arabic_text_style))

    story.append(Paragraph("student presence week 1:",styleH_1))
    story.append(Spacer(1,8))
    story.append(Paragraph(bidi_text_presence_week_1,arabic_text_style))

    story.append(Paragraph("student presence week 2:",styleH_1))
    story.append(Spacer(1,8))
    story.append(Paragraph(bidi_text_presence_week_2,arabic_text_style))
    
    story.append(Paragraph("student presence week 3:",styleH_1))
    story.append(Spacer(1,8))
    story.append(Paragraph(bidi_text_presence_week_3,arabic_text_style))

    story.append(Paragraph("student presence week 4:",styleH_1))
    story.append(Spacer(1,8))
    story.append(Paragraph(bidi_text_presence_week_4,arabic_text_style))

    output_dir = 'generated_docs'
    os.makedirs(output_dir, exist_ok=True)

    doc = SimpleDocTemplate(os.path.join("generated_docs",f'student_{i}.pdf'), pagesize=letter)
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
        # Get the sheet names
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names

        # Display the sheet names and allow the user to select one
        sheet_name = st.selectbox("Select a sheet", sheet_names)

        # Read the selected sheet into a DataFrame
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name)

        # Display the DataFrame (optional)
        st.dataframe(df)

        # Process and Generate PDFs
        for i, (index, row) in enumerate(df.iterrows(), start=1):
            generate_arabic_pdf(row, i)

        st.success("PDFs generated successfully!")

if __name__ == '__main__':
    main()