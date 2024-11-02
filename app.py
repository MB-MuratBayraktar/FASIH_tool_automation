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

def crop_df_to_student_name_header(df, keyword="اسم الطالب"):
    # Find the row index where the keyword is found
    header_row_index = None
    for i, row in df.iterrows():
        if keyword in row.values:
            header_row_index = i
            break

    # If the keyword is not found, create it in the first column at the top
    if header_row_index is None:
        new_row = pd.DataFrame({df.columns[0]: [keyword]})
        df = pd.concat([new_row, df], ignore_index=True)
        header_row_index = 0  # New keyword is at the top row

    # Set the row containing the keyword as the new header
    df.columns = df.iloc[header_row_index]
    df = df.iloc[header_row_index + 1:].reset_index(drop=True)

    # Rename duplicate columns by adding suffixes
    def deduplicate_columns(columns):
        counts = {}
        new_columns = []
        for col in columns:
            if col in counts:
                counts[col] += 1
                new_columns.append(f"{col}_{counts[col]}")
            else:
                counts[col] = 0
                new_columns.append(col)
        return new_columns
    
    df.columns = deduplicate_columns(df.columns)
    
    return df



def generate_arabic_pdf(data, i, folder_name):
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

    arabic_text_name = str(data['اسم الطالب'])
    class_presence_week_1 = str(data[3])
    class_presence_week_2 = str(data[4])
    class_presence_week_3 = str(data[5])
    class_presence_week_4 = str(data[6])

    # Reshape and apply bidi algorithm
    rehaped_text_name = arabic_reshaper.reshape(arabic_text_name)
    rehaped_text_presence_week_1 = arabic_reshaper.reshape(class_presence_week_1)
    rehaped_text_presence_week_2 = arabic_reshaper.reshape(class_presence_week_2)
    rehaped_text_presence_week_3 = arabic_reshaper.reshape(class_presence_week_3)
    rehaped_text_presence_week_4 = arabic_reshaper.reshape(class_presence_week_4)

    bidi_text_name = get_display(rehaped_text_name)
    bidi_text_presence_week_1 = get_display(rehaped_text_presence_week_1)
    bidi_text_presence_week_2 = get_display(rehaped_text_presence_week_2)
    bidi_text_presence_week_3 = get_display(rehaped_text_presence_week_3)
    bidi_text_presence_week_4 = get_display(rehaped_text_presence_week_4)

    story = []
    story.append(Paragraph("Student name",styleH_1))
    story.append(Spacer(1,8))
    story.append(Paragraph(bidi_text_name,arabic_text_style))

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
    
    doc = SimpleDocTemplate(os.path.join(folder_name,f'student_{i}.pdf'), pagesize=letter)
    doc.build(story)

def main():
    st.title("PDF Generator from Excel")

    # File Uploader
    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

    if uploaded_file is not None:
        # Get the sheet names
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names

        for sheet_name in sheet_names:
            # Read the selected sheet into a DataFrame
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
            # Display the DataFrame (optional)
            st.dataframe(df)
            cropped_df = crop_df_to_student_name_header(df)
            st.dataframe(cropped_df)
            
            # Create a directory for the selected sheet
            output_dir = os.path.join('generated_docs', sheet_name)
            os.makedirs(output_dir, exist_ok=True)

            # Process and Generate PDFs
            try:
                for i, (index, row) in enumerate(cropped_df.iterrows(), start=1):
                    generate_arabic_pdf(row, i, output_dir)
                st.success(f"PDFs generated successfully for sheet: {sheet_name}")
                    
            except Exception as e:
                st.error(f"An error occurred while processing sheet {sheet_name}: {e}")



if __name__ == '__main__':
    main()