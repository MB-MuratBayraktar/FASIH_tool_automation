import os
import streamlit as st
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import arabic_reshaper
from reportlab.pdfgen import canvas
from reportlab.pdfbase.ttfonts import TTFont
from bidi.algorithm import get_display
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.units import inch
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.platypus import Flowable
from PIL import Image




styles = getSampleStyleSheet()

pdfmetrics.registerFont(TTFont("dejavu", "DejaVuSansCondensed.ttf"))
pdfmetrics.registerFont(TTFont("Cairo_medium", "Cairo-Medium.ttf"))



class ImageAndHeader(Flowable):
    def __init__(self, img_path, header_text, width):
        Flowable.__init__(self)
        self.img_path = img_path
        self.header_text = header_text
        self.width = width
        self.height = 3 * inch  # Adjust this value to match the desired height

    def wrap(self, availWidth, availHeight):
        return (self.width, self.height)

    def draw(self):
        canvas = self.canv
        # Draw background image
        canvas.drawImage(self.img_path, 0, 0, self.width, self.height)
        # Draw header text
        canvas.setFont("dejavu", 18)
        text_width = stringWidth(self.header_text, "dejavu", 18)
        canvas.setFillColorRGB(1, 1, 1)  # White color for text
        canvas.drawString(6*inch, self.height - 1*inch, self.header_text)

def calculate_attendance_percentage(row, attendance_columns):
  total_weeks = len(attendance_columns)
  attendance_values = row[attendance_columns].values

  present_count = sum(1 for val in attendance_values if val == "ملتزم بالحضور")
  middle_count = sum(1 for val in attendance_values if val == "متوسط الالتزام")
  # Count empty cells and None values as absent
  not_present = sum(1 for val in attendance_values if val == "لا يحضر")
  not_verified = sum(1 for val in attendance_values if val == None)
  empty_cell = sum(1 for val in attendance_values if val == "")
  absent_count = not_present + not_verified + empty_cell

  presence_percentage = (present_count / total_weeks) * 100
  middle_percentage = (middle_count / total_weeks) * 100
  absence_percentage = (absent_count / total_weeks) * 100

  return presence_percentage, middle_percentage, absence_percentage, present_count, middle_count, absent_count


def generate_arabic_pdf(data, i, folder_name, attendance_columns):
    styleN = styles['Normal']
    styleH_1 = styles['Heading1']

    header_text_style = ParagraphStyle('header_style',parent=styleN,fontName="dejavu")
    
    arabic_text_name = str(data[0])
    student_name_label = 'الاسم'

    # Reshape and apply bidi algorithm
    reshaped_student_name_label = arabic_reshaper.reshape(student_name_label)
    rehaped_text_name = arabic_reshaper.reshape(arabic_text_name)
    reshaped_header_text = arabic_reshaper.reshape("تقرير الطالب الشهري")

    bidi_text_name = get_display(rehaped_text_name)
    bidi_text_student_name_label = get_display(reshaped_student_name_label)
    bidi_text_header = get_display(reshaped_header_text)

    # Calculate attendance percentages
    presence_percentage, middle_percentage, absence_percentage, presence_count, middle_count, absence_count = calculate_attendance_percentage(data, attendance_columns)

    pdf_path = os.path.join(folder_name, f'student_{i}.pdf')
    c = canvas.Canvas(pdf_path, pagesize=letter)
    width, height = letter

    # Draw the background image
    image_path = './media/about-bg.png'
    c.drawImage(image_path, 0, height - 219, width, 219)

    
    # Draw the header text
    c.setFont("dejavu", 24)
    header_text = bidi_text_header
    text_width = stringWidth(header_text, "dejavu", 24)
    c.drawString((width - text_width), height - 1.5*inch, header_text)

    # Save the canvas state
    c.save()

    # Create the story for the rest of the content
    doc = SimpleDocTemplate(os.path.join(folder_name, f'student_{i}.pdf'),
                             pagesize=letter, leftMargin=0, rightMargin=0, bottomMargin=0, topMargin=0.5*inch)
    
    width, height = letter
    
    story = []
    image_and_header = ImageAndHeader('./media/about-bg.png', bidi_text_header, width)
    story.append(image_and_header)
    
    # Add some space after the image
    story.append(Spacer(1, 20))

    
    student_name_data = [["", Paragraph(bidi_text_name, header_text_style),Paragraph(bidi_text_student_name_label, header_text_style), ]]
    student_name_table = Table(student_name_data, colWidths=[1*inch, 2*inch])
    student_name_table.hAlign = 'RIGHT'


    # Style the table
    style = TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
        ('FONTNAME', (0, 0), (-1, -1), 'dejavu'),
        ('FONTSIZE', (0, 0), (-1, -1), 12),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
    ])

    student_name_table.setStyle(style)
    story.append(student_name_table)

    # Add attendance percentages
    story.append(Spacer(1, 14))

    data = [["Metric", "Count","Percentage"],
            [f"Total middle count:", presence_count,f"{presence_percentage:.2f}%"],
            [f"Total middle count:", middle_count,f"{middle_percentage:.2f}%"],
            [f"Total absence count:", absence_count,f"{absence_percentage:.2f}%"]]
    
    table = Table(data)
    table.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ('FONTSIZE', (0, 0), (-1, 0), 12),
    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
    ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ]))

    story.append(table)
    doc.build(story)

def crop_df_to_student_name_header(df):
    specific_word = "اسم الطالب"
    column_with_word = None
    for column in df.columns:
        if specific_word in df[column].astype(str).values:
            column_with_word = column
            break

    if column_with_word:
        columns = [column_with_word] + [col for col in df.columns if col != column_with_word]
        df = df[columns]
    return df

def identify_attendance_columns(df):
    attendance_columns = []
    for col in df.columns:
        col_str = str(col)
        if any(keyword in col_str for keyword in ["الحضور والغياب الاسبوع الأول", "الحضور والغياب الاسبوع الثاني", "الحضور والغياب الاسبوع الثالث", "الحضور والغياب الاسبوع الرابع"]):
            attendance_columns.append(col_str)
    return attendance_columns

def adjust_columns_if_needed(df):
    attendance_keywords = ["الحضور والغياب الاسبوع الأول", "الحضور والغياب الاسبوع الثاني", "الحضور والغياب الاسبوع الثالث", "الحضور والغياب الاسبوع الرابع"]
    for keyword in attendance_keywords:
        if keyword not in df.columns:
            for col in df.columns:
                if isinstance(df[col], pd.Series) and df[col].astype(str).str.contains(keyword).any():
                    df.columns = deduplicate_columns(df.iloc[0])
                    df = df[1:]
                    break
    return df

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
            # Adjust columns if needed
            df = adjust_columns_if_needed(df)
            # Deduplicate columns
            df.columns = deduplicate_columns(df.columns)
            # Display the DataFrame (optional)
            st.dataframe(df)
            cropped_df = crop_df_to_student_name_header(df)
            st.dataframe(cropped_df)
            
            # Identify attendance columns
            attendance_columns = identify_attendance_columns(cropped_df)
            if not attendance_columns:
                st.error(f"No attendance columns found in sheet: {sheet_name}")
                continue
            
            # Create a directory for the selected sheet
            output_dir = os.path.join('generated_docs', sheet_name)
            os.makedirs(output_dir, exist_ok=True)

            # Process and Generate PDFs
            try:
                for i, (index, row) in enumerate(cropped_df.iterrows(), start=1):
                    generate_arabic_pdf(row, i, output_dir, attendance_columns)
                st.success(f"PDFs generated successfully for sheet: {sheet_name}")
                st.text(f"the followings are the attendance columns: {attendance_columns}")
                    
            except Exception as e:
                st.error(f"An error occurred while processing sheet {sheet_name}: {e}")

if __name__ == '__main__':
    main()