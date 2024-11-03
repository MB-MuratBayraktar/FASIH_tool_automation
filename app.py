import os
import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import arabic_reshaper
from reportlab.pdfbase.ttfonts import TTFont
from bidi.algorithm import get_display
from reportlab.pdfbase import pdfmetrics
styles = getSampleStyleSheet()

pdfmetrics.registerFont(TTFont("Khayal", "Khayal-Font-Demo.ttf"))
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

  return presence_percentage, middle_percentage, absence_percentage, present_count, middle_count, absent_count, empty_cell, not_verified, not_present


def generate_arabic_pdf(data, i, folder_name, attendance_columns):
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
    
    arabic_text_name = str(data[0])
    class_presence_week_1 = str(data[attendance_columns[0]])
    class_presence_week_2 = str(data[attendance_columns[1]])
    class_presence_week_3 = str(data[attendance_columns[2]])
    class_presence_week_4 = str(data[attendance_columns[3]])

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

    # Calculate attendance percentages
    presence_percentage, middle_percentage, absence_percentage, presence_count, middle_count, absence_count, empty_cell, not_verified, not_present = calculate_attendance_percentage(data, attendance_columns)

    story = []
    image_path = './logo.png'  # Replace with your image path
    img = Image(image_path)
    img.hAlign = 'CENTER'
    story.append(img)
    story.append(Spacer(1, 12))

    story.append(Paragraph("Student name", styleH_1))
    story.append(Spacer(1, 8))
    story.append(Paragraph(bidi_text_name, arabic_text_style))
    story.append(Spacer(1, 14))

    story.append(Paragraph("student presence week 1:", styleH_1))
    story.append(Spacer(1, 8))
    story.append(Paragraph(bidi_text_presence_week_1, arabic_text_style))

    story.append(Paragraph("student presence week 2:", styleH_1))
    story.append(Spacer(1, 8))
    story.append(Paragraph(bidi_text_presence_week_2, arabic_text_style))
    
    story.append(Paragraph("student presence week 3:", styleH_1))
    story.append(Spacer(1, 8))
    story.append(Paragraph(bidi_text_presence_week_3, arabic_text_style))

    story.append(Paragraph("student presence week 4:", styleH_1))
    story.append(Spacer(1, 8))
    story.append(Paragraph(bidi_text_presence_week_4, arabic_text_style))

    # Add attendance percentages
    story.append(Spacer(1, 14))
    story.append(Paragraph(f"Presence Percentage: {presence_percentage:.2f}%", styleN))
    story.append(Paragraph(f"Middle Percentage: {middle_percentage:.2f}%", styleN))
    story.append(Paragraph(f"Absence Percentage: {absence_percentage:.2f}%", styleN))
    story.append(Spacer(1, 14))
    story.append(Paragraph(f"Total presence count: {presence_count}", styleN))
    story.append(Paragraph(f"Total middle count: {middle_count}", styleN))
    story.append(Paragraph(f"Total absence count: {absence_count}", styleN))
    story.append(Paragraph(f"Total not verified count: {not_verified}", styleN))
    story.append(Paragraph(f"Total not present count: {not_present}", styleN))
    story.append(Paragraph(f"Total empty cell count: {empty_cell}", styleN))
    
    doc = SimpleDocTemplate(os.path.join(folder_name, f'student_{i}.pdf'), pagesize=letter)
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