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
from reportlab.lib.enums import TA_RIGHT





styles = getSampleStyleSheet()

pdfmetrics.registerFont(TTFont("dejavu", "DejaVuSansCondensed.ttf"))
pdfmetrics.registerFont(TTFont("Cairo_medium", "Cairo-Medium.ttf"))


class ImageAndHeader(Flowable):
    def __init__(self, img_path, header_text, width, height):
        Flowable.__init__(self)
        self.img_path = img_path
        self.header_text = header_text
        self.width = width
        self.height = height

    def draw(self):
        canvas = self.canv
        # Draw background image
        canvas.drawImage(self.img_path, 0, 0, self.width, self.height)

        # Draw header text
        canvas.setFont("dejavu", 24)
        canvas.setFillColor(colors.white)
        reshaped_text = arabic_reshaper.reshape(self.header_text)
        bidi_text = get_display(reshaped_text)
        text_width = stringWidth(bidi_text, "dejavu", 24)
        canvas.drawString((self.width - text_width)/2, self.height - 0.5 * inch, bidi_text)


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

    # Story content
    story = []

    styleN = styles['Normal']
    styleH_1 = styles['Heading1']
    header_text_style = ParagraphStyle('header_style', parent=styleN, fontName="dejavu")
    
    header_text_style_right = ParagraphStyle(
        'HeaderTextRight',
        parent=header_text_style,
        alignment=2,  # 2 represents right alignment,
        fontName="dejavu"
    )

    arabic_text_name = str(data[0])
    fields_to_check = {
        "monthly_evaluation" : str(data.get("التقييم الشهري للطالب","")),
        "monthly_evaluation_1" : str(data.get("التقييم الشهري للطالب_1","")),
        "monthly_evaluation_2" : str(data.get("التقييم الشهري للطالب_2","")),
        "monthly_evaluation_3" : str(data.get("التقييم الشهري للطالب_3","")),
        "monthly_evaluation_4" : str(data.get("التقييم الشهري للطالب_4","")),
        "monthly_evaluation_5" : str(data.get("التقييم الشهري للطالب_5","")),
        "monthly_evaluation_6" : str(data.get("التقييم الشهري للطالب_6","")),
        "monthly_evaluation_7" : str(data.get( "التقييم الشهري للطالب_7","")),

    }
    
    # Reshape and apply bidi algorithm
    reshaped_student_name_label = arabic_reshaper.reshape("الاسم")
    rehaped_text_name = arabic_reshaper.reshape(arabic_text_name)


    bidi_text_name = get_display(rehaped_text_name)
    bidi_text_student_name_label = get_display(reshaped_student_name_label)
    student_name_data = [["", Paragraph(bidi_text_name, header_text_style), Paragraph(bidi_text_student_name_label, header_text_style_right)]]
    
    student_name_table = Table(student_name_data)

    # Define the table style
    table_style = TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),  # Align all cells to the right
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),  # Vertically align content to the middle
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),  # Set text color
    ])

    # Apply the style to the table
    student_name_table.setStyle(table_style)



    # Calculate attendance percentages
    presence_percentage, middle_percentage, absence_percentage, presence_count, middle_count, absence_count = calculate_attendance_percentage(data, attendance_columns)

    # PDF Path
    pdf_path = os.path.join(folder_name, f'student_{i}.pdf')

    # Create the document
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=letter,
        leftMargin=0,
        rightMargin=0,
        bottomMargin=0,
        topMargin=1 * inch
    )


    # Add the ImageAndHeader at the top
    image_and_header = ImageAndHeader('./media/about-bg.png', "تقرير الطالب الشهري", letter[0], 2 * inch)
    story.append(image_and_header)
    story.append(Spacer(1, 20))

    # Student name
    story.append(student_name_table)
    story.append(Spacer(1, 20))

    total_evaluations_label = "التقييم الشهري من المعلمين"
    reshaped_total_evaluations_label = arabic_reshaper.reshape(total_evaluations_label)
    bidi_total_evaluations_label = get_display(reshaped_total_evaluations_label)
    story.append(Paragraph(bidi_total_evaluations_label, header_text_style_right))
    story.append(Spacer(1, 20))

    total_monthly_evaluations = []
    for label, value in fields_to_check.items():
        if value:
            reshaped_label = arabic_reshaper.reshape(label)
            reshaped_value = arabic_reshaper.reshape(str(value))
            bidi_label = get_display(reshaped_label)
            bidi_value = get_display(reshaped_value)
            if bidi_value != "nan":
                total_monthly_evaluations.append(bidi_value)
    
    if total_monthly_evaluations:
        combined_values = ", ".join(total_monthly_evaluations)
        right_aligned_style = ParagraphStyle(
            'RightAligned',
            parent=header_text_style,
            alignment=TA_RIGHT
        )
        story.append(Paragraph(combined_values, right_aligned_style))




    # Attendance table
    attendance_data = [
        ["Metric", "Count", "Percentage"],
        ["Present:", presence_count, f"{presence_percentage:.2f}%"],
        ["Moderately present:", middle_count, f"{middle_percentage:.2f}%"],
        ["Absent:", absence_count, f"{absence_percentage:.2f}%"],
    ]
    attendance_table = Table(attendance_data)
    attendance_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, -1), 'dejavu'),
        ('FONTSIZE', (0, 0), (-1, -1), 12),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ]))
    story.append(Spacer(1, 20))
    story.append(attendance_table)

    # Build the document
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

def detect_and_rename_column(df, keyword):
    for col in df.columns:
        if keyword in col:
            df.rename(columns={col: keyword}, inplace=True)
            return df
        if df[col].astype(str).str.contains(keyword).any():
            df.columns = df.iloc[0]
            df = df[1:]
            df.rename(columns={col: keyword}, inplace=True)
            return df
    return df

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

    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

    if uploaded_file is not None:
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names

        for sheet_name in sheet_names:
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
            df = adjust_columns_if_needed(df)
            df.columns = deduplicate_columns(df.columns)
            #st.dataframe(df)
            cropped_df = crop_df_to_student_name_header(df)
            #st.dataframe(cropped_df)
            
            attendance_columns = identify_attendance_columns(cropped_df)
            if not attendance_columns:
                #st.error(f"No attendance columns found in sheet: {sheet_name}")
                continue
            
            output_dir = os.path.join('generated_docs', sheet_name)
            os.makedirs(output_dir, exist_ok=True)

            try:
                total_students = len(cropped_df)
                progress_bar = st.progress(0)
                status_text = st.empty()

                for i, (index, row) in enumerate(cropped_df.iterrows(), start=1):
                    pdf_path = os.path.join(output_dir, f'student_{i}.pdf')
                    generate_arabic_pdf(row, i, output_dir, attendance_columns)
                    
                    if os.path.exists(pdf_path):
                        status_text.text(f"Generated PDF for student {i}/{total_students} in sheet: {sheet_name}")
                    else:
                        status_text.warning(f"PDF not generated for student {i}/{total_students} in sheet: {sheet_name}")
                    
                    # Update progress bar
                    progress = int(i / total_students * 100)
                    progress_bar.progress(progress)

                progress_bar.empty()
                status_text.success(f"PDF generation completed for sheet: {sheet_name}")
                st.text(f"The following are the attendance columns: {attendance_columns}")
                    
            except Exception as e:
                st.error(f"An error occurred while processing sheet {sheet_name}: {e}")


if __name__ == '__main__':
    main()