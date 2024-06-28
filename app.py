import matplotlib as mpl
mpl.use('Agg')
import os
from flask import Flask, render_template, request, redirect, url_for, send_file
import pandas as pd
import matplotlib.pyplot as plt
from reportlab.lib.utils import ImageReader
import seaborn as sns
from io import BytesIO
import base64
import threading
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import json

from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, Frame, PageTemplate, PageBreak

app = Flask(__name__)

DASHBOARD_DATA_FILE = 'dashboard_data.json'

def run_flask():
    app.run(debug=True, use_reloader=False)

def start_flask():
    thread = threading.Thread(target=run_flask)
    thread.start()

def plot_visualizations(df):
    try:
        # Convert all columns to numeric, coercing errors to NaN
        df = df.apply(pd.to_numeric, errors='coerce')
        # Visualize the class average
        class_average_path = 'static/class_average_plot.png'
        plt.figure(figsize=(10, 6))
        sns.barplot(x=df.columns[2:-2], y=df.iloc[:, 2:-2].mean())
        plt.title("Class Average")
        plt.xlabel("Subjects")
        plt.ylabel("Average Marks")
        plt.savefig(class_average_path)
        plt.close('all')  # Close all figures

        # Visualize overall percentage distribution
        overall_percentage_path = 'static/overall_percentage_plot.png'
        plt.figure(figsize=(8, 5))
        sns.histplot(df['OverallPercentage'], kde=True)
        plt.title("Overall Percentage Distribution")
        plt.xlabel("Overall Percentage")
        plt.ylabel("Frequency")
        plt.savefig(overall_percentage_path)
        plt.close('all')  # Close all figures

        return class_average_path, overall_percentage_path
    except Exception as e:
        print(f"Error plotting visualizations: {e}")
        return None, None


@app.route('/')
def index():
    dashboard_data = load_dashboard_data()
    return render_template('index.html', dashboard_data=dashboard_data)

@app.route('/upload', methods=['POST'])
def upload():
    try:
        dashboard_data = load_dashboard_data()
        if not dashboard_data:
            return render_template('index.html', message='Please save the dashboard data first')

        file = request.files['file']
        if file.filename == '':
            return render_template('index.html', message='No file selected')

        if file and allowed_file(file.filename):
            df = pd.read_excel(file, engine='openpyxl')
            process_data(df)
            plot_paths = plot_visualizations(df)

            result_file_path = 'static/uploaded_data.xlsx'
            df.to_excel(result_file_path, index=False, engine='openpyxl')

            return render_template(
                'index.html', 
                message='File uploaded successfully!', 
                dashboard_data=dashboard_data, 
                download_link=result_file_path,
                plot_paths=plot_paths,
                show_plots_link=True,
                generate_report_cards_link=True
            )
        else:
            return render_template('index.html', message='Invalid file format. Supported formats: Excel')
    except Exception as e:
        print(f"Error processing file: {e}")
        return render_template('index.html', message=f'Error processing file: {e}', dashboard_data=dashboard_data)


def process_data(df):
    try:
        # Calculate the total for each subject
        df['Total'] = df.iloc[:, 2:].apply(pd.to_numeric, errors='coerce').sum(axis=1)

        # Calculate overall percentage
        total_marks = len(df.columns) - 2  # excluding 'RollNumber' and 'Name' columns
        df['OverallPercentage'] = round((df['Total'] / ((total_marks-1) * 100)) * 100,2)

        # Rank students by overall percentage
        df['Rank'] = df['OverallPercentage'].rank(ascending=False)
    except Exception as e:
        print(f"Error processing data: {e}")


@app.route('/show_plots')
def show_plots():
    try:
        plot_paths = [
            'static/class_average_plot.png',
            'static/overall_percentage_plot.png'
        ]

        return render_template('show_results.html', plot_paths=plot_paths)
    except Exception as e:
        print(f"Error showing plots: {e}")
        return render_template('show_results.html', message='Error showing plots')
    

@app.route('/report_card_template', methods=['POST'])
def report_card_template():
    try:
        # Read the Excel file
        df = pd.read_excel('static/uploaded_data.xlsx', engine='openpyxl')

        # Get the paths of the generated plots
        plot_paths = [
            'static/class_average_plot.png',
            'static/overall_percentage_plot.png'
        ]

        return render_template('report_card_template.html', student_data=df.to_dict(orient='records'), plot_paths=plot_paths)
    except Exception as e:
        print(f"Error rendering report card template: {e}")
        return render_template('index.html', message='Error rendering report card template')


def load_dashboard_data():
    if os.path.exists(DASHBOARD_DATA_FILE):
        with open(DASHBOARD_DATA_FILE, 'r') as f:
            return json.load(f)
    return {}

def save_dashboard_data(data):
    with open(DASHBOARD_DATA_FILE, 'w') as f:
        json.dump(data, f)

@app.route('/save_dashboard', methods=['POST'])
def save_dashboard():
    try:
        school_name = request.form['school_name']
        principal_signature = request.files['principal_signature']
        school_logo = request.files['school_logo']
        class_name = request.form['class_name']
        section = request.form['section']

        dashboard_data = {
            'school_name': school_name,
            'class_name': class_name,
            'section': section
        }

        if principal_signature:
            principal_signature_path = os.path.join('static', 'principal_signature.png')
            principal_signature.save(principal_signature_path)
            dashboard_data['principal_signature_path'] = principal_signature_path

        if school_logo:
            school_logo_path = os.path.join('static', 'school_logo.png')
            school_logo.save(school_logo_path)
            dashboard_data['school_logo_path'] = school_logo_path

        save_dashboard_data(dashboard_data)
        return redirect(url_for('index'))
    except Exception as e:
        print(f"Error saving dashboard data: {e}")
        return render_template('index.html', message='Error saving dashboard data')



@app.route('/generate_pdf', methods=['POST'])
def generate_pdf():
    try:
        # Read the Excel file
        df = pd.read_excel('static/uploaded_data.xlsx', engine='openpyxl')

        dashboard_data = load_dashboard_data()
        student_data = df.to_dict(orient='records')

        pdf_path = generate_pdf_document(student_data, dashboard_data)

        if pdf_path:
            return send_file(pdf_path, as_attachment=True, mimetype='application/pdf', download_name='report_cards.pdf')
        else:
            return render_template('index.html', message='Error generating PDF', dashboard_data=dashboard_data)
    except Exception as e:
        print(f"Error generating PDF: {e}")
        return render_template('index.html', message='Error generating PDF', dashboard_data=dashboard_data)


def generate_pdf_document(student_data, dashboard_data):
    try:
        pdf_file_path = 'static/report_cards.pdf'
        doc = SimpleDocTemplate(pdf_file_path, pagesize=letter,
                                rightMargin=0.5*inch, leftMargin=0.5*inch,
                                topMargin=0.5*inch, bottomMargin=0.5*inch)

        elements = []

        school_logo_path = dashboard_data.get('school_logo_path', '')
        principal_signature_path = dashboard_data.get('principal_signature_path', '')

        styles = getSampleStyleSheet()
        styles.add(ParagraphStyle(name='MyTitle', fontSize=18, leading=22, spaceAfter=12))
        styles.add(ParagraphStyle(name='MyHeading2', fontSize=14, leading=18, spaceAfter=10))
        styles.add(ParagraphStyle(name='MyHeading3', fontSize=12, leading=14, spaceAfter=8))

        for student in student_data:
            if isinstance(student, dict):
                # Header with logo and school information
                header_data = []
                if school_logo_path and os.path.exists(school_logo_path):
                    header_data.append([Image(school_logo_path, width=1*inch, height=1*inch)])
                else:
                    header_data.append([''])

                school_info = f"<b>{dashboard_data.get('school_name', '')}</b><br/>Class: {dashboard_data.get('class_name', '')}<br/>Section: {dashboard_data.get('section', '')}"
                header_data[0].append(Paragraph(school_info, styles['MyHeading2']))

                header_table = Table(header_data, colWidths=[2*inch, 4*inch])
                header_table.setStyle(TableStyle([
                    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                    ('ALIGN', (1, 0), (1, 0), 'CENTER'),
                ]))
                elements.append(header_table)
                elements.append(Spacer(1, 0.2 * inch))

                # Student Information
                elements.append(Paragraph(f"<b>Name:</b> {student.get('Name', '')}", styles['MyHeading3']))
                elements.append(Paragraph(f"<b>Roll Number:</b> {student.get('Roll Number', '')}", styles['MyHeading3']))
                elements.append(Spacer(1, 0.2 * inch))

                # Academic Performance Table
                elements.append(Paragraph("<b>Academic Performance</b>", styles['MyHeading2']))
                data = [['Subject', 'Marks']]
                subjects = student.keys() - {'Name', 'Roll Number', 'OverallPercentage', 'Total', 'Rank'}
                for subject in subjects:
                    data.append([subject, str(student.get(subject, ''))])

                # Total, Overall Percentage, and Rank
                data.append(['Total', str(student.get('Total', ''))])
                data.append(['Overall Percentage', f"{student.get('OverallPercentage', '')}%"])
                data.append(['Rank', str(student.get('Rank', ''))])

                table = Table(data, colWidths=[3 * inch, 2.5 * inch])
                table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ]))
                elements.append(table)
                elements.append(Spacer(1, 0.2 * inch))

                # Extra Curricular Table
                elements.append(Paragraph("<b>Extra Curricular</b>", styles['MyHeading2']))
                extra_curricular_data = [['Activity', 'Grade'],
                                         ['Public Speaking', ''],
                                         ['Dancing', ''],
                                         ['Singing', '']]
                extra_curricular_table = Table(extra_curricular_data, colWidths=[3 * inch, 2.5 * inch])
                extra_curricular_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ]))
                elements.append(extra_curricular_table)
                elements.append(Spacer(1, 0.2 * inch))

                # Remarks Section
                elements.append(Paragraph("<b>Remarks</b>", styles['MyHeading2']))
                remarks_table = Table([[' ' * 80]], colWidths=[6.5 * inch])
                remarks_table.setStyle(TableStyle([
                    ('BOX', (0, 0), (-1, -1), 1, colors.black),
                    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                    ('TOPPADDING', (0, 0), (-1, -1), 12),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
                ]))
                elements.append(remarks_table)
                elements.append(Spacer(1, 0.2 * inch))

                # Principal's Signature
                if principal_signature_path and os.path.exists(principal_signature_path):
                    elements.append(Spacer(1, 0.2 * inch))
                    elements.append(Paragraph("Principal's Signature", styles['MyHeading3']))
                    signature_image = Image(principal_signature_path, width=1.5 * inch, height=0.75 * inch)
                    elements.append(signature_image)

                elements.append(PageBreak())

        doc.build(elements)
        return pdf_file_path
    except Exception as e:
        print(f"Error generating PDF: {e}")
        return None


    

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in {'xlsx'}

if __name__ == '__main__':
    app.run(debug=True, use_reloader=False)