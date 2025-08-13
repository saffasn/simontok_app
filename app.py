from io import BytesIO
import re
import sqlite3
from unicodedata import category
import uuid
import pandas as pd
from werkzeug.utils import secure_filename
import psycopg2
from datetime import datetime
from flask import Flask, render_template, request, redirect, send_file, url_for, session, flash
from uuid import uuid4
from werkzeug.security import generate_password_hash, check_password_hash
import os
import logging
import xlsxwriter
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib.units import cm


# Configure logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

app = Flask(__name__)
app.secret_key = 'your-secret-key-123'
app.config['UPLOAD_FOLDER'] = 'static/uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Valid jenis perwakilan options
VALID_JENIS_PWK = ['KBRI', 'KJRI', 'PTRI', 'KRI', 'KDEI', 'PJB']

# ==============================================
# DATABASE FUNCTIONS
# ==============================================

def get_db_connection():
    try:
        conn = psycopg2.connect(
            host="localhost",
            database="db_simontok",
            user="postgres",
            password="sswatuniS4"
        )
        logger.debug("Database connection successful")
        return conn
    except psycopg2.Error as e:
        logger.error(f'Database connection error: {e}')
        flash(f'Database connection error: {e}', 'error')
        return None

def execute_query(query, params=None, fetch=False, fetch_one=False, commit=False):
    try:
        with get_db_connection() as conn:
            with conn.cursor() as cur:
                cur.execute(query, params or ())
                if commit:
                    conn.commit()
                    return True
                if fetch_one:
                    result = cur.fetchone()
                    return result if result else None
                if fetch:
                    return cur.fetchall()
                return True
    except Exception as e:
        logger.error(f"Database error: {str(e)}")
        if commit:
            conn.rollback()
        raise e

def get_next_urutan():
    """Get the next auto-increment value for no_urutan"""
    result = execute_query(
        "SELECT COALESCE(MAX(NO_URUTAN), 0) + 1 FROM REF_PERWAKILAN",
        fetch_one=True
    )
    return result[0] if result else 1

def generate_next_kategori_id():
    try:
        # Get the last ID from the database
        last_id = execute_query(
            "SELECT id FROM ref_kategori_sistem ORDER BY id DESC LIMIT 1",
            fetch_one=True
        )
        
        if last_id and last_id[0]:  # Check if result exists and has at least one element
            current_id = last_id[0]
            # Extract numeric part if format is K0001
            if current_id.startswith('K'):
                try:
                    num = int(current_id[1:]) + 1
                    return f"K{num:04d}"
                except ValueError:
                    return "K0001"
            return "K0001"  # Fallback if format is unexpected
        return "K0001"  # First record
    except Exception as e:
        logger.error(f"Error generating ID: {str(e)}")
        return "K0001"  # Fallback on error

def get_next_no_perwakilan():
    """Get the next auto-increment value for no_perwakilan"""
    result = execute_query(
        "SELECT COALESCE(MAX(NO_PERWAKILAN), 0) + 1 FROM REF_PERWAKILAN",
        fetch_one=True
    )
    return result[0] if result else 1

def generate_tipe_palsan_id():
    """Generate ID Tipe Palsan otomatis (P0001, P0002, dst)"""
    last_id = execute_query(
        "SELECT id_tipe FROM tipe_palsan ORDER BY id_tipe DESC LIMIT 1",
        fetch_one=True
    )
    
    if last_id and last_id[0]:
        last_num = int(last_id[0][1:])
        return f"P{last_num + 1:04d}"
    return "P0001"

def generate_palsan_id():
    """Generate ID Palsan otomatis (PL001, PL002, dst)"""
    last_id = execute_query(
        "SELECT id_palsan FROM tabel_palsan ORDER BY id_palsan DESC LIMIT 1",
        fetch_one=True
    )
    
    if last_id and last_id[0]:
        last_num = int(last_id[0][2:])  # Ambil angka setelah 'PL'
        return f"PL{last_num + 1:03d}"
    return "PL001"

def generate_distribution_pdf(data):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, 
                          leftMargin=2*cm, rightMargin=2*cm,
                          topMargin=1.5*cm, bottomMargin=1.5*cm)
    
    styles = getSampleStyleSheet()
    elements = []
    
    # Custom styles
    header_style = ParagraphStyle(
        'Header',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=12,
        alignment=1,  # center
        spaceAfter=6
    )
   
    subheader_style = ParagraphStyle(
        'Header',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=12,
        alignment=1,  # center
        spaceAfter=6
    )
    
    normal_style = ParagraphStyle(
        'Normal',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=10,
        leading=12,
        spaceAfter=6
    )
    
    # Create a custom style with specific underline properties
    underlined_header_style = ParagraphStyle(
        'UnderlinedHeader',
        fontName='Helvetica-Bold',
        fontSize=14,
        alignment=1,  # center alignment
        spaceAfter=6,
        textDecoration='underline',  # This enables the underline
        underlineWidth=1.5,         # Line thickness
        underlineOffset=-4,         # Position adjustment (negative moves it down)
        underlineGap=2,             # Space between text and underline
        underlineColor=colors.black # Line color
    )
    
    
    # Header section
    elements.append(Paragraph("KEMENTERIAN LUAR NEGERI REPUBLIK INDONESIA", header_style))
    elements.append(Paragraph("PUSAT TEKNOLOGI INFORMASI DAN KOMUNIKASI", header_style))
    elements.append(Paragraph("KEMENTERIAN DAN PERWAKILAN", header_style))
    elements.append(Spacer(1, 12))
    
    # Document title
    # Create a table with underline effect
    t = Table([["TANDA TERIMA"]], style=[
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,-1), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,-1), 14),
        ('LINEBELOW', (0,0), (-1,-1), 1, colors.black),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ])
    elements.append(t)
    elements.append(Paragraph(f"NOMOR: {data['id_palsan']}/PUSTIKKP/{datetime.now().strftime('%m/%Y')}", subheader_style))
    elements.append(Spacer(1, 12))
    
    # Body text
    body_text = Paragraph(
        "Bersama ini diserahterimakan kelengkapan yang akan digunakan pada perangkat pengamanan dokumen " +
        "Kementerian Luar Negeri dan Perwakilan RI, dengan data-data sebagai berikut:", 
        normal_style
    )
    elements.append(body_text)
    elements.append(Spacer(1, 12))
    
    # Items table - only showing the actually input item
    items_data = [
        ["No", "Jenis", "Merk/Model", "Jumlah (unit)", "Serial Number"],
        ["1.", data['tipe_palsan'], "-", "1", data['serial_number']]
    ]
    
    items_table = Table(items_data, colWidths=[1.5*cm, 4*cm, 4*cm, 3*cm, 4*cm])
    items_table.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTNAME', (0,1), (-1,-1), 'Helvetica'),
        ('FONTSIZE', (0,0), (-1,-1), 10),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('GRID', (0,0), (-1,-1), 1, colors.black),
        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
    ]))
    elements.append(items_table)
    elements.append(Spacer(1, 12))
    
    # Footer notes
    notes = Paragraph(
        "Tanda terima ini dibuat 2 (dua) rangkap untuk diserahkan kepada:<br/>" +
        "1. Kepala Bidang Pengembangan TIK<br/>" +
        "2. Yang bersangkutan<br/><br/>" +
        "Mohon tanda terima ini dapat dikembalikan kepada Pokja Persandian Diplomatik dan softcopy dapat " +
        "dikirim ke email vpn sus.pustkktp@vpn.kemlu.go.id.",
        normal_style
    )
    elements.append(notes)
    elements.append(Spacer(1, 24))
    
    
    # Create a table with 3 columns (left signature, spacer, right signature)
    signatures_data = [
        ["","",Paragraph(f"Jakarta, {datetime.now().strftime('%d %B %Y')}", normal_style)],
        ["Yang Menerima", "", "Yang Menyerahkan"],
        ["", "", ""],  # Empty row for spacing
        ["PK Satker BPO", "", Paragraph("Staff Pokja Persandian Diplomatik", styles['Normal'])],
        ["", "", ""],  # Empty row for spacing
        ["", "", Table([[""]], style=[
        ], colWidths=[5*cm])],  # Signature line
        [data['nama_peminjam'], "", data['penyerah']],
        [Paragraph(f"NIP: {data['nip_peminjam']}", styles['Normal']), "", 
        Paragraph(f"NIP: {data['nip_penyerah']}", styles['Normal'])]
    ]
    
    signatures_table = Table(signatures_data, colWidths=[7*cm, 2*cm, 7*cm])
    signatures_table.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),
        ('FONTSIZE', (0,0), (-1,-1), 10),
        ('ALIGN', (1,0), (0,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('SPAN', (1,0), (1,-1)),  # Merge the spacer column
    ]))
    
    elements.append(signatures_table)
    
    doc.build(elements)
    buffer.seek(0)
    return buffer

# ==============================================
# AUTHENTICATION ROUTES
# ==============================================

@app.route('/')
def home():
    return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username').strip()
        password = request.form.get('password').strip()

        user = execute_query(
            "SELECT id_pengguna, username, password, role, id_pwk FROM tabel_pengguna WHERE username = %s",
            (username,),
            fetch_one=True
        )

        if user and check_password_hash(user[2], password):
            session['user_id'] = user[0]
            session['username'] = user[1]
            session['role'] = user[3]
            session['trigram'] = user[4].upper() if user[4] else None  # Ensure uppercase
            flash('Login berhasil', 'success')
            return redirect(url_for('dashboard'))
        else:
            flash('Username atau password salah!', 'error')
    
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    flash('Anda telah logout', 'info')
    return redirect(url_for('login'))

# ==============================================
# DASHBOARD
# ==============================================

@app.route('/dashboard')
def dashboard():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    try:
        # Data untuk dashboard
        stats = {
            'users': execute_query("SELECT COUNT(*) FROM TABEL_PENGGUNA", fetch_one=True)[0] or 0,
            'perwakilan': execute_query("SELECT COUNT(*) FROM REF_PERWAKILAN", fetch_one=True)[0] or 0,
            'jenis_sistem': execute_query("SELECT COUNT(*) FROM REF_JENIS_SISTEM", fetch_one=True)[0] or 0,
            'sistem': execute_query("SELECT COUNT(*) FROM TABEL_SISTEM", fetch_one=True)[0] or 0
        }
        
        # Sistem terbaru (limit 5 untuk dashboard)
        recent_systems = execute_query("""
            SELECT s.ID_SISTEM, s.TAHUN, s.ID_JENIS, s.NO_SISTEM, 
                   s.NAMA_SISTEM, s.JML_LEMBAR, s.STATUS,
                   j.JENIS, p.NAMA_PERWAKILAN
            FROM TABEL_SISTEM s
            LEFT JOIN REF_JENIS_SISTEM j ON s.ID_JENIS = j.ID_JENIS
            LEFT JOIN REF_PERWAKILAN p ON j.TRIGRAM_PWK = p.TRIGRAM
            ORDER BY s.DATE_INPUT DESC
            LIMIT 5
        """, fetch=True) or []
        
        return render_template('dashboard.html', 
                            stats=stats, 
                            recent_systems=recent_systems,
                            is_dashboard=True)
    except Exception as e:
        logger.error(f"Dashboard error: {str(e)}")
        flash('Terjadi kesalahan saat memuat dashboard', 'error')
        return render_template('dashboard.html', stats={}, recent_systems=[], is_dashboard=True)
    
# ==============================================
# PENGGUNA CRUD ROUTES
# ==============================================

@app.route('/pengguna/export/pdf')
def export_pengguna_pdf():
    if 'user_id' not in session or session.get('role') != 0:  # Admin only
        return redirect(url_for('login'))
    
    try:
        # Get all parameters from the request
        search = request.args.get('search', '').strip()
        sort_column = request.args.get('sort', 'id_pengguna')
        sort_direction = request.args.get('dir', 'asc')

        # Validate sort column
        valid_columns = {
            'id_pengguna': 'id_pengguna',
            'nama_pengguna': 'nama_pengguna',
            'username': 'username',
            'role': 'role',
            'id_pwk': 'id_pwk',
            'date_input': 'date_input'
        }
        sort_column = valid_columns.get(sort_column, 'id_pengguna')
        sort_direction = 'DESC' if sort_direction.lower() == 'desc' else 'ASC'

        # Base query with join to get perwakilan name
        query = f"""
            SELECT p.id_pengguna, p.nama_pengguna, p.username, p.role, 
                   p.id_pwk, r.NAMA_PERWAKILAN,
                   p.date_input
            FROM tabel_pengguna p
            LEFT JOIN ref_perwakilan r ON p.id_pwk = r.TRIGRAM
            WHERE p.nama_pengguna ILIKE %s OR 
                  p.username ILIKE %s OR
                  r.NAMA_PERWAKILAN ILIKE %s OR
                  p.id_pengguna ILIKE %s
            ORDER BY {sort_column} {sort_direction}
        """

        # Parameters for the query
        search_param = f'%{search}%'
        params = [search_param for _ in range(4)]

        # Execute query to get all data
        pengguna_list = execute_query(query, params, fetch=True) or []

        # Create PDF with A4 size (portrait orientation)
        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, 
                              pagesize=A4,
                              leftMargin=15,
                              rightMargin=15,
                              topMargin=20,
                              bottomMargin=20)
        elements = []
        
        # Styles configuration
        styles = getSampleStyleSheet()
        
        # Title style
        title_style = ParagraphStyle(
            'Title',
            parent=styles['Title'],
            fontSize=14,
            alignment=TA_CENTER,
            spaceAfter=12
        )
        
        # Metadata style
        meta_style = ParagraphStyle(
            'Meta',
            parent=styles['Normal'],
            fontSize=9,
            spaceAfter=6
        )
        
        # Add title
        elements.append(Paragraph("LAPORAN DATA PENGGUNA", title_style))
        
        # Add metadata
        metadata = [
            f"Tanggal Export: {datetime.now().strftime('%d-%m-%Y %H:%M')}"
        ]
        if search:
            metadata.insert(0, f"Filter: {search}")
        
        for meta in metadata:
            elements.append(Paragraph(meta, meta_style))
        
        # Prepare table data
        data = []
        headers = [
            "No",
            "ID Pengguna",
            "Nama Pengguna",
            "Username",
            "Role",
            "Perwakilan",
            "Tanggal Input"
        ]
        data.append(headers)
        
        for idx, item in enumerate(pengguna_list, 1):
            row = [
                str(idx),
                item[0] or '-',  # ID Pengguna
                item[1] or '-',  # Nama Pengguna
                item[2] or '-',  # Username
                'Admin' if item[3] == 0 else 'User',  # Role
                item[5] or '-',  # Perwakilan
                item[6].strftime('%d-%m-%Y') if item[6] else '-'  # Tanggal Input
            ]
            data.append(row)
        
        # Calculate available width using A4 width
        available_width = A4[0] - doc.leftMargin - doc.rightMargin
        
        # Define column distribution
        col_ratios = [0.05, 0.1, 0.2, 0.15, 0.1, 0.25, 0.15]
        col_widths = [available_width * ratio for ratio in col_ratios]
        
        # Ensure minimum widths
        min_widths = [25, 50, 70, 60, 40, 80, 60]
        col_widths = [max(w, min_w) for w, min_w in zip(col_widths, min_widths)]
        
        # Create table with adjusted column widths
        table = Table(data, 
                     colWidths=col_widths,
                     repeatRows=1,
                     hAlign='LEFT')
        
        # Table style with adjusted font size and padding
        style = TableStyle([
            # Header style
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#4472C4')),
            ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,0), 7),
            ('ALIGN', (0,0), (-1,0), 'CENTER'),
            ('BOTTOMPADDING', (0,0), (-1,0), 4),
            
            # Cell style
            ('FONTSIZE', (0,1), (-1,-1), 7),
            ('LEADING', (0,0), (-1,-1), 8),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey),
            ('WORDWRAP', (0,0), (-1,-1), True),
            
            # Alternating row colors
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.whitesmoke]),
            
            # Specific column alignments
            ('ALIGN', (0,0), (0,-1), 'CENTER'),
            ('ALIGN', (1,0), (1,-1), 'LEFT'),
            ('ALIGN', (3,0), (3,-1), 'LEFT'),
            ('ALIGN', (4,0), (4,-1), 'CENTER'),
            ('ALIGN', (-1,0), (-1,-1), 'CENTER')
        ])
        
        table.setStyle(style)
        elements.append(table)
        
        # Build PDF
        doc.build(elements)
        buffer.seek(0)
        
        return send_file(
            buffer,
            as_attachment=True,
            download_name=f"data_pengguna_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
            mimetype='application/pdf'
        )
        
    except Exception as e:
        print(f"PDF Export Error: {str(e)}")
        flash('Gagal membuat PDF: ' + str(e), 'error')
        return redirect(url_for('list_pengguna'))
    
@app.route('/pengguna/export/excel')
def export_pengguna_excel():
    if 'user_id' not in session or session.get('role') != 0:  # Admin only
        return redirect(url_for('login'))
    
    try:
        # Get all parameters from the request
        search = request.args.get('search', '').strip()
        sort_column = request.args.get('sort', 'id_pengguna')
        sort_direction = request.args.get('dir', 'asc')

        # Validate sort column
        valid_columns = {
            'id_pengguna': 'id_pengguna',
            'nama_pengguna': 'nama_pengguna',
            'username': 'username',
            'role': 'role',
            'id_pwk': 'id_pwk',
            'date_input': 'date_input'
        }
        sort_column = valid_columns.get(sort_column, 'id_pengguna')
        sort_direction = 'DESC' if sort_direction.lower() == 'desc' else 'ASC'

        # Base query with join to get perwakilan name
        query = f"""
            SELECT p.id_pengguna, p.nama_pengguna, p.username, p.role, 
                   p.id_pwk, r.NAMA_PERWAKILAN,
                   p.date_input
            FROM tabel_pengguna p
            LEFT JOIN ref_perwakilan r ON p.id_pwk = r.TRIGRAM
            WHERE p.nama_pengguna ILIKE %s OR 
                  p.username ILIKE %s OR
                  r.NAMA_PERWAKILAN ILIKE %s OR
                  p.id_pengguna ILIKE %s
            ORDER BY {sort_column} {sort_direction}
        """

        # Parameters for the query
        search_param = f'%{search}%'
        params = [search_param for _ in range(4)]

        # Execute query to get all data
        pengguna_list = execute_query(query, params, fetch=True) or []

        # Create Excel file in memory
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet("Data Pengguna")
        
        # Add formats
        title_format = workbook.add_format({
            'bold': True,
            'font_size': 14,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        subtitle_format = workbook.add_format({
            'font_size': 10,
            'align': 'left',
            'valign': 'vcenter'
        })
        
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#4472C4',
            'font_color': '#FFFFFF',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True
        })
        
        # Data formats
        data_format = workbook.add_format({
            'border': 1,
            'text_wrap': True,
            'valign': 'top'
        })
        
        center_format = workbook.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        # Write title and metadata
        worksheet.merge_range('A1:G1', 'LAPORAN DATA PENGGUNA', title_format)
        
        current_row = 1
        if search:
            worksheet.write(current_row, 0, f"Filter: {search}", subtitle_format)
            current_row += 1
            
        worksheet.write(current_row, 0, f"Tanggal Export: {datetime.now().strftime('%d-%m-%Y %H:%M')}", subtitle_format)
        current_row += 2
        
        # Write headers
        headers = [
            "No",
            "ID Pengguna",
            "Nama Pengguna",
            "Username",
            "Role",
            "Perwakilan",
            "Tanggal Input"
        ]
        
        for col, header in enumerate(headers):
            worksheet.write(current_row, col, header, header_format)
        
        current_row += 1
        
        # Write data
        for idx, item in enumerate(pengguna_list, 1):
            worksheet.write(current_row, 0, idx, center_format)  # No
            worksheet.write(current_row, 1, item[0] or '-', data_format)  # ID Pengguna
            worksheet.write(current_row, 2, item[1] or '-', data_format)  # Nama Pengguna
            worksheet.write(current_row, 3, item[2] or '-', data_format)  # Username
            worksheet.write(current_row, 4, 'Admin' if item[3] == 0 else 'User', center_format)  # Role
            worksheet.write(current_row, 5, item[5] or '-', data_format)  # Perwakilan
            worksheet.write(current_row, 6, item[6].strftime('%d-%m-%Y') if item[6] else '-', center_format)  # Tanggal Input
            
            current_row += 1
        
        # Auto-adjust column widths based on content
        col_widths = [
            ("No", 5),
            ("ID Pengguna", 15),
            ("Nama Pengguna", 25),
            ("Username", 20),
            ("Role", 12),
            ("Perwakilan", 30),
            ("Tanggal Input", 15)
        ]
        
        for i, (header, default_width) in enumerate(col_widths):
            max_length = len(header)
            for row in pengguna_list:
                if i == 0:  # No column
                    val_length = len(str(idx))
                elif i == 1:  # ID Pengguna
                    val_length = len(str(row[0] or ''))
                elif i == 2:  # Nama Pengguna
                    val_length = len(str(row[1] or ''))
                elif i == 3:  # Username
                    val_length = len(str(row[2] or ''))
                elif i == 4:  # Role
                    val_length = len('Admin' if row[3] == 0 else 'User')
                elif i == 5:  # Perwakilan
                    val_length = len(str(row[5] or ''))
                elif i == 6:  # Tanggal Input
                    val_length = len(row[6].strftime('%d-%m-%Y') if row[6] else '-')
                
                if val_length > max_length:
                    max_length = val_length
            
            # Set column width with some padding
            worksheet.set_column(i, i, min(max_length + 2, 50))
        
        workbook.close()
        output.seek(0)
        
        return send_file(
            output,
            as_attachment=True,
            download_name=f"data_pengguna_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        print(f"Excel Export Error: {str(e)}")
        flash('Gagal membuat Excel: ' + str(e), 'error')
        return redirect(url_for('list_pengguna'))

@app.route('/pengguna')
def list_pengguna():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    logger.debug(f"Current user role: {session.get('role')}")
    
    # Only admin can access pengguna management
    if session.get('role') != 0:  # Assuming 0 is admin role
        flash('Anda tidak memiliki akses ke halaman ini', 'error')
        return redirect(url_for('dashboard'))
    
    # Get search and pagination parameters
    search = request.args.get('search', '').strip()
    page = request.args.get('page', 1, type=int)
    per_page = 20
    sort_column = request.args.get('sort', 'id_pengguna')
    sort_direction = request.args.get('dir', 'asc')

    # Validate sort column
    valid_columns = {
        'id_pengguna': 'id_pengguna',
        'nama_pengguna': 'nama_pengguna',
        'username': 'username',
        'role': 'role',
        'id_pwk': 'id_pwk',
        'date_input': 'date_input'
    }
    sort_column = valid_columns.get(sort_column, 'id_pengguna')
    
    # Validate sort direction
    sort_direction = 'DESC' if sort_direction.lower() == 'desc' else 'ASC'

    # Base query with join to get perwakilan name
    query = f"""
        SELECT p.id_pengguna, p.nama_pengguna, p.username, p.role, 
               p.id_pwk, r.NAMA_PERWAKILAN,
               p.user_input, p.date_input, p.user_update, p.date_update
        FROM tabel_pengguna p
        LEFT JOIN ref_perwakilan r ON p.id_pwk = r.TRIGRAM
        WHERE p.nama_pengguna ILIKE %s OR 
              p.username ILIKE %s OR
              r.NAMA_PERWAKILAN ILIKE %s OR
              p.id_pengguna ILIKE %s
        ORDER BY {sort_column} {sort_direction}
    """

    # Count query for pagination
    count_query = """
        SELECT COUNT(*)
        FROM tabel_pengguna p
        LEFT JOIN ref_perwakilan r ON p.id_pwk = r.TRIGRAM
        WHERE p.nama_pengguna ILIKE %s OR 
              p.username ILIKE %s OR
              r.NAMA_PERWAKILAN ILIKE %s OR
              p.id_pengguna ILIKE %s
    """

    search_param = f'%{search}%'
    
    # Get total count
    total = execute_query(count_query, 
                         (search_param, search_param, search_param, search_param), 
                         fetch_one=True)[0] or 0
    
    # Add pagination to main query
    paginated_query = query + f" LIMIT {per_page} OFFSET {(page - 1) * per_page}"
    
    # Execute query
    pengguna_list = execute_query(paginated_query, 
                                (search_param, search_param, search_param, search_param), 
                                fetch=True) or []
    
    total_pages = (total + per_page - 1) // per_page
    
    # Get list of perwakilan for filter
    perwakilan_list = execute_query(
        "SELECT TRIGRAM, NAMA_PERWAKILAN FROM ref_perwakilan ORDER BY NAMA_PERWAKILAN",
        fetch=True
    ) or []
    
    return render_template('pengguna/list.html', 
                         pengguna_list=pengguna_list,
                         perwakilan_list=perwakilan_list,
                         search=search,
                         page=page,
                         per_page=per_page,
                         total=total,
                         total_pages=total_pages,
                         sort_column=sort_column,
                         sort_direction=sort_direction)

@app.route('/pengguna/create', methods=['GET', 'POST'])
def create_pengguna():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    if session.get('role') != 0:
        flash('Anda tidak memiliki akses ke halaman ini', 'error')
        return redirect(url_for('dashboard'))
    
    perwakilan_list = execute_query(
        "SELECT TRIGRAM, NAMA_PERWAKILAN FROM ref_perwakilan ORDER BY NAMA_PERWAKILAN",
        fetch=True
    ) or []
    
    if request.method == 'POST':
        try:
            # Generate user ID (U0001 format)
            max_id_result = execute_query(
                "SELECT MAX(CAST(SUBSTRING(id_pengguna, 2) AS INTEGER)) FROM tabel_pengguna",
                fetch_one=True
            )
            max_id = max_id_result[0] if max_id_result and max_id_result[0] else 0
            user_id = f'U{max_id + 1:04}'
            
            # Get form data
            nama_pengguna = request.form.get('nama_pengguna', '').strip()
            username = request.form.get('username', '').strip()
            password = request.form.get('password', '').strip()
            role = request.form.get('role', '1').strip()  # Default to 1 (regular user)
            id_pwk = request.form.get('id_pwk', '').strip() or None
            
            # Validate required fields
            if not all([nama_pengguna, username, password]):
                flash('Nama Pengguna, Username, dan Password wajib diisi', 'error')
                return render_template('pengguna/create.html', 
                                     perwakilan_list=perwakilan_list,
                                     form_data=request.form)
            
            # Check if username already exists
            username_exists = execute_query(
                "SELECT 1 FROM tabel_pengguna WHERE username = %s",
                (username,),
                fetch_one=True
            )
            if username_exists:
                flash('Username sudah digunakan', 'error')
                return render_template('pengguna/create.html', 
                                     perwakilan_list=perwakilan_list,
                                     form_data=request.form)
            
            hashed_password = generate_password_hash(password, method='pbkdf2:sha256')
            
            # Prepare data
            data = (
                user_id,
                nama_pengguna,
                username,
                hashed_password,
                int(role),
                id_pwk,
                session.get('username', 'system'),
                datetime.now()
            )
            
            success = execute_query("""
                INSERT INTO tabel_pengguna 
                (id_pengguna, nama_pengguna, username, password, role, id_pwk, user_input, date_input)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
            """, data, commit=True)
            
            if success:
                flash('Pengguna berhasil ditambahkan', 'success')
                return redirect(url_for('list_pengguna'))
            else:
                flash('Gagal menambahkan pengguna: Tidak ada baris yang terpengaruh', 'error')
        except Exception as e:
            logger.error(f"Error creating pengguna: {str(e)}", exc_info=True)  # Add exc_info for full traceback
            flash(f'Terjadi kesalahan spesifik: {str(e)}', 'error')  # Show specific error
        except psycopg2.Error as e:  # Add specific database error handling
            logger.error(f"Database error creating pengguna: {str(e)}", exc_info=True)
            flash(f'Database error: {str(e)}', 'error')
    
    return render_template('pengguna/create.html', 
                         perwakilan_list=perwakilan_list,
                         role_options=[(0, 'Admin'), (1, 'User')])

@app.route('/pengguna/edit/<user_id>', methods=['GET', 'POST'])
def edit_pengguna(user_id):
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    if session.get('role') != 0:
        flash('Anda tidak memiliki akses ke halaman ini', 'error')
        return redirect(url_for('dashboard'))
    
    pengguna = execute_query(
        """SELECT id_pengguna, nama_pengguna, username, password, role, id_pwk 
           FROM tabel_pengguna WHERE id_pengguna = %s""",
        (user_id,),
        fetch_one=True
    )
    
    if not pengguna:
        flash('Data pengguna tidak ditemukan', 'error')
        return redirect(url_for('list_pengguna'))
    
    perwakilan_list = execute_query(
        "SELECT TRIGRAM, NAMA_PERWAKILAN FROM ref_perwakilan ORDER BY NAMA_PERWAKILAN",
        fetch=True
    ) or []
    
    if request.method == 'POST':
        try:
            nama_pengguna = request.form.get('nama_pengguna', '').strip()
            username = request.form.get('username', '').strip()
            password = request.form.get('password', '').strip()
            role = request.form.get('role', '1').strip()
            id_pwk = request.form.get('id_pwk', '').strip() or None
            
            if not all([nama_pengguna, username]):
                flash('Nama Pengguna dan Username wajib diisi', 'error')
                return redirect(url_for('edit_pengguna', user_id=user_id))
            
            username_exists = execute_query(
                "SELECT 1 FROM tabel_pengguna WHERE username = %s AND id_pengguna != %s",
                (username, user_id),
                fetch_one=True
            )
            if username_exists:
                flash('Username sudah digunakan', 'error')
                return redirect(url_for('edit_pengguna', user_id=user_id))
            
            # Prepare update data
            update_data = [
                nama_pengguna,
                username,
                int(role),
                id_pwk,
                session.get('username', 'system'),
                datetime.now(),
                user_id
            ]
            
            # If password is provided, update it
            if password:
                hashed_password = generate_password_hash(password, method='pbkdf2:sha256')
                update_query = """
                    UPDATE tabel_pengguna SET
                        nama_pengguna = %s,
                        username = %s,
                        password = %s,
                        role = %s,
                        id_pwk = %s,
                        user_update = %s,
                        date_update = %s
                    WHERE id_pengguna = %s
                """
                update_data.insert(2, hashed_password)
            else:
                update_query = """
                    UPDATE tabel_pengguna SET
                        nama_pengguna = %s,
                        username = %s,
                        role = %s,
                        id_pwk = %s,
                        user_update = %s,
                        date_update = %s
                    WHERE id_pengguna = %s
                """
            
            success = execute_query(update_query, tuple(update_data), commit=True)
            
            if success:
                flash('Pengguna berhasil diperbarui', 'success')
                return redirect(url_for('list_pengguna'))
            else:
                flash('Gagal memperbarui pengguna', 'error')
                
        except Exception as e:
            logger.error(f"Error updating pengguna: {str(e)}")
            flash(f'Terjadi kesalahan: {str(e)}', 'error')
    
    return render_template('pengguna/edit.html', 
                         pengguna=pengguna,
                         perwakilan_list=perwakilan_list,
                         role_options=[(0, 'Admin'), (1, 'User')])

@app.route('/pengguna/delete/<user_id>', methods=['POST'])
def delete_pengguna(user_id):
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    # Only admin can delete pengguna
    if session.get('role') != 0:
        flash('Anda tidak memiliki akses ke halaman ini', 'error')
        return redirect(url_for('dashboard'))
    
    # Prevent self-deletion
    if user_id == session.get('user_id'):
        flash('Anda tidak dapat menghapus akun sendiri', 'error')
        return redirect(url_for('list_pengguna'))
    
    try:
        success = execute_query(
            "DELETE FROM tabel_pengguna WHERE id_pengguna = %s",
            (user_id,),
            commit=True
        )
        
        if success:
            flash('Pengguna berhasil dihapus', 'success')
        else:
            flash('Gagal menghapus pengguna', 'error')
    except Exception as e:
        logger.error(f"Error deleting pengguna: {str(e)}")
        flash('Terjadi kesalahan saat menghapus data', 'error')
    
    return redirect(url_for('list_pengguna'))
    

# ==============================================
# PERWAKILAN CRUD (IMPROVED VERSION)
# ==============================================

@app.route('/perwakilan/export/pdf')
def export_perwakilan_pdf():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    try:
        # Get all parameters from the request
        search = request.args.get('search', '').strip()
        sort_column = request.args.get('sort', 'NO_URUTAN')
        sort_direction = request.args.get('dir', 'asc')

        # Validate sort column
        valid_columns = {
            'TRIGRAM': 'TRIGRAM',
            'BIGRAM': 'BIGRAM',
            'NAMA_PERWAKILAN': 'NAMA_PERWAKILAN',
            'NEGARA': 'NEGARA',
            'JENIS_PWK': 'JENIS_PWK',
            'NO_URUTAN': 'NO_URUTAN'
        }
        sort_column = valid_columns.get(sort_column, 'NO_URUTAN')
        sort_direction = 'DESC' if sort_direction.lower() == 'desc' else 'ASC'

        # Base query
        query = """
            SELECT TRIGRAM, BIGRAM, NAMA_PERWAKILAN, NEGARA, JENIS_PWK, NO_URUTAN
            FROM REF_PERWAKILAN 
            WHERE NAMA_PERWAKILAN ILIKE %s OR 
                  NEGARA ILIKE %s OR 
                  TRIGRAM ILIKE %s OR
                  BIGRAM ILIKE %s OR
                  JENIS_PWK ILIKE %s
            ORDER BY {} {}
        """.format(sort_column, sort_direction)
        
        # Parameters for the query
        search_param = f'%{search}%'
        params = [search_param for _ in range(5)]

        # Execute query to get all data
        perwakilan_list = execute_query(query, params, fetch=True) or []

        # Create PDF with A4 size (portrait orientation)
        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, 
                              pagesize=A4,
                              leftMargin=15,
                              rightMargin=15,
                              topMargin=20,
                              bottomMargin=20)
        elements = []
        
        # Styles configuration
        styles = getSampleStyleSheet()
        
        # Title style
        title_style = ParagraphStyle(
            'Title',
            parent=styles['Title'],
            fontSize=14,
            alignment=TA_CENTER,
            spaceAfter=12
        )
        
        # Metadata style
        meta_style = ParagraphStyle(
            'Meta',
            parent=styles['Normal'],
            fontSize=9,
            spaceAfter=6
        )
        
        # Add title
        elements.append(Paragraph("LAPORAN DATA PERWAKILAN", title_style))
        
        # Add metadata
        metadata = [
            f"Tanggal Export: {datetime.now().strftime('%d-%m-%Y %H:%M')}"
        ]
        if search:
            metadata.insert(0, f"Filter: {search}")
        
        for meta in metadata:
            elements.append(Paragraph(meta, meta_style))
        
        # Prepare table data
        data = []
        headers = [
            "No",
            "Trigram",
            "Bigram",
            "Nama Perwakilan",
            "Negara",
            "Jenis PWK",
        ]
        data.append(headers)
        
        for idx, item in enumerate(perwakilan_list, 1):
            row = [
                str(idx),
                item[0] or '-',  # Trigram
                item[1] or '-',  # Bigram
                item[2] or '-',  # Nama Perwakilan
                item[3] or '-',  # Negara
                item[4] or '-',  # Jenis PWK
            ]
            data.append(row)
        
        # Calculate available width using A4 width
        available_width = A4[0] - doc.leftMargin - doc.rightMargin
        
        # Define column distribution
        col_ratios = [0.05, 0.1, 0.1, 0.25, 0.4, 0.02]
        col_widths = [available_width * ratio for ratio in col_ratios]
        
        # Ensure minimum widths
        min_widths = [25, 40, 40, 70, 120, 50]
        col_widths = [max(w, min_w) for w, min_w in zip(col_widths, min_widths)]
        
        # Create table with adjusted column widths
        table = Table(data, 
                     colWidths=col_widths,
                     repeatRows=1,
                     hAlign='LEFT')
        
        # Table style with adjusted font size and padding
        style = TableStyle([
            # Header style
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#4472C4')),
            ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,0), 7),
            ('ALIGN', (0,0), (-1,0), 'CENTER'),
            ('BOTTOMPADDING', (0,0), (-1,0), 4),
            
            # Cell style
            ('FONTSIZE', (0,1), (-1,-1), 7),
            ('LEADING', (0,0), (-1,-1), 8),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey),
            ('WORDWRAP', (0,0), (-1,-1), True),
            
            # Alternating row colors
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.whitesmoke]),
            
            # Specific column alignments
            ('ALIGN', (0,0), (0,-1), 'CENTER'),
            ('ALIGN', (1,0), (2,-1), 'CENTER'),
            ('ALIGN', (-1,0), (-1,-1), 'CENTER')
        ])
        
        table.setStyle(style)
        elements.append(table)
        
        # Build PDF
        doc.build(elements)
        buffer.seek(0)
        
        return send_file(
            buffer,
            as_attachment=True,
            download_name=f"data_perwakilan_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
            mimetype='application/pdf'
        )
        
    except Exception as e:
        print(f"PDF Export Error: {str(e)}")
        flash('Gagal membuat PDF: ' + str(e), 'error')
        return redirect(url_for('list_perwakilan'))

@app.route('/perwakilan/export/excel')
def export_perwakilan_excel():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    try:
        # Get all parameters from the request
        search = request.args.get('search', '').strip()
        sort_column = request.args.get('sort', 'NO_URUTAN')
        sort_direction = request.args.get('dir', 'asc')

        # Validate sort column
        valid_columns = {
            'TRIGRAM': 'TRIGRAM',
            'BIGRAM': 'BIGRAM',
            'NAMA_PERWAKILAN': 'NAMA_PERWAKILAN',
            'NEGARA': 'NEGARA',
            'JENIS_PWK': 'JENIS_PWK',
            'NO_URUTAN': 'NO_URUTAN'
        }
        sort_column = valid_columns.get(sort_column, 'NO_URUTAN')
        sort_direction = 'DESC' if sort_direction.lower() == 'desc' else 'ASC'

        # Base query
        query = """
            SELECT TRIGRAM, BIGRAM, NAMA_PERWAKILAN, NEGARA, JENIS_PWK, NO_URUTAN
            FROM REF_PERWAKILAN 
            WHERE NAMA_PERWAKILAN ILIKE %s OR 
                  NEGARA ILIKE %s OR 
                  TRIGRAM ILIKE %s OR
                  BIGRAM ILIKE %s OR
                  JENIS_PWK ILIKE %s
            ORDER BY {} {}
        """.format(sort_column, sort_direction)
        
        # Parameters for the query
        search_param = f'%{search}%'
        params = [search_param for _ in range(5)]

        # Execute query to get all data
        perwakilan_list = execute_query(query, params, fetch=True) or []

        # Create Excel file in memory
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet("Data Perwakilan")
        
        # Add formats
        title_format = workbook.add_format({
            'bold': True,
            'font_size': 14,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        subtitle_format = workbook.add_format({
            'font_size': 10,
            'align': 'left',
            'valign': 'vcenter'
        })
        
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#4472C4',
            'font_color': '#FFFFFF',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True
        })
        
        # Data formats
        data_format = workbook.add_format({
            'border': 1,
            'text_wrap': True,
            'valign': 'top'
        })
        
        center_format = workbook.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        # Write title and metadata
        worksheet.merge_range('A1:G1', 'LAPORAN DATA PERWAKILAN', title_format)
        
        current_row = 1
        if search:
            worksheet.write(current_row, 0, f"Filter: {search}", subtitle_format)
            current_row += 1
            
        worksheet.write(current_row, 0, f"Tanggal Export: {datetime.now().strftime('%d-%m-%Y %H:%M')}", subtitle_format)
        current_row += 2
        
        # Write headers
        headers = [
            "No",
            "Trigram",
            "Bigram",
            "Nama Perwakilan",
            "Negara",
            "Jenis PWK",
        ]
        
        for col, header in enumerate(headers):
            worksheet.write(current_row, col, header, header_format)
        
        current_row += 1
        
        # Write data
        for idx, item in enumerate(perwakilan_list, 1):
            worksheet.write(current_row, 0, idx, center_format)  # No
            worksheet.write(current_row, 1, item[0] or '-', center_format)  # Trigram
            worksheet.write(current_row, 2, item[1] or '-', center_format)  # Bigram
            worksheet.write(current_row, 3, item[2] or '-', data_format)  # Nama Perwakilan
            worksheet.write(current_row, 4, item[3] or '-', data_format)  # Negara
            worksheet.write(current_row, 5, item[4] or '-', center_format)  # Jenis PWK
            
            current_row += 1
        
        # Auto-adjust column widths based on content
        col_widths = [
            ("No", 5),
            ("Trigram", 10),
            ("Bigram", 10),
            ("Nama Perwakilan", 30),
            ("Negara", 20),
            ("Jenis PWK", 15),
        ]
        
        for i, (header, default_width) in enumerate(col_widths):
            max_length = len(header)
            for row in perwakilan_list:
                if i == 0:  # No column
                    val_length = len(str(idx))
                elif i == 1:  # Trigram
                    val_length = len(str(row[0] or ''))
                elif i == 2:  # Bigram
                    val_length = len(str(row[1] or ''))
                elif i == 3:  # Nama Perwakilan
                    val_length = len(str(row[2] or ''))
                elif i == 4:  # Negara
                    val_length = len(str(row[3] or ''))
                elif i == 5:  # Jenis PWK
                    val_length = len(str(row[4] or ''))
                elif i == 6:  # No Urut
                    val_length = len(str(row[5] or ''))
                
                if val_length > max_length:
                    max_length = val_length
            
            # Set column width with some padding
            worksheet.set_column(i, i, min(max_length + 2, 50))
        
        workbook.close()
        output.seek(0)
        
        return send_file(
            output,
            as_attachment=True,
            download_name=f"data_perwakilan_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        print(f"Excel Export Error: {str(e)}")
        flash('Gagal membuat Excel: ' + str(e), 'error')
        return redirect(url_for('list_perwakilan'))

@app.route('/perwakilan')
def list_perwakilan():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    search = request.args.get('search', '').strip()
    page = request.args.get('page', 1, type=int)
    per_page = 20
    sort_column = request.args.get('sort', 'NO_URUTAN')
    sort_direction = request.args.get('dir', 'asc')

    # Validate sort column
    valid_columns = {
        'TRIGRAM': 'TRIGRAM',
        'BIGRAM': 'BIGRAM',
        'NAMA_PERWAKILAN': 'NAMA_PERWAKILAN',
        'NEGARA': 'NEGARA',
        'JENIS_PWK': 'JENIS_PWK',
        'NO_URUTAN': 'NO_URUTAN'
    }
    sort_column = valid_columns.get(sort_column, 'NO_URUTAN')
    
    # Validate sort direction
    sort_direction = 'DESC' if sort_direction.lower() == 'desc' else 'ASC'

    # Base query
    query = """
        SELECT TRIGRAM, BIGRAM, NAMA_PERWAKILAN, NEGARA, JENIS_PWK
        FROM REF_PERWAKILAN 
        WHERE NAMA_PERWAKILAN ILIKE %s OR 
              NEGARA ILIKE %s OR 
              TRIGRAM ILIKE %s OR
              BIGRAM ILIKE %s OR
              JENIS_PWK ILIKE %s
        ORDER BY {} {}
    """.format(sort_column, sort_direction)
    
    search_param = f'%{search}%'
    
    # Get total count
    count_query = """
        SELECT COUNT(*) FROM REF_PERWAKILAN 
        WHERE NAMA_PERWAKILAN ILIKE %s OR 
              NEGARA ILIKE %s OR 
              TRIGRAM ILIKE %s OR
              BIGRAM ILIKE %s OR
              JENIS_PWK ILIKE %s
    """
    total = execute_query(count_query, 
                         (search_param, search_param, search_param, 
                          search_param, search_param), 
                         fetch_one=True)[0] or 0
    
    # Add pagination
    paginated_query = query + f" LIMIT {per_page} OFFSET {(page - 1) * per_page}"
    perwakilan_list = execute_query(paginated_query, 
                                  (search_param, search_param, search_param, 
                                   search_param, search_param), 
                                  fetch=True) or []
    
    total_pages = (total + per_page - 1) // per_page
    
    return render_template('perwakilan/list.html', 
                         perwakilan_list=perwakilan_list,
                         search=search,
                         page=page,
                         per_page=per_page,
                         total=total,
                         total_pages=total_pages,
                         sort_column=sort_column,
                         sort_direction=sort_direction)

@app.route('/perwakilan/create', methods=['GET', 'POST'])
def create_perwakilan():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    if request.method == 'POST':
        # Validate jenis_pwk
        jenis_pwk = request.form.get('jenis_pwk', '').strip()
        if not jenis_pwk or jenis_pwk not in VALID_JENIS_PWK:
            flash('Jenis perwakilan tidak valid', 'error')
            return redirect(url_for('create_perwakilan'))
        
        try:
            # Get auto-increment numbers
            next_urutan = get_next_urutan()
            next_no_perwakilan = get_next_no_perwakilan()
            
            data = (
                request.form.get('trigram', '').strip().upper(),
                request.form.get('bigram', '').strip().upper(),
                request.form.get('nama_perwakilan', '').strip().upper(),
                request.form.get('negara', '').strip(),
                jenis_pwk,
                next_no_perwakilan,
                next_urutan,
                session.get('username', 'system'),
                datetime.now(),
                session.get('username', 'system'),
                datetime.now()
            )
            
            # Validate required fields
            if not all(data[:4]):
                flash('Semua field wajib diisi', 'error')
                return redirect(url_for('create_perwakilan'))
            
            success = execute_query("""
                INSERT INTO REF_PERWAKILAN 
                (TRIGRAM, BIGRAM, NAMA_PERWAKILAN, NEGARA, JENIS_PWK, 
                 NO_PERWAKILAN, NO_URUTAN, USER_INPUT, DATE_INPUT, 
                 USER_UPDATE, DATE_UPDATE)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """, data, commit=True)
            
            if success:
                flash('Data perwakilan berhasil ditambahkan', 'success')
                return redirect(url_for('list_perwakilan'))
            else:
                flash('Gagal menambahkan data perwakilan', 'error')
                
        except Exception as e:
            logger.error(f"Error creating perwakilan: {str(e)}")
            flash('Terjadi kesalahan saat menambahkan data', 'error')
    
    # For GET request
    next_urutan = get_next_urutan()
    next_no_perwakilan = get_next_no_perwakilan()
    return render_template('perwakilan/create.html', 
                         next_urutan=next_urutan,
                         next_no_perwakilan=next_no_perwakilan,
                         jenis_pwk_options=VALID_JENIS_PWK)

@app.route('/perwakilan/edit/<trigram>', methods=['GET', 'POST'])
def edit_perwakilan(trigram):
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    if request.method == 'POST':
        # Validate jenis_pwk
        jenis_pwk = request.form.get('jenis_pwk')
        if jenis_pwk not in VALID_JENIS_PWK:
            flash('Jenis perwakilan tidak valid', 'error')
            return redirect(url_for('edit_perwakilan', trigram=trigram))
        
        # Handle no_perwakilan - convert empty string to None or default value
        no_perwakilan = request.form.get('no_perwakilan')
        try:
            no_perwakilan = int(no_perwakilan) if no_perwakilan else 0
        except ValueError:
            no_perwakilan = 0
        
        data = (
            request.form.get('trigram', '').strip().upper(),  # Uppercase
            request.form.get('bigram', '').strip().upper(),   # Uppercase
            request.form.get('nama_perwakilan', '').strip().upper(),  # Uppercase
            request.form.get('negara', '').strip(),
            jenis_pwk,
            no_perwakilan,
            request.form.get('no_urutan'),
            session.get('username', 'system'),
            datetime.now(),
            trigram
        )
        
        success = execute_query("""
            UPDATE REF_PERWAKILAN SET
                TRIGRAM = %s,
                BIGRAM = %s,
                NAMA_PERWAKILAN = %s,
                NEGARA = %s,
                JENIS_PWK = %s,
                NO_PERWAKILAN = %s,
                NO_URUTAN = %s,
                USER_UPDATE = %s,
                DATE_UPDATE = %s
            WHERE TRIGRAM = %s
        """, data, commit=True)
        
        if success:
            flash('Data perwakilan berhasil diperbarui', 'success')
            return redirect(url_for('list_perwakilan'))
        else:
            flash('Gagal memperbarui data perwakilan', 'error')
    
    perwakilan = execute_query(
        "SELECT * FROM REF_PERWAKILAN WHERE TRIGRAM = %s",
        (trigram,),
        fetch_one=True
    )
    
    if not perwakilan:
        flash('Data perwakilan tidak ditemukan', 'error')
        return redirect(url_for('list_perwakilan'))
    
    return render_template('perwakilan/edit.html', 
                         perwakilan=perwakilan,
                         jenis_pwk_options=VALID_JENIS_PWK)

@app.route('/perwakilan/delete/<trigram>', methods=['POST'])
def delete_perwakilan(trigram):
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    success = execute_query(
        "DELETE FROM REF_PERWAKILAN WHERE TRIGRAM = %s",
        (trigram,),
        commit=True
    )
    
    if success:
        flash('Data perwakilan berhasil dihapus', 'success')
    else:
        flash('Gagal menghapus data perwakilan', 'error')
    
    return redirect(url_for('list_perwakilan'))

# ==============================================
# KEPRI CRUD ROUTES
# ==============================================

@app.route('/kepri/export/pdf')
def export_kepri_pdf():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    try:
        # Get all parameters from the request
        search = request.args.get('search', '').strip()
        sort_column = request.args.get('sort', 'no')
        sort_direction = request.args.get('dir', 'asc')

        # Validate sort column
        valid_columns = {
            'no': 'k.no',
            'nama': 'k.nama',
            'tahun': 'k.tahun',
            'perwakilan': 'r.nama_perwakilan'
        }
        sort_column = valid_columns.get(sort_column, 'k.no')
        sort_direction = 'DESC' if sort_direction.lower() == 'desc' else 'ASC'

        # Base query with joins
        query = """
            SELECT k.no, k.nama, k.tahun, k.id_pwk, r.nama_perwakilan, k.status, k.keterangan
            FROM tabel_kepri k
            LEFT JOIN ref_perwakilan r ON k.id_pwk = r.trigram
            WHERE (k.nama ILIKE %s OR 
                  r.nama_perwakilan ILIKE %s OR
                  k.id_pwk ILIKE %s OR
                  k.tahun::text ILIKE %s)
        """

        # Parameters for the query
        search_param = f'%{search}%'
        params = [search_param for _ in range(4)]

        # Add perwakilan filter for non-admin users
        if session.get('role') != 0:  # If not admin
            user_trigram = session.get('trigram')
            if user_trigram:
                query += " AND k.id_pwk = %s"
                params.append(user_trigram)

        # Sort by the selected column
        query += f" ORDER BY {sort_column} {sort_direction}"

        # Execute query to get all data
        kepri_list = execute_query(query, params, fetch=True) or []

        # Create PDF with landscape orientation
        buffer = BytesIO()
        doc_width, doc_height = landscape(letter)
        doc = SimpleDocTemplate(buffer, 
                              pagesize=landscape(letter),
                              leftMargin=15,
                              rightMargin=15,
                              topMargin=20,
                              bottomMargin=20)
        elements = []
        
        # Styles configuration
        styles = getSampleStyleSheet()
        
        # Title style
        title_style = ParagraphStyle(
            'Title',
            parent=styles['Title'],
            fontSize=14,
            alignment=TA_CENTER,
            spaceAfter=12
        )
        
        # Metadata style
        meta_style = ParagraphStyle(
            'Meta',
            parent=styles['Normal'],
            fontSize=9,
            spaceAfter=6
        )
        
        # Add title
        elements.append(Paragraph("LAPORAN DATA KEPRI", title_style))
        
        # Add metadata
        metadata = [
            f"Tanggal Export: {datetime.now().strftime('%d-%m-%Y %H:%M')}"
        ]
        if search:
            metadata.insert(0, f"Filter: {search}")
        
        for meta in metadata:
            elements.append(Paragraph(meta, meta_style))
        
        # Prepare table data
        data = []
        headers = [
            "No", 
            "Perwakilan", 
            "Nama", 
            "Tahun", 
            "Status",
            "Keterangan"
        ]
        data.append(headers)
        
        for idx, item in enumerate(kepri_list, 1):
            status = "Aktif" if item[5] == 1 else "Tidak Aktif"
            row = [
                str(idx),
                item[4] or '-',  # nama_perwakilan
                item[1] or '-',  # nama
                str(item[2]) if item[2] else '-',  # tahun
                status,
                item[6] or '-' if item[5] == 0 else '-'  # keterangan only if not active
            ]
            data.append(row)
        
        # Calculate available width (subtract margins)
        available_width = doc_width - doc.leftMargin - doc.rightMargin
        
        # Define column distribution
        col_ratios = [0.05, 0.25, 0.25, 0.1, 0.15, 0.2]  # Sum should be 1.0
        col_widths = [available_width * ratio for ratio in col_ratios]
        
        # Ensure minimum widths
        min_widths = [30, 80, 80, 40, 60, 80]  # Minimum widths in points
        col_widths = [max(w, min_w) for w, min_w in zip(col_widths, min_widths)]
        
        # Create table
        table = Table(data, 
                     colWidths=col_widths,
                     repeatRows=1,
                     hAlign='LEFT')
        
        # Table style
        style = TableStyle([
            # Header style
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#4472C4')),
            ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,0), 8),
            ('ALIGN', (0,0), (-1,0), 'CENTER'),
            ('BOTTOMPADDING', (0,0), (-1,0), 6),
            
            # Cell style
            ('FONTSIZE', (0,1), (-1,-1), 8),
            ('LEADING', (0,0), (-1,-1), 9),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey),
            ('WORDWRAP', (0,0), (-1,-1), True),
            
            # Alternating row colors
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.whitesmoke]),
            
            # Specific column alignments
            ('ALIGN', (0,0), (0,-1), 'CENTER'),    # No column centered
            ('ALIGN', (1,0), (1,-1), 'LEFT'),      # Perwakilan left-aligned
            ('ALIGN', (2,0), (2,-1), 'LEFT'),      # Nama left-aligned
            ('ALIGN', (3,0), (3,-1), 'CENTER'),    # Tahun centered
            ('ALIGN', (4,0), (4,-1), 'CENTER'),    # Status centered
            ('ALIGN', (5,0), (5,-1), 'LEFT')       # Keterangan left-aligned
        ])
        
        table.setStyle(style)
        elements.append(table)
        
        # Build PDF
        doc.build(elements)
        buffer.seek(0)
        
        return send_file(
            buffer,
            as_attachment=True,
            download_name=f"data_kepri_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
            mimetype='application/pdf'
        )
        
    except Exception as e:
        print(f"PDF Export Error: {str(e)}")
        flash('Gagal membuat PDF: ' + str(e), 'error')
        return redirect(url_for('list_kepri'))

@app.route('/kepri/export/excel')
def export_kepri_excel():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    try:
        # Get all parameters from the request
        search = request.args.get('search', '').strip()
        sort_column = request.args.get('sort', 'no')
        sort_direction = request.args.get('dir', 'asc')

        # Validate sort column
        valid_columns = {
            'no': 'k.no',
            'nama': 'k.nama',
            'tahun': 'k.tahun',
            'perwakilan': 'r.nama_perwakilan'
        }
        sort_column = valid_columns.get(sort_column, 'k.no')
        sort_direction = 'DESC' if sort_direction.lower() == 'desc' else 'ASC'

        # Base query with joins
        query = """
            SELECT k.no, k.nama, k.tahun, k.id_pwk, r.nama_perwakilan, k.status, k.keterangan
            FROM tabel_kepri k
            LEFT JOIN ref_perwakilan r ON k.id_pwk = r.trigram
            WHERE (k.nama ILIKE %s OR 
                  r.nama_perwakilan ILIKE %s OR
                  k.id_pwk ILIKE %s OR
                  k.tahun::text ILIKE %s)
        """

        # Parameters for the query
        search_param = f'%{search}%'
        params = [search_param for _ in range(4)]

        # Add perwakilan filter for non-admin users
        if session.get('role') != 0:  # If not admin
            user_trigram = session.get('trigram')
            if user_trigram:
                query += " AND k.id_pwk = %s"
                params.append(user_trigram)

        # Sort by the selected column
        query += f" ORDER BY {sort_column} {sort_direction}"

        # Execute query to get all data
        kepri_list = execute_query(query, params, fetch=True) or []

        # Create Excel file in memory
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet("Data KEPRI")
        
        # Add formats
        title_format = workbook.add_format({
            'bold': True,
            'font_size': 14,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        subtitle_format = workbook.add_format({
            'font_size': 10,
            'align': 'left',
            'valign': 'vcenter'
        })
        
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#4472C4',
            'font_color': '#FFFFFF',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True
        })
        
        # Data formats
        data_format = workbook.add_format({
            'border': 1,
            'text_wrap': True,
            'valign': 'top'
        })
        
        center_format = workbook.add_format({
            'border': 1,
            'text_wrap': True,
            'valign': 'top',
            'align': 'center'
        })
        
        # Write title and metadata
        worksheet.merge_range('A1:F1', 'LAPORAN DATA KEPRI', title_format)
        
        current_row = 1
        if search:
            worksheet.write(current_row, 0, f"Filter: {search}", subtitle_format)
            current_row += 1
            
        worksheet.write(current_row, 0, f"Tanggal Export: {datetime.now().strftime('%d-%m-%Y %H:%M')}", subtitle_format)
        current_row += 2
        
        # Write headers
        headers = [
            "No", 
            "Perwakilan", 
            "Nama", 
            "Tahun", 
            "Status",
            "Keterangan"
        ]
        
        for col, header in enumerate(headers):
            worksheet.write(current_row, col, header, header_format)
        
        current_row += 1
        
        # Write data
        for idx, item in enumerate(kepri_list, 1):
            status = "Aktif" if item[5] == 1 else "Tidak Aktif"
            keterangan = item[6] if item[5] == 0 else '-'
            
            row = [
                idx,
                item[4] or '-',  # nama_perwakilan
                item[1] or '-',  # nama
                item[2] or '-',  # tahun
                status,
                keterangan
            ]
            
            # Apply different formats based on column
            worksheet.write(current_row, 0, row[0], center_format)  # No - centered
            worksheet.write(current_row, 1, row[1], data_format)     # Perwakilan
            worksheet.write(current_row, 2, row[2], data_format)     # Nama
            worksheet.write(current_row, 3, row[3], center_format)  # Tahun - centered
            worksheet.write(current_row, 4, row[4], center_format)  # Status - centered
            worksheet.write(current_row, 5, row[5], data_format)     # Keterangan
            
            current_row += 1
        
        # Auto-adjust column widths based on content
        for i, header in enumerate(headers):
            max_length = len(header)
            for row in kepri_list:
                if i == 0:  # No column
                    val_length = len(str(row[0]))
                elif i == 1:  # Perwakilan
                    val_length = len(str(row[4] or ''))
                elif i == 2:  # Nama
                    val_length = len(str(row[1] or ''))
                elif i == 3:  # Tahun
                    val_length = len(str(row[2] or ''))
                elif i == 4:  # Status
                    val_length = 6  # "Aktif" or "Tidak Aktif"
                elif i == 5:  # Keterangan
                    val_length = len(str(row[6] or '')) if row[5] == 0 else 0
                
                if val_length > max_length:
                    max_length = val_length
            
            # Set column width with some padding
            worksheet.set_column(i, i, min(max_length + 2, 50))  # Max width 50 chars
        
        workbook.close()
        output.seek(0)
        
        return send_file(
            output,
            as_attachment=True,
            download_name=f"data_kepri_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        print(f"Excel Export Error: {str(e)}")
        flash('Gagal membuat Excel: ' + str(e), 'error')
        return redirect(url_for('list_kepri'))

@app.route('/kepri')
def list_kepri():
    if 'user_id' not in session:
        return redirect(url_for('login'))

    # Get search and pagination parameters
    search = request.args.get('search', '').strip()
    page = request.args.get('page', 1, type=int)
    per_page = 20
    sort_column = request.args.get('sort', 'no')
    sort_direction = request.args.get('dir', 'asc')

    # Validate sort column
    valid_columns = {
        'no': 'k.no',
        'nama': 'k.nama',
        'tahun': 'k.tahun',
        'perwakilan': 'r.nama_perwakilan'
    }
    sort_column = valid_columns.get(sort_column, 'k.no')
    
    # Validate sort direction
    sort_direction = 'DESC' if sort_direction.lower() == 'desc' else 'ASC'

    # Base query with join to get perwakilan name
    query = """
        SELECT k.no, k.nama, k.tahun, k.id_pwk, r.nama_perwakilan, k.status, k.keterangan
        FROM tabel_kepri k
        LEFT JOIN ref_perwakilan r ON k.id_pwk = r.trigram
        WHERE (k.nama ILIKE %s OR 
              r.nama_perwakilan ILIKE %s OR
              k.id_pwk ILIKE %s OR
              k.tahun::text ILIKE %s)
    """

    # Parameters for the query
    params = [f'%{search}%' for _ in range(4)]

    # Add perwakilan filter for non-admin users
    if session.get('role') != 0:  # If not admin
        user_trigram = session.get('trigram')
        if user_trigram:
            query += " AND k.id_pwk = %s"
            params.append(user_trigram)

    # Add sorting
    query += f" ORDER BY {sort_column} {sort_direction}"

    # Count query for pagination - simplified to just count rows
    count_query = """
        SELECT COUNT(*) 
        FROM tabel_kepri k
        LEFT JOIN ref_perwakilan r ON k.id_pwk = r.trigram
        WHERE (k.nama ILIKE %s OR 
              r.nama_perwakilan ILIKE %s OR
              k.id_pwk ILIKE %s OR
              k.tahun::text ILIKE %s)
    """
    
    # Add perwakilan filter to count query if needed
    if session.get('role') != 0 and session.get('trigram'):
        count_query += " AND k.id_pwk = %s"

    # Get total count
    total = execute_query(count_query, params, fetch_one=True)[0] or 0
    
    # Add pagination to main query
    paginated_query = query + f" LIMIT {per_page} OFFSET {(page - 1) * per_page}"
    
    # Execute query
    kepri_list = execute_query(paginated_query, params, fetch=True) or []
    
    total_pages = (total + per_page - 1) // per_page if total > 0 else 1
    
    return render_template('kepri/list.html',
                         kepri_list=kepri_list,
                         search=search,
                         page=page,
                         per_page=per_page,
                         total=total,
                         total_pages=total_pages,
                         sort_column=sort_column.replace('k.', ''),
                         sort_direction=sort_direction)

@app.route('/kepri/create', methods=['GET', 'POST'])
def create_kepri():
    if 'user_id' not in session:
        return redirect(url_for('login'))

    # Get perwakilan list based on user role
    if session.get('role') == 0:  # Admin
        perwakilan_list = execute_query(
            "SELECT trigram, nama_perwakilan FROM ref_perwakilan ORDER BY nama_perwakilan",
            fetch=True
        ) or []
    else:  # Regular user
        user_trigram = session.get('trigram')
        perwakilan_list = execute_query(
            "SELECT trigram, nama_perwakilan FROM ref_perwakilan WHERE trigram = %s",
            (user_trigram,),
            fetch=True
        ) or []

    if request.method == 'POST':
        nama = request.form.get('nama', '').strip()
        tahun = request.form.get('tahun', '').strip()
        id_pwk = request.form.get('id_pwk', '').strip().upper()
        status = int(request.form.get('status', 1))  # Default aktif
        keterangan = request.form.get('keterangan', '').strip() if status == 0 else None

        # Validate inputs
        if not nama:
            flash('Nama wajib diisi', 'error')
            return render_template('kepri/create.html', perwakilan_list=perwakilan_list)
        
        if not tahun or not tahun.isdigit():
            flash('Tahun harus berupa angka', 'error')
            return render_template('kepri/create.html', perwakilan_list=perwakilan_list)
        
        # For non-admin users, verify they're not trying to create for another perwakilan
        if session.get('role') != 0 and id_pwk != session.get('trigram'):
            flash('Anda hanya bisa membuat data untuk perwakilan Anda sendiri', 'error')
            return render_template('kepri/create.html', perwakilan_list=perwakilan_list)

        try:
            data = (nama, int(tahun), id_pwk, status, keterangan)
            success = execute_query(
                "INSERT INTO tabel_kepri (nama, tahun, id_pwk, status, keterangan) VALUES (%s, %s, %s, %s, %s)",
                data,
                commit=True
            )
            
            if success:
                flash('Data KEPRI berhasil ditambahkan', 'success')
                return redirect(url_for('list_kepri'))
            else:
                flash('Gagal menambahkan data KEPRI', 'error')
        except Exception as e:
            logger.error(f"Error creating KEPRI: {str(e)}")
            flash('Terjadi kesalahan saat menyimpan data', 'error')

    return render_template('kepri/create.html', perwakilan_list=perwakilan_list)

@app.route('/kepri/edit/<int:no>', methods=['GET', 'POST'])
def edit_kepri(no):
    if 'user_id' not in session:
        return redirect(url_for('login'))

    # Get existing data with perwakilan info
    kepri = execute_query(
        """SELECT k.no, k.nama, k.tahun, k.id_pwk, r.nama_perwakilan, k.status, k.keterangan 
           FROM tabel_kepri k
           LEFT JOIN ref_perwakilan r ON k.id_pwk = r.trigram
           WHERE k.no = %s""",
        (no,),
        fetch_one=True
    )
    
    if not kepri:
        flash('Data KEPRI tidak ditemukan', 'error')
        return redirect(url_for('list_kepri'))

    # Authorization check for non-admin users
    if session.get('role') != 0:  # If not admin
        user_trigram = session.get('trigram')
        if user_trigram != kepri[3]:  # kepri[3] is id_pwk
            flash('Anda tidak memiliki akses untuk mengedit data ini', 'error')
            return redirect(url_for('list_kepri'))

    # Get perwakilan list based on user role
    if session.get('role') == 0:  # Admin
        perwakilan_list = execute_query(
            "SELECT trigram, nama_perwakilan FROM ref_perwakilan ORDER BY nama_perwakilan",
            fetch=True
        ) or []
    else:  # Regular user
        perwakilan_list = execute_query(
            "SELECT trigram, nama_perwakilan FROM ref_perwakilan WHERE trigram = %s",
            (session.get('trigram'),),
            fetch=True
        ) or []

    if request.method == 'POST':
        nama = request.form.get('nama', '').strip()
        tahun = request.form.get('tahun', '').strip()
        id_pwk = request.form.get('id_pwk', '').strip().upper()
        status = int(request.form.get('status', 1))
        keterangan = request.form.get('keterangan', '').strip() if status == 0 else ''

        # Debugging - print form data
        print("Form Data:", {
            'nama': nama,
            'tahun': tahun,
            'id_pwk': id_pwk,
            'status': status,
            'keterangan': keterangan
        })

        # Validate inputs
        if not nama:
            flash('Nama wajib diisi', 'error')
            return render_template('kepri/edit.html', kepri=kepri, perwakilan_list=perwakilan_list)
        
        if not tahun or not tahun.isdigit():
            flash('Tahun harus berupa angka', 'error')
            return render_template('kepri/edit.html', kepri=kepri, perwakilan_list=perwakilan_list)
        
        if not id_pwk:
            flash('Perwakilan wajib dipilih', 'error')
            return render_template('kepri/edit.html', kepri=kepri, perwakilan_list=perwakilan_list)

        try:
            # Prepare data - pastikan urutan sesuai dengan query
            data = (
                nama,          # %s - nama
                int(tahun),    # %s - tahun
                id_pwk,        # %s - id_pwk
                status,        # %s - status
                keterangan,    # %s - keterangan
                no             # %s - WHERE no
            )
            
            # Debugging - print prepared data
            print("Prepared Data:", data)
            
            # Execute query
            success = execute_query(
                """UPDATE tabel_kepri 
                   SET nama = %s, tahun = %s, id_pwk = %s, status = %s, keterangan = %s 
                   WHERE no = %s""",
                data,
                commit=True
            )
            
            if success:
                flash('Data KEPRI berhasil diperbarui', 'success')
                # Redirect dengan menyertakan parameter sorting/pagination yang sama
                return redirect(url_for('list_kepri', 
                                     page=request.args.get('page', 1),
                                     sort=request.args.get('sort', 'no'),
                                     dir=request.args.get('dir', 'asc')))
            else:
                flash('Gagal memperbarui data KEPRI', 'error')
        except Exception as e:
            logger.error(f"Error updating KEPRI: {str(e)}", exc_info=True)
            flash(f'Terjadi kesalahan saat memperbarui data: {str(e)}', 'error')
    
    return render_template('kepri/edit.html', kepri=kepri, perwakilan_list=perwakilan_list)

@app.route('/kepri/delete/<int:no>', methods=['POST'])
def delete_kepri(no):
    if 'user_id' not in session:
        return redirect(url_for('login'))

    try:
        success = execute_query(
            "DELETE FROM tabel_kepri WHERE no = %s",
            (no,),
            commit=True
        )
        
        if success:
            flash('Data KEPRI berhasil dihapus', 'success')
        else:
            flash('Gagal menghapus data KEPRI', 'error')
    except Exception as e:
        logger.error(f"Error deleting KEPRI: {str(e)}")
        flash('Terjadi kesalahan saat menghapus data', 'error')
    
    return redirect(url_for('list_kepri'))

# ==============================================
# JABATAN CRUD ROUTES
# ==============================================

@app.route('/jabatan')
def list_jabatan():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    # Get search and pagination parameters
    search = request.args.get('search', '').strip()
    page = request.args.get('page', 1, type=int)
    per_page = 20
    sort_column = request.args.get('sort', 'no')
    sort_direction = request.args.get('dir', 'asc')

    # Validate sort column
    valid_columns = {
        'no': 'no',
        'nama': 'nama',
        'singkatan': 'singkatan'
    }
    sort_column = valid_columns.get(sort_column, 'no')
    
    # Validate sort direction
    sort_direction = 'DESC' if sort_direction.lower() == 'desc' else 'ASC'

    # Base query
    query = f"""
        SELECT no, nama, singkatan 
        FROM tabel_jabatan
        WHERE nama ILIKE %s OR 
              singkatan ILIKE %s
        ORDER BY {sort_column} {sort_direction}
    """

    # Count query for pagination
    count_query = """
        SELECT COUNT(*)
        FROM tabel_jabatan
        WHERE nama ILIKE %s OR 
              singkatan ILIKE %s
    """

    search_param = f'%{search}%'
    
    # Get total count
    total = execute_query(count_query, 
                        (search_param, search_param), 
                        fetch_one=True)[0] or 0
    
    # Add pagination to main query
    paginated_query = query + f" LIMIT {per_page} OFFSET {(page - 1) * per_page}"
    
    # Execute query
    jabatan_list = execute_query(paginated_query, 
                               (search_param, search_param), 
                               fetch=True) or []
    
    total_pages = (total + per_page - 1) // per_page
    
    return render_template('jabatan/list.html',
                         jabatan_list=jabatan_list,
                         search=search,
                         page=page,
                         per_page=per_page,
                         total=total,
                         total_pages=total_pages,
                         sort_column=sort_column,
                         sort_direction=sort_direction)

@app.route('/jabatan/create', methods=['GET', 'POST'])
def create_jabatan():
    if 'user_id' not in session:
        return redirect(url_for('login'))

    if request.method == 'POST':
        nama = request.form.get('nama', '').strip()
        singkatan = request.form.get('singkatan', '').strip()

        # Validate inputs
        if not nama:
            flash('Nama wajib diisi', 'error')
            return redirect(url_for('create_jabatan'))
        
        try:
            data = (nama, singkatan)
            success = execute_query(
                "INSERT INTO tabel_jabatan (nama, singkatan) VALUES (%s, %s)",
                data,
                commit=True
            )
            
            if success:
                flash('Data jabatan berhasil ditambahkan', 'success')
                return redirect(url_for('list_jabatan'))
            else:
                flash('Gagal menambahkan data jabatan', 'error')
        except Exception as e:
            logger.error(f"Error creating jabatan: {str(e)}")
            flash('Terjadi kesalahan saat menyimpan data', 'error')

    return render_template('jabatan/create.html')

@app.route('/jabatan/edit/<int:no>', methods=['GET', 'POST'])
def edit_jabatan(no):
    if 'user_id' not in session:
        return redirect(url_for('login'))

    # Get existing data
    jabatan = execute_query(
        "SELECT no, nama, singkatan FROM tabel_jabatan WHERE no = %s",
        (no,),
        fetch_one=True
    )
    
    if not jabatan:
        flash('Data jabatan tidak ditemukan', 'error')
        return redirect(url_for('list_jabatan'))

    if request.method == 'POST':
        nama = request.form.get('nama', '').strip()
        singkatan = request.form.get('singkatan', '').strip()

        # Validate inputs
        if not nama:
            flash('Nama wajib diisi', 'error')
            return redirect(url_for('edit_jabatan', no=no))

        try:
            data = (nama, singkatan, no)
            success = execute_query(
                "UPDATE tabel_jabatan SET nama = %s, singkatan = %s WHERE no = %s",
                data,
                commit=True
            )
            
            if success:
                flash('Data jabatan berhasil diperbarui', 'success')
                return redirect(url_for('list_jabatan'))
            else:
                flash('Gagal memperbarui data jabatan', 'error')
        except Exception as e:
            logger.error(f"Error updating jabatan: {str(e)}")
            flash('Terjadi kesalahan saat memperbarui data', 'error')

    return render_template('jabatan/edit.html', jabatan=jabatan)

@app.route('/jabatan/delete/<int:no>', methods=['POST'])
def delete_jabatan(no):
    if 'user_id' not in session:
        return redirect(url_for('login'))

    try:
        success = execute_query(
            "DELETE FROM tabel_jabatan WHERE no = %s",
            (no,),
            commit=True
        )
        
        if success:
            flash('Data jabatan berhasil dihapus', 'success')
        else:
            flash('Gagal menghapus data jabatan', 'error')
    except Exception as e:
        logger.error(f"Error deleting jabatan: {str(e)}")
        flash('Terjadi kesalahan saat menghapus data', 'error')
    
    return redirect(url_for('list_jabatan'))

# ==============================================
# PERSONEL CRUD ROUTES (UPDATED)
# ==============================================

# Define pangkat/golongan options
PANGKAT_GOLONGAN = [
    'IV E/Pembina Utama',
    'IV D/Pembina Utama Madya',
    'IV C/Pembina Utama Madya',
    'IV B/Pembina Tingkat 1',
    'IV A/Pembina',
    'III D/Penata Tingkat 1',
    'III C/Penata',
    'III B/Penata Muda Tingkat 1',
    'III A/Penata Muda',
    'II D/Pengatur Tingkat 1',
    'II C/Pengatur',
    'II B/Pengatur Muda Tingkat 1',
    'II A/Pengatur Muda'
]

PENEMPATAN_OPTIONS = list(range(1, 8))  # 1-7

@app.route('/personel/export/pdf')
def export_personel_pdf():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    try:
        # Get all parameters from the request
        search = request.args.get('search', '').strip()
        sort_column = request.args.get('sort', 'no')
        sort_direction = request.args.get('dir', 'asc')

        # Validate sort column
        valid_columns = {
            'no': 'p.no',
            'nama': 'p.nama',
            'nip': 'p.nip',
            'pangkat_gol': 'p.pangkat_gol',
            'jabatan': 'j.nama',
            'perwakilan': 'r.nama_perwakilan'
        }
        sort_column = valid_columns.get(sort_column, 'p.no')
        sort_direction = 'DESC' if sort_direction.lower() == 'desc' else 'ASC'

        # Base query with joins
        query = """
            SELECT p.no, p.nama, p.nip, p.pangkat_gol, p.tmt_pangkat, 
                   p.id_jabatan, j.nama as nama_jabatan, p.tmt_jabatan,
                   p.penempatan, p.tmt_penempatan, p.id_pwk, r.nama_perwakilan
            FROM tabel_personel p
            LEFT JOIN tabel_jabatan j ON p.id_jabatan = j.no
            LEFT JOIN ref_perwakilan r ON p.id_pwk = r.trigram
            WHERE (p.nama ILIKE %s OR 
                  p.nip ILIKE %s OR
                  p.pangkat_gol ILIKE %s OR
                  j.nama ILIKE %s OR
                  r.nama_perwakilan ILIKE %s)
        """

        # Parameters for the query
        search_param = f'%{search}%'
        params = [search_param for _ in range(5)]

        # Add perwakilan filter for non-admin users
        if session.get('role') != 0:  # If not admin
            user_trigram = session.get('trigram')
            if user_trigram:
                query += " AND p.id_pwk = %s"
                params.append(user_trigram)

        # Sort by the selected column
        query += f" ORDER BY {sort_column} {sort_direction}"

        # Execute query to get all data
        personel_list = execute_query(query, params, fetch=True) or []

        # Create PDF with landscape orientation
        buffer = BytesIO()
        doc_width, doc_height = landscape(letter)
        doc = SimpleDocTemplate(buffer, 
                              pagesize=landscape(letter),
                              leftMargin=15,
                              rightMargin=15,
                              topMargin=20,
                              bottomMargin=20)
        elements = []
        
        # Styles configuration
        styles = getSampleStyleSheet()
        
        # Title style
        title_style = ParagraphStyle(
            'Title',
            parent=styles['Title'],
            fontSize=14,
            alignment=TA_CENTER,
            spaceAfter=12
        )
        
        # Metadata style
        meta_style = ParagraphStyle(
            'Meta',
            parent=styles['Normal'],
            fontSize=9,
            spaceAfter=6
        )
        
        # Add title
        elements.append(Paragraph("LAPORAN DATA PERSONEL", title_style))
        
        # Add metadata
        metadata = [
            f"Tanggal Export: {datetime.now().strftime('%d-%m-%Y %H:%M')}"
        ]
        if search:
            metadata.insert(0, f"Filter: {search}")
        
        for meta in metadata:
            elements.append(Paragraph(meta, meta_style))
        
        # Prepare table data
        data = []
        headers = [
            "No", 
            "Perwakilan", 
            "Nama", 
            "NIP", 
            "Pangkat/Gol", 
            "Jabatan", 
            "Penempatan Ke-"
        ]
        data.append(headers)
        
        for idx, item in enumerate(personel_list, 1):
            row = [
                str(idx),
                item[11] or '-',
                item[1] or '-',
                item[2] or '-',
                item[3] or '-',
                item[6] or '-',
                str(item[8]) if item[8] is not None else '-'
            ]
            data.append(row)
        
        # Calculate available width (subtract margins)
        available_width = doc_width - doc.leftMargin - doc.rightMargin
        
        # Define column distribution (adjust these ratios as needed)
        col_ratios = [0.05, 0.2, 0.15, 0.15, 0.2, 0.15, 0.1]  # Sum should be 1.0
        col_widths = [available_width * ratio for ratio in col_ratios]
        
        # Ensure minimum widths
        min_widths = [30, 60, 60, 60, 80, 60, 40]  # Minimum widths in points
        col_widths = [max(w, min_w) for w, min_w in zip(col_widths, min_widths)]
        
        # Create table
        table = Table(data, 
                     colWidths=col_widths,
                     repeatRows=1,
                     hAlign='LEFT')
        
        # Table style
        style = TableStyle([
            # Header style
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#4472C4')),
            ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,0), 8),
            ('ALIGN', (0,0), (-1,0), 'CENTER'),
            ('BOTTOMPADDING', (0,0), (-1,0), 6),
            
            # Cell style
            ('FONTSIZE', (0,1), (-1,-1), 8),
            ('LEADING', (0,0), (-1,-1), 9),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey),
            ('WORDWRAP', (0,0), (-1,-1), True),
            
            # Alternating row colors
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.whitesmoke]),
            
            # Specific column alignments
            ('ALIGN', (0,0), (0,-1), 'CENTER'),    # No column centered
            ('ALIGN', (1,0), (1,-1), 'LEFT'),      # Perwakilan left-aligned
            ('ALIGN', (2,0), (2,-1), 'LEFT'),      # Nama left-aligned
            ('ALIGN', (3,0), (3,-1), 'LEFT'),      # NIP left-aligned
            ('ALIGN', (4,0), (4,-1), 'LEFT'),      # Pangkat/Gol left-aligned
            ('ALIGN', (5,0), (5,-1), 'LEFT'),      # Jabatan left-aligned
            ('ALIGN', (6,0), (6,-1), 'CENTER')     # Penempatan centered
        ])
        
        table.setStyle(style)
        elements.append(table)
        
        # Build PDF
        doc.build(elements)
        buffer.seek(0)
        
        return send_file(
            buffer,
            as_attachment=True,
            download_name=f"data_personel_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
            mimetype='application/pdf'
        )
        
    except Exception as e:
        print(f"PDF Export Error: {str(e)}")
        flash('Gagal membuat PDF: ' + str(e), 'error')
        return redirect(url_for('list_personel'))

@app.route('/personel/export/excel')
def export_personel_excel():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    try:
        # Get all parameters from the request
        search = request.args.get('search', '').strip()
        sort_column = request.args.get('sort', 'no')
        sort_direction = request.args.get('dir', 'asc')

        # Validate sort column
        valid_columns = {
            'no': 'p.no',
            'nama': 'p.nama',
            'nip': 'p.nip',
            'pangkat_gol': 'p.pangkat_gol',
            'jabatan': 'j.nama',
            'perwakilan': 'r.nama_perwakilan'
        }
        sort_column = valid_columns.get(sort_column, 'p.no')
        sort_direction = 'DESC' if sort_direction.lower() == 'desc' else 'ASC'

        # Base query with joins
        query = """
            SELECT p.no, p.nama, p.nip, p.pangkat_gol, p.tmt_pangkat, 
                   p.id_jabatan, j.nama as nama_jabatan, p.tmt_jabatan,
                   p.penempatan, p.tmt_penempatan, p.id_pwk, r.nama_perwakilan
            FROM tabel_personel p
            LEFT JOIN tabel_jabatan j ON p.id_jabatan = j.no
            LEFT JOIN ref_perwakilan r ON p.id_pwk = r.trigram
            WHERE (p.nama ILIKE %s OR 
                  p.nip ILIKE %s OR
                  p.pangkat_gol ILIKE %s OR
                  j.nama ILIKE %s OR
                  r.nama_perwakilan ILIKE %s)
        """

        # Parameters for the query
        search_param = f'%{search}%'
        params = [search_param for _ in range(5)]

        # Add perwakilan filter for non-admin users
        if session.get('role') != 0:  # If not admin
            user_trigram = session.get('trigram')
            if user_trigram:
                query += " AND p.id_pwk = %s"
                params.append(user_trigram)

        # Sort by the selected column
        query += f" ORDER BY {sort_column} {sort_direction}"

        # Execute query to get all data
        personel_list = execute_query(query, params, fetch=True) or []

        # Create Excel file in memory
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet("Data Personel")
        
        # Add formats
        title_format = workbook.add_format({
            'bold': True,
            'font_size': 14,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        subtitle_format = workbook.add_format({
            'font_size': 10,
            'align': 'left',
            'valign': 'vcenter'
        })
        
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#4472C4',
            'font_color': '#FFFFFF',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True
        })
        
        # Data formats
        data_format = workbook.add_format({
            'border': 1,
            'text_wrap': True,
            'valign': 'top'
        })
        
        # Write title and metadata
        worksheet.merge_range('A1:G1', 'LAPORAN DATA PERSONEL', title_format)
        
        current_row = 1
        if search:
            worksheet.write(current_row, 0, f"Filter: {search}", subtitle_format)
            current_row += 1
            
        worksheet.write(current_row, 0, f"Tanggal Export: {datetime.now().strftime('%d-%m-%Y %H:%M')}", subtitle_format)
        current_row += 2
        
        # Write headers
        headers = [
            "No", 
            "Perwakilan", 
            "Nama", 
            "NIP", 
            "Pangkat/Gol", 
            "Jabatan", 
            "Penempatan Ke-"
        ]
        
        for col, header in enumerate(headers):
            worksheet.write(current_row, col, header, header_format)
        
        current_row += 1
        
        # Write data
        for idx, item in enumerate(personel_list, 1):
            row = [
                idx,
                item[11] or '-',
                item[1] or '-',
                item[2] or '-',
                item[3] or '-',
                item[6] or '-',
                item[8] or '-'
            ]
            
            for col, value in enumerate(row):
                worksheet.write(current_row, col, value, data_format)
            
            current_row += 1
        
        # Auto-adjust column widths based on content
        for i, header in enumerate(headers):
            max_length = len(header)
            for row in personel_list:
                if i == 0:  # No column
                    val_length = len(str(row[0]))
                elif i == 1:  # Perwakilan
                    val_length = len(str(row[11] or ''))
                elif i == 2:  # Nama
                    val_length = len(str(row[1] or ''))
                elif i == 3:  # NIP
                    val_length = len(str(row[2] or ''))
                elif i == 4:  # Pangkat/Gol
                    val_length = len(str(row[3] or ''))
                elif i == 5:  # Jabatan
                    val_length = len(str(row[6] or ''))
                elif i == 6:  # Penempatan
                    val_length = len(str(row[8] or ''))
                
                if val_length > max_length:
                    max_length = val_length
            
            # Set column width with some padding
            worksheet.set_column(i, i, min(max_length + 2, 50))  # Max width 50 chars
        
        workbook.close()
        output.seek(0)
        
        return send_file(
            output,
            as_attachment=True,
            download_name=f"data_personel_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        print(f"Excel Export Error: {str(e)}")
        flash('Gagal membuat Excel: ' + str(e), 'error')
        return redirect(url_for('list_personel'))

@app.route('/personel')
def list_personel():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    try:
        # Get search and pagination parameters
        search = request.args.get('search', '').strip()
        page = request.args.get('page', 1, type=int)
        per_page = 20

        # Base query
        query = """
            SELECT p.no, p.nama, p.nip, p.pangkat_gol, p.tmt_pangkat, 
                   p.id_jabatan, j.nama as nama_jabatan, p.tmt_jabatan,
                   p.penempatan, p.tmt_penempatan, p.id_pwk,pwk.nama_perwakilan
            FROM tabel_personel p
            LEFT JOIN tabel_jabatan j ON p.id_jabatan = j.no
            LEFT JOIN ref_perwakilan pwk ON pwk.trigram = p.id_pwk
        """
        params = []

        # Add search filter if exists
        if search:
            query += """
                WHERE (p.nama ILIKE %s OR 
                      p.nip ILIKE %s OR
                      p.pangkat_gol ILIKE %s OR
                      j.nama ILIKE %s OR
                      p.id_pwk ILIKE %s)
            """
            search_param = f'%{search}%'
            params.extend([search_param]*5)

        # Add perwakilan filter for non-admin users
        if session.get('role') != 0:  # If not admin
            user_trigram = session.get('trigram')
            if user_trigram:
                if search:
                    query += " AND UPPER(TRIM(p.id_pwk)) = UPPER(TRIM(%s))"
                else:
                    query += " WHERE UPPER(TRIM(p.id_pwk)) = UPPER(TRIM(%s))"
                params.append(user_trigram.strip().upper())

        # Count total records
        count_query = "SELECT COUNT(*) FROM (" + query + ") AS subquery"
        total = execute_query(count_query, params, fetch_one=True)[0] or 0

        # Add pagination
        query += " ORDER BY p.no LIMIT %s OFFSET %s"
        params.extend([per_page, (page - 1) * per_page])
        
        # Execute query
        personel_list = execute_query(query, params, fetch=True) or []

        total_pages = (total + per_page - 1) // per_page if total > 0 else 1

        return render_template('personel/list.html',
                            personel_list=personel_list,
                            search=search,
                            page=page,
                            per_page=per_page,
                            total=total,
                            total_pages=total_pages)

    except Exception as e:
        logger.error(f"Error in list_personel: {str(e)}")
        # Return empty result if error occurs
        return render_template('personel/list.html',
                            personel_list=[],
                            search='',
                            page=1,
                            per_page=20,
                            total=0,
                            total_pages=1)

@app.route('/personel/create', methods=['GET', 'POST'])
def create_personel():
    if 'user_id' not in session:
        return redirect(url_for('login'))

    # Get list of jabatan for dropdown
    jabatan_list = execute_query(
        "SELECT no, nama FROM tabel_jabatan ORDER BY nama",
        fetch=True
    ) or []

    # Get perwakilan list based on user role
    # Get perwakilan list based on user role
    if session.get('role') == 0:  # Admin
        perwakilan_list = execute_query(
            "SELECT trigram, nama_perwakilan FROM ref_perwakilan ORDER BY nama_perwakilan",
            fetch=True
        ) or []
    else:  # Regular user
        user_trigram = session.get('trigram')
        perwakilan_list = execute_query(
            "SELECT trigram, nama_perwakilan FROM ref_perwakilan WHERE trigram = %s",
            (user_trigram,),
            fetch=True
        ) or []

    if request.method == 'POST':
        try:
            # Get form data
            form_data = request.form
            
            # Prepare data tuple with proper null handling
            data = (
                form_data.get('nama', '').strip() or None,  # Required field
                form_data.get('t_lahir', '').strip() or None,
                form_data.get('tgl_lahir') or None,  # Date field
                form_data.get('nip', '').strip() or None,
                form_data.get('pangkat_gol') or None,
                form_data.get('tmt_pangkat') or None,  # Date field
                int(form_data.get('id_jabatan')) if form_data.get('id_jabatan') else None,
                form_data.get('tmt_jabatan') or None,  # Date field
                int(form_data.get('penempatan')) if form_data.get('penempatan') else None,
                form_data.get('tmt_penempatan') or None,  # Date field
                form_data.get('alamat', '').strip() or None,
                form_data.get('telp', '').strip() or None,
                form_data.get('email', '').strip() or None,
                form_data.get('tmt_selesai_penempatan') or None,  # Date field
                form_data.get('id_pwk', '').strip().upper() or None
            )

            # Validate required fields
            if not data[0]:  # nama
                flash('Nama wajib diisi', 'error')
                return render_template('personel/create.html', 
                                    jabatan_list=jabatan_list,
                                    perwakilan_list=perwakilan_list,
                                    pangkat_golongan=PANGKAT_GOLONGAN,
                                    penempatan_options=PENEMPATAN_OPTIONS,
                                    form_data=request.form)

            # For non-admin users, verify they're not trying to create for another perwakilan
            # For non-admin users, verify they're not trying to create for another perwakilan
            if session.get('role') != 0 and data[14] != session.get('trigram'):
                flash('Anda hanya bisa membuat data untuk perwakilan Anda sendiri', 'error')
                return render_template('personel/create.html', 
                                    jabatan_list=jabatan_list,
                                    perwakilan_list=perwakilan_list,
                                    pangkat_golongan=PANGKAT_GOLONGAN,
                                    penempatan_options=PENEMPATAN_OPTIONS,
                                    form_data=request.form)

            # Execute query with proper NULL handling
            success = execute_query("""
                INSERT INTO tabel_personel 
                (nama, t_lahir, tgl_lahir, nip, pangkat_gol, tmt_pangkat, 
                 id_jabatan, tmt_jabatan, penempatan, tmt_penempatan, 
                 alamat, telp, email, tmt_selesai_penempatan, id_pwk)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """, data, commit=True)
            
            if success:
                flash('Data personel berhasil ditambahkan', 'success')
                return redirect(url_for('list_personel'))
            else:
                flash('Gagal menambahkan data personel', 'error')
        except ValueError as e:
            logger.error(f"ValueError creating personel: {str(e)}")
            flash('Format data tidak valid. Pastikan semua field diisi dengan benar', 'error')
        except Exception as e:
            logger.error(f"Error creating personel: {str(e)}")
            flash('Terjadi kesalahan saat menyimpan data', 'error')

    return render_template('personel/create.html', 
                         jabatan_list=jabatan_list,
                         perwakilan_list=perwakilan_list,
                         pangkat_golongan=PANGKAT_GOLONGAN,
                         penempatan_options=PENEMPATAN_OPTIONS)

@app.route('/personel/edit/<int:no>', methods=['GET', 'POST'])
def edit_personel(no):
    if 'user_id' not in session:
        return redirect(url_for('login'))

    # Get existing data with perwakilan info
    personel = execute_query(
        """SELECT p.no, p.nama, p.t_lahir, p.tgl_lahir, p.nip, p.pangkat_gol, p.tmt_pangkat, 
           p.id_jabatan, p.tmt_jabatan, p.penempatan, p.tmt_penempatan, p.alamat, p.telp, p.email, 
           p.tmt_selesai_penempatan, p.id_pwk, r.nama_perwakilan
           FROM tabel_personel p
           LEFT JOIN ref_perwakilan r ON p.id_pwk = r.trigram
           WHERE p.no = %s""",
        (no,),
        fetch_one=True
    )
    
    if not personel:
        flash('Data personel tidak ditemukan', 'error')
        return redirect(url_for('list_personel'))

    # Authorization check for non-admin users
    # Authorization check for non-admin users
    if session.get('role') != 0:  # If not admin
        user_trigram = session.get('trigram')
        if user_trigram != personel[15]:  # personel[15] is id_pwk
            flash('Anda tidak memiliki akses untuk mengedit data ini', 'error')
            return redirect(url_for('list_personel'))

    # Get list of jabatan for dropdown
    jabatan_list = execute_query(
        "SELECT no, nama FROM tabel_jabatan ORDER BY nama",
        fetch=True
    ) or []

    # Get perwakilan list based on user role
    if session.get('role') == 0:  # Admin
        perwakilan_list = execute_query(
            "SELECT trigram, nama_perwakilan FROM ref_perwakilan ORDER BY nama_perwakilan",
            fetch=True
        ) or []
    else:  # Regular user
        perwakilan_list = execute_query(
            "SELECT trigram, nama_perwakilan FROM ref_perwakilan WHERE trigram = %s",
            (session.get('trigram'),),
            fetch=True
        ) or []

    if request.method == 'POST':
        try:
            # Get form data
            form_data = request.form
            
            # Prepare data tuple with proper null handling
            data = (
                form_data.get('nama', '').strip() or None,  # Required field
                form_data.get('t_lahir', '').strip() or None,
                form_data.get('tgl_lahir') or None,  # Date field
                form_data.get('nip', '').strip() or None,
                form_data.get('pangkat_gol') or None,
                form_data.get('tmt_pangkat') or None,  # Date field
                int(form_data.get('id_jabatan')) if form_data.get('id_jabatan') else None,
                form_data.get('tmt_jabatan') or None,  # Date field
                int(form_data.get('penempatan')) if form_data.get('penempatan') else None,
                form_data.get('tmt_penempatan') or None,  # Date field
                form_data.get('alamat', '').strip() or None,
                form_data.get('telp', '').strip() or None,
                form_data.get('email', '').strip() or None,
                form_data.get('tmt_selesai_penempatan') or None,  # Date field
                form_data.get('id_pwk', '').strip().upper() or None,
                no
            )

            # Validate required fields
            if not data[0]:  # nama
                flash('Nama wajib diisi', 'error')
                return redirect(url_for('edit_personel', no=no))

            # Check if perwakilan exists
            if data[14]:  # Only check if id_pwk is provided
                perwakilan_exists = execute_query(
                    "SELECT 1 FROM ref_perwakilan WHERE trigram = %s",
                    (data[14],),
                    fetch_one=True
                )
                
                if not perwakilan_exists:
                    flash('Perwakilan tidak valid', 'error')
                    return redirect(url_for('edit_personel', no=no))

            # Execute query with proper NULL handling
            success = execute_query("""
                UPDATE tabel_personel SET
                    nama = %s,
                    t_lahir = %s,
                    tgl_lahir = %s,
                    nip = %s,
                    pangkat_gol = %s,
                    tmt_pangkat = %s,
                    id_jabatan = %s,
                    tmt_jabatan = %s,
                    penempatan = %s,
                    tmt_penempatan = %s,
                    alamat = %s,
                    telp = %s,
                    email = %s,
                    tmt_selesai_penempatan = %s,
                    id_pwk = %s
                WHERE no = %s
            """, data, commit=True)
            
            if success:
                flash('Data personel berhasil diperbarui', 'success')
                return redirect(url_for('list_personel'))
            else:
                flash('Gagal memperbarui data personel', 'error')
        except ValueError as e:
            logger.error(f"ValueError updating personel: {str(e)}")
            flash('Format data tidak valid. Pastikan semua field diisi dengan benar', 'error')
        except Exception as e:
            logger.error(f"Error updating personel: {str(e)}")
            flash('Terjadi kesalahan saat memperbarui data', 'error')

    return render_template('personel/edit.html', 
                         personel=personel,
                         jabatan_list=jabatan_list,
                         perwakilan_list=perwakilan_list,
                         pangkat_golongan=PANGKAT_GOLONGAN,
                         penempatan_options=PENEMPATAN_OPTIONS)

@app.route('/personel/delete/<int:no>', methods=['POST'])
def delete_personel(no):
    if 'user_id' not in session:
        return redirect(url_for('login'))

    try:
        success = execute_query(
            "DELETE FROM tabel_personel WHERE no = %s",
            (no,),
            commit=True
        )
        
        if success:
            flash('Data personel berhasil dihapus', 'success')
        else:
            flash('Gagal menghapus data personel', 'error')
    except Exception as e:
        logger.error(f"Error deleting personel: {str(e)}")
        flash('Terjadi kesalahan saat menghapus data', 'error')
    
    return redirect(url_for('list_personel'))

# ==============================================
# PEGAWAI SETEMPAT CRUD ROUTES
# ==============================================

@app.route('/pegawai-setempat/export/pdf')
def export_pegawai_setempat_pdf():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    try:
        # Get all parameters from the request
        search = request.args.get('search', '').strip()
        sort_column = request.args.get('sort', 'no')
        sort_direction = request.args.get('dir', 'asc')

        # Validate sort column
        valid_columns = {
            'no': 'p.no',
            'nama': 'p.nama',
            'nik': 'p.nik',
            'tmt_penempatan': 'p.tmt_penempatan',
            'perwakilan': 'p.id_pwk'
        }
        sort_column = valid_columns.get(sort_column, 'p.no')
        sort_direction = 'DESC' if sort_direction.lower() == 'desc' else 'ASC'

        # Base query with joins
        query = """
            SELECT p.no, p.nama, p.t_lahir, p.tgl_lahir, p.nik, p.telp, p.email, 
                   p.tmt_penempatan, p.tmt_selesai_penempatan, p.id_pwk, r.nama_perwakilan
            FROM tabel_pegawai_setempat p
            LEFT JOIN ref_perwakilan r ON p.id_pwk = r.trigram
            WHERE (p.nama ILIKE %s OR 
                  p.nik ILIKE %s OR
                  p.telp ILIKE %s OR
                  p.email ILIKE %s OR
                  p.id_pwk ILIKE %s OR
                  r.nama_perwakilan ILIKE %s)
        """

        # Parameters for the query
        search_param = f'%{search}%'
        params = [search_param for _ in range(6)]

        # Add perwakilan filter for non-admin users
        if session.get('role') != 0:  # If not admin
            user_trigram = session.get('trigram')
            if user_trigram:
                query += " AND p.id_pwk = %s"
                params.append(user_trigram)

        # Sort by the selected column
        query += f" ORDER BY {sort_column} {sort_direction}"

        # Execute query to get all data
        pegawai_list = execute_query(query, params, fetch=True) or []

        # Create PDF with A4 size (portrait orientation)
        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, 
                              pagesize=A4,
                              leftMargin=15,
                              rightMargin=15,
                              topMargin=20,
                              bottomMargin=20)
        elements = []
        
        # Styles configuration
        styles = getSampleStyleSheet()
        
        # Title style
        title_style = ParagraphStyle(
            'Title',
            parent=styles['Title'],
            fontSize=14,
            alignment=TA_CENTER,
            spaceAfter=12
        )
        
        # Metadata style
        meta_style = ParagraphStyle(
            'Meta',
            parent=styles['Normal'],
            fontSize=9,
            spaceAfter=6
        )
        
        # Add title
        elements.append(Paragraph("LAPORAN DATA PEGAWAI SETEMPAT", title_style))
        
        # Add metadata
        metadata = [
            f"Tanggal Export: {datetime.now().strftime('%d-%m-%Y %H:%M')}"
        ]
        if search:
            metadata.insert(0, f"Filter: {search}")
        
        for meta in metadata:
            elements.append(Paragraph(meta, meta_style))
        
        # Prepare table data
        data = []
        headers = [
            "No", 
            "Perwakilan" if session.get('role') == 0 else None,  # Only show for admin
            "Nama", 
            "Tempat/Tgl Lahir", 
            "NIK", 
            "Telp/Email",
        ]
        # Filter out None values (for non-admin users)
        headers = [h for h in headers if h is not None]
        data.append(headers)
        
        for idx, item in enumerate(pegawai_list, 1):
            tempat_tgl_lahir = f"{item[2] if item[2] else '-'}, {item[3].strftime('%d-%m-%Y') if item[3] else '-'}"
            telp_email = f"{item[5] if item[5] else '-'} / {item[6] if item[6] else '-'}"
            
            row = [
                str(idx),
                item[10] or '-' if session.get('role') == 0 else None,  # Perwakilan (only for admin)
                item[1] or '-',   # Nama
                tempat_tgl_lahir,
                item[4] or '-',   # NIK
                telp_email,
            ]
            # Filter out None values (for non-admin users)
            row = [r for r in row if r is not None]
            data.append(row)
        
        # Calculate available width using A4 width
        available_width = A4[0] - doc.leftMargin - doc.rightMargin
        
        # Define column distribution - adjusted for A4 portrait
        if session.get('role') == 0:  # Admin
            col_ratios = [0.05, 0.15, 0.2, 0.15, 0.2, 0.25]
        else:  # Non-admin
            col_ratios = [0.05, 0.2, 0.15, 0.25, 0.35]
        
        col_widths = [available_width * ratio for ratio in col_ratios]
        
        # Ensure minimum widths
        if session.get('role') == 0:  # Admin
            min_widths = [25, 50, 60, 60, 80, 80]
        else:  # Non-admin
            min_widths = [25, 60, 60, 90, 100]
            
        col_widths = [max(w, min_w) for w, min_w in zip(col_widths, min_widths)]
        
        # Create table with adjusted column widths
        table = Table(data, 
                     colWidths=col_widths,
                     repeatRows=1,
                     hAlign='LEFT')
        
        # Table style with adjusted font size and padding
        style = TableStyle([
            # Header style
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#4472C4')),
            ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,0), 7),  # Reduced font size
            ('ALIGN', (0,0), (-1,0), 'CENTER'),
            ('BOTTOMPADDING', (0,0), (-1,0), 4),  # Reduced padding
            
            # Cell style
            ('FONTSIZE', (0,1), (-1,-1), 7),  # Reduced font size
            ('LEADING', (0,0), (-1,-1), 8),   # Reduced leading
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey),
            ('WORDWRAP', (0,0), (-1,-1), True),
            
            # Alternating row colors
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.whitesmoke]),
            
            # Specific column alignments
            ('ALIGN', (0,0), (0,-1), 'CENTER'),    # No column centered
            ('ALIGN', (1,0), (1,-1), 'LEFT'),      # Perwakilan left-aligned
            ('ALIGN', (2,0), (2,-1), 'LEFT'),      # Nama left-aligned
            ('ALIGN', (3,0), (3,-1), 'LEFT'),      # Tempat/Tgl Lahir left-aligned
            ('ALIGN', (4,0), (4,-1), 'LEFT'),      # NIK left-aligned
            ('ALIGN', (5,0), (5,-1), 'LEFT'),      # Telp/Email left-aligned
            ('ALIGN', (-2,0), (-2,-1), 'CENTER'),  # TMT Penempatan centered
            ('ALIGN', (-1,0), (-1,-1), 'CENTER')   # TMT Selesai centered
        ])
        
        table.setStyle(style)
        elements.append(table)
        
        # Build PDF
        doc.build(elements)
        buffer.seek(0)
        
        return send_file(
            buffer,
            as_attachment=True,
            download_name=f"data_pegawai_setempat_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
            mimetype='application/pdf'
        )
        
    except Exception as e:
        print(f"PDF Export Error: {str(e)}")
        flash('Gagal membuat PDF: ' + str(e), 'error')
        return redirect(url_for('list_pegawai_setempat'))

@app.route('/pegawai-setempat/export/excel')
def export_pegawai_setempat_excel():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    try:
        # Get all parameters from the request
        search = request.args.get('search', '').strip()
        sort_column = request.args.get('sort', 'no')
        sort_direction = request.args.get('dir', 'asc')

        # Validate sort column
        valid_columns = {
            'no': 'p.no',
            'nama': 'p.nama',
            'nik': 'p.nik',
            'tmt_penempatan': 'p.tmt_penempatan',
            'perwakilan': 'p.id_pwk'
        }
        sort_column = valid_columns.get(sort_column, 'p.no')
        sort_direction = 'DESC' if sort_direction.lower() == 'desc' else 'ASC'

        # Base query with joins
        query = """
            SELECT p.no, p.nama, p.t_lahir, p.tgl_lahir, p.nik, p.telp, p.email, 
                   p.tmt_penempatan, p.tmt_selesai_penempatan, p.id_pwk, r.nama_perwakilan
            FROM tabel_pegawai_setempat p
            LEFT JOIN ref_perwakilan r ON p.id_pwk = r.trigram
            WHERE (p.nama ILIKE %s OR 
                  p.nik ILIKE %s OR
                  p.telp ILIKE %s OR
                  p.email ILIKE %s OR
                  p.id_pwk ILIKE %s OR
                  r.nama_perwakilan ILIKE %s)
        """

        # Parameters for the query
        search_param = f'%{search}%'
        params = [search_param for _ in range(6)]

        # Add perwakilan filter for non-admin users
        if session.get('role') != 0:  # If not admin
            user_trigram = session.get('trigram')
            if user_trigram:
                query += " AND p.id_pwk = %s"
                params.append(user_trigram)

        # Sort by the selected column
        query += f" ORDER BY {sort_column} {sort_direction}"

        # Execute query to get all data
        pegawai_list = execute_query(query, params, fetch=True) or []

        # Create Excel file in memory
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet("Data Pegawai Setempat")
        
        # Add formats
        title_format = workbook.add_format({
            'bold': True,
            'font_size': 14,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        subtitle_format = workbook.add_format({
            'font_size': 10,
            'align': 'left',
            'valign': 'vcenter'
        })
        
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#4472C4',
            'font_color': '#FFFFFF',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True
        })
        
        # Data formats
        data_format = workbook.add_format({
            'border': 1,
            'text_wrap': True,
            'valign': 'top'
        })
        
        date_format = workbook.add_format({
            'border': 1,
            'text_wrap': True,
            'valign': 'top',
            'num_format': 'dd-mm-yyyy'
        })
        
        # Write title and metadata
        worksheet.merge_range('A1:H1', 'LAPORAN DATA PEGAWAI SETEMPAT', title_format)
        
        current_row = 1
        if search:
            worksheet.write(current_row, 0, f"Filter: {search}", subtitle_format)
            current_row += 1
            
        worksheet.write(current_row, 0, f"Tanggal Export: {datetime.now().strftime('%d-%m-%Y %H:%M')}", subtitle_format)
        current_row += 2
        
        # Write headers
        headers = [
            "No", 
            "Perwakilan", 
            "Nama", 
            "Tempat/Tgl Lahir", 
            "NIK", 
            "Telp/Email",
            "TMT Penempatan",
            "TMT Selesai"
        ]
        
        for col, header in enumerate(headers):
            worksheet.write(current_row, col, header, header_format)
        
        current_row += 1
        
        # Write data
        for idx, item in enumerate(pegawai_list, 1):
            tempat_tgl_lahir = f"{item[2] if item[2] else '-'}, {item[3].strftime('%d-%m-%Y') if item[3] else '-'}"
            telp_email = f"{item[5] if item[5] else '-'} / {item[6] if item[6] else '-'}"
            
            row = [
                idx,
                item[10] or '-',  # Perwakilan
                item[1] or '-',   # Nama
                tempat_tgl_lahir,
                item[4] or '-',   # NIK
                telp_email,
                item[7],  # TMT Penempatan (date object)
                item[8]   # TMT Selesai (date object)
            ]
            
            for col, value in enumerate(row):
                if col in [6, 7]:  # Date columns
                    if value:  # Only write if date exists
                        worksheet.write_datetime(current_row, col, value, date_format)
                    else:
                        worksheet.write(current_row, col, '-', data_format)
                else:
                    worksheet.write(current_row, col, value, data_format)
            
            current_row += 1
        
        # Auto-adjust column widths based on content
        for i, header in enumerate(headers):
            max_length = len(header)
            for row in pegawai_list:
                if i == 0:  # No column
                    val_length = len(str(row[0]))
                elif i == 1:  # Perwakilan
                    val_length = len(str(row[10] or ''))
                elif i == 2:  # Nama
                    val_length = len(str(row[1] or ''))
                elif i == 3:  # Tempat/Tgl Lahir
                    tempat = str(row[2] or '')
                    tgl = row[3].strftime('%d-%m-%Y') if row[3] else ''
                    val_length = len(f"{tempat}, {tgl}")
                elif i == 4:  # NIK
                    val_length = len(str(row[4] or ''))
                elif i == 5:  # Telp/Email
                    telp = str(row[5] or '')
                    email = str(row[6] or '')
                    val_length = len(f"{telp} / {email}")
                elif i == 6:  # TMT Penempatan
                    val_length = 10  # Fixed length for date format
                elif i == 7:  # TMT Selesai
                    val_length = 10  # Fixed length for date format
                
                if val_length > max_length:
                    max_length = val_length
            
            # Set column width with some padding
            worksheet.set_column(i, i, min(max_length + 2, 50))  # Max width 50 chars
        
        # Freeze header row
        worksheet.freeze_panes(1, 0)
        
        workbook.close()
        output.seek(0)
        
        return send_file(
            output,
            as_attachment=True,
            download_name=f"data_pegawai_setempat_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        print(f"Excel Export Error: {str(e)}")
        flash('Gagal membuat Excel: ' + str(e), 'error')
        return redirect(url_for('list_pegawai_setempat'))

@app.route('/pegawai-setempat')
def list_pegawai_setempat():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    # Get search and pagination parameters
    search = request.args.get('search', '').strip()
    page = request.args.get('page', 1, type=int)
    per_page = 20
    sort_column = request.args.get('sort', 'no')
    sort_direction = request.args.get('dir', 'asc')

    # Validate sort column
    valid_columns = {
        'no': 'p.no',
        'nama': 'p.nama',
        'nik': 'p.nik',
        'tmt_penempatan': 'p.tmt_penempatan',
        'perwakilan': 'p.id_pwk'
    }
    sort_column = valid_columns.get(sort_column, 'p.no')
    
    # Validate sort direction
    sort_direction = 'DESC' if sort_direction.lower() == 'desc' else 'ASC'

    # Base query
    query = """
        SELECT p.no, p.nama, p.t_lahir, p.tgl_lahir, p.nik, p.telp, p.email, 
               p.tmt_penempatan, p.tmt_selesai_penempatan, p.id_pwk, r.nama_perwakilan
        FROM tabel_pegawai_setempat p
        LEFT JOIN ref_perwakilan r ON p.id_pwk = r.trigram
        WHERE (p.nama ILIKE %s OR 
              p.nik ILIKE %s OR
              p.telp ILIKE %s OR
              p.email ILIKE %s OR
              p.id_pwk ILIKE %s OR
              r.nama_perwakilan ILIKE %s)
    """

    # Parameters for the query
    search_param = f'%{search}%'
    params = [search_param for _ in range(6)]

    # Add perwakilan filter for non-admin users
    if session.get('role') != 0:  # If not admin
        user_trigram = session.get('trigram')
        if user_trigram:
            query += " AND p.id_pwk = %s"
            params.append(user_trigram)

    # Add sorting
    query += f" ORDER BY {sort_column} {sort_direction}"

    # Count query for pagination
    count_query = query.replace(
        "SELECT p.no, p.nama, p.t_lahir, p.tgl_lahir, p.nik, p.telp, p.email, " +
        "p.tmt_penempatan, p.tmt_selesai_penempatan, p.id_pwk, r.nama_perwakilan", 
        "SELECT COUNT(*)"
    ).split("ORDER BY")[0]  # Remove ORDER BY for count query

    # Get total count - handle None result safely
    count_result = execute_query(count_query, params, fetch_one=True)
    total = count_result[0] if count_result else 0
    
    # Add pagination to main query
    paginated_query = query + f" LIMIT {per_page} OFFSET {(page - 1) * per_page}"
    
    # Execute query
    pegawai_list = execute_query(paginated_query, params, fetch=True) or []
    
    total_pages = (total + per_page - 1) // per_page if total > 0 else 1
    
    return render_template('pegawai_setempat/list.html',
                         pegawai_list=pegawai_list,
                         search=search,
                         page=page,
                         per_page=per_page,
                         total=total,
                         total_pages=total_pages,
                         sort_column=sort_column.replace('p.', ''),
                         sort_direction=sort_direction)

@app.route('/pegawai-setempat/create', methods=['GET', 'POST'])
def create_pegawai_setempat():
    if 'user_id' not in session:
        return redirect(url_for('login'))

    # Get perwakilan list based on user role
    if session.get('role') == 0:  # Admin
        perwakilan_list = execute_query(
            "SELECT trigram, nama_perwakilan FROM ref_perwakilan ORDER BY nama_perwakilan",
            fetch=True
        ) or []
    else:  # Regular user
        user_trigram = session.get('trigram')
        perwakilan_list = execute_query(
            "SELECT trigram, nama_perwakilan FROM ref_perwakilan WHERE trigram = %s",
            (user_trigram,),
            fetch=True
        ) or []

    if request.method == 'POST':
        try:
            data = (
                request.form.get('nama', '').strip(),
                request.form.get('t_lahir', '').strip(),
                request.form.get('tgl_lahir'),
                request.form.get('nik', '').strip(),
                request.form.get('telp', '').strip(),
                request.form.get('email', '').strip(),
                request.form.get('tmt_penempatan'),
                request.form.get('tmt_selesai_penempatan'),
                request.form.get('id_pwk', '').strip().upper()
            )

            # Validate required fields
            if not data[0]:  # nama
                flash('Nama wajib diisi', 'error')
                return render_template('pegawai_setempat/create.html', 
                                    perwakilan_list=perwakilan_list,
                                    form_data=request.form)

            # For non-admin users, verify they're not trying to create for another perwakilan
            if session.get('role') != 0 and data[8] != session.get('trigram'):
                flash('Anda hanya bisa membuat data untuk perwakilan Anda sendiri', 'error')
                return render_template('pegawai_setempat/create.html', 
                                    perwakilan_list=perwakilan_list,
                                    form_data=request.form)

            # Convert empty strings to None for date fields
            date_fields = [2, 6, 7]  # indices of date fields
            data = list(data)
            for i in date_fields:
                if data[i] == '':
                    data[i] = None

            success = execute_query("""
                INSERT INTO tabel_pegawai_setempat 
                (nama, t_lahir, tgl_lahir, nik, telp, email, 
                 tmt_penempatan, tmt_selesai_penempatan, id_pwk)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
            """, tuple(data), commit=True)
            
            if success:
                flash('Data pegawai setempat berhasil ditambahkan', 'success')
                return redirect(url_for('list_pegawai_setempat'))
            else:
                flash('Gagal menambahkan data pegawai setempat', 'error')
        except Exception as e:
            logger.error(f"Error creating pegawai setempat: {str(e)}")
            flash('Terjadi kesalahan saat menyimpan data', 'error')

    return render_template('pegawai_setempat/create.html', 
                         perwakilan_list=perwakilan_list)

@app.route('/pegawai-setempat/edit/<int:no>', methods=['GET', 'POST'])
def edit_pegawai_setempat(no):
    if 'user_id' not in session:
        return redirect(url_for('login'))

    # Get existing data with perwakilan info
    pegawai = execute_query(
        """SELECT p.no, p.nama, p.t_lahir, p.tgl_lahir, p.nik, p.telp, p.email, 
           p.tmt_penempatan, p.tmt_selesai_penempatan, p.id_pwk, r.nama_perwakilan
           FROM tabel_pegawai_setempat p
           LEFT JOIN ref_perwakilan r ON p.id_pwk = r.trigram
           WHERE p.no = %s""",
        (no,),
        fetch_one=True
    )
    
    if not pegawai:
        flash('Data pegawai setempat tidak ditemukan', 'error')
        return redirect(url_for('list_pegawai_setempat'))

    # Authorization check for non-admin users
    if session.get('role') != 0:  # If not admin
        user_trigram = session.get('trigram')
        if user_trigram != pegawai[9]:  # pegawai[9] is id_pwk
            flash('Anda tidak memiliki akses untuk mengedit data ini', 'error')
            return redirect(url_for('list_pegawai_setempat'))

    # Get perwakilan list based on user role
    if session.get('role') == 0:  # Admin
        perwakilan_list = execute_query(
            "SELECT trigram, nama_perwakilan FROM ref_perwakilan ORDER BY nama_perwakilan",
            fetch=True
        ) or []
    else:  # Regular user
        perwakilan_list = execute_query(
            "SELECT trigram, nama_perwakilan FROM ref_perwakilan WHERE trigram = %s",
            (session.get('trigram'),),
            fetch=True
        ) or []

    if request.method == 'POST':
        try:
            data = (
                request.form.get('nama', '').strip(),
                request.form.get('t_lahir', '').strip(),
                request.form.get('tgl_lahir'),
                request.form.get('nik', '').strip(),
                request.form.get('telp', '').strip(),
                request.form.get('email', '').strip(),
                request.form.get('tmt_penempatan'),
                request.form.get('tmt_selesai_penempatan'),
                request.form.get('id_pwk', '').strip().upper(),
                no
            )

            # Validate required fields
            if not data[0]:  # nama
                flash('Nama wajib diisi', 'error')
                return redirect(url_for('edit_pegawai_setempat', no=no))

            # Check if perwakilan exists
            perwakilan_exists = execute_query(
                "SELECT 1 FROM ref_perwakilan WHERE trigram = %s",
                (data[8],),
                fetch_one=True
            )
            
            if not perwakilan_exists:
                flash('Perwakilan tidak valid', 'error')
                return redirect(url_for('edit_pegawai_setempat', no=no))

            # Convert empty strings to None for date fields
            date_fields = [2, 6, 7]  # indices of date fields
            data = list(data)
            for i in date_fields:
                if data[i] == '':
                    data[i] = None

            success = execute_query("""
                UPDATE tabel_pegawai_setempat SET
                    nama = %s,
                    t_lahir = %s,
                    tgl_lahir = %s,
                    nik = %s,
                    telp = %s,
                    email = %s,
                    tmt_penempatan = %s,
                    tmt_selesai_penempatan = %s,
                    id_pwk = %s
                WHERE no = %s
            """, tuple(data), commit=True)
            
            if success:
                flash('Data pegawai setempat berhasil diperbarui', 'success')
                return redirect(url_for('list_pegawai_setempat'))
            else:
                flash('Gagal memperbarui data pegawai setempat', 'error')
        except Exception as e:
            logger.error(f"Error updating pegawai setempat: {str(e)}")
            flash('Terjadi kesalahan saat memperbarui data', 'error')

    return render_template('pegawai_setempat/edit.html', 
                         pegawai=pegawai,
                         perwakilan_list=perwakilan_list)

@app.route('/pegawai-setempat/delete/<int:no>', methods=['POST'])
def delete_pegawai_setempat(no):
    if 'user_id' not in session:
        return redirect(url_for('login'))

    # Authorization check
    if session.get('role') != 0:  # If not admin
        pegawai = execute_query(
            "SELECT id_pwk FROM tabel_pegawai_setempat WHERE no = %s",
            (no,),
            fetch_one=True
        )
        if not pegawai or pegawai[0] != session.get('trigram'):
            flash('Anda tidak memiliki akses untuk menghapus data ini', 'error')
            return redirect(url_for('list_pegawai_setempat'))

    try:
        success = execute_query(
            "DELETE FROM tabel_pegawai_setempat WHERE no = %s",
            (no,),
            commit=True
        )
        
        if success:
            flash('Data pegawai setempat berhasil dihapus', 'success')
        else:
            flash('Gagal menghapus data pegawai setempat', 'error')
    except Exception as e:
        logger.error(f"Error deleting pegawai setempat: {str(e)}")
        flash('Terjadi kesalahan saat menghapus data', 'error')
    
    return redirect(url_for('list_pegawai_setempat'))

# ==============================================
# JENIS PENDIDIKAN CRUD ROUTES
# ==============================================

@app.route('/jenis-pendidikan')
def list_jenis_pendidikan():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    # Get search and pagination parameters
    search = request.args.get('search', '').strip()
    page = request.args.get('page', 1, type=int)
    per_page = 20
    sort_column = request.args.get('sort', 'no')
    sort_direction = request.args.get('dir', 'asc')

    # Validate sort column
    valid_columns = {
        'no': 'no',
        'jenis_pend': 'jenis_pend'
    }
    sort_column = valid_columns.get(sort_column, 'no')
    
    # Validate sort direction
    sort_direction = 'DESC' if sort_direction.lower() == 'desc' else 'ASC'

    # Base query
    query = f"""
        SELECT no, jenis_pend
        FROM tabel_jenis_pendidikan
        WHERE jenis_pend ILIKE %s
        ORDER BY {sort_column} {sort_direction}
    """

    # Count query for pagination
    count_query = """
        SELECT COUNT(*)
        FROM tabel_jenis_pendidikan
        WHERE jenis_pend ILIKE %s
    """

    search_param = f'%{search}%'
    
    # Get total count
    total = execute_query(count_query, (search_param,), fetch_one=True)[0] or 0
    
    # Add pagination to main query
    paginated_query = query + f" LIMIT {per_page} OFFSET {(page - 1) * per_page}"
    
    # Execute query
    jenis_pendidikan_list = execute_query(paginated_query, (search_param,), fetch=True) or []
    
    total_pages = (total + per_page - 1) // per_page
    
    return render_template('jenis_pendidikan/list.html',
                         jenis_pendidikan_list=jenis_pendidikan_list,
                         search=search,
                         page=page,
                         per_page=per_page,
                         total=total,
                         total_pages=total_pages,
                         sort_column=sort_column,
                         sort_direction=sort_direction)

@app.route('/jenis-pendidikan/create', methods=['GET', 'POST'])
def create_jenis_pendidikan():
    if 'user_id' not in session:
        return redirect(url_for('login'))

    # Daftar pilihan untuk dropdown
    pilihan_jenis = ['umum', 'militer', 'pim', 'sandi', 'teknis']

    if request.method == 'POST':
        try:
            jenis_pend = request.form.get('jenis_pend', '').strip()

            # Validate required fields
            if not jenis_pend:
                flash('Jenis pendidikan wajib diisi', 'error')
                return render_template('jenis_pendidikan/create.html', 
                                     pilihan_jenis=pilihan_jenis,
                                     form_data=request.form)

            # Get next auto increment number
            next_no = execute_query("""
                SELECT COALESCE(MAX(no), 0) + 1 
                FROM tabel_jenis_pendidikan
            """, fetch_one=True)[0]

            success = execute_query("""
                INSERT INTO tabel_jenis_pendidikan 
                (no, jenis_pend)
                VALUES (%s, %s)
            """, (next_no, jenis_pend), commit=True)
            
            if success:
                flash('Data jenis pendidikan berhasil ditambahkan', 'success')
                return redirect(url_for('list_jenis_pendidikan'))
            else:
                flash('Gagal menambahkan data jenis pendidikan', 'error')
        except Exception as e:
            logger.error(f"Error creating jenis pendidikan: {str(e)}")
            flash('Terjadi kesalahan saat menyimpan data', 'error')

    return render_template('jenis_pendidikan/create.html', 
                         pilihan_jenis=pilihan_jenis)

@app.route('/jenis-pendidikan/edit/<int:no>', methods=['GET', 'POST'])
def edit_jenis_pendidikan(no):
    if 'user_id' not in session:
        return redirect(url_for('login'))

    # Daftar pilihan untuk dropdown
    pilihan_jenis = ['umum', 'militer', 'pim', 'sandi', 'teknis']

    # Get existing data
    jenis_pendidikan = execute_query(
        "SELECT no, jenis_pend FROM tabel_jenis_pendidikan WHERE no = %s",
        (no,),
        fetch_one=True
    )
    
    if not jenis_pendidikan:
        flash('Data jenis pendidikan tidak ditemukan', 'error')
        return redirect(url_for('list_jenis_pendidikan'))

    if request.method == 'POST':
        try:
            jenis_pend = request.form.get('jenis_pend', '').strip()

            # Validate required fields
            if not jenis_pend:
                flash('Jenis pendidikan wajib diisi', 'error')
                return redirect(url_for('edit_jenis_pendidikan', no=no))

            success = execute_query("""
                UPDATE tabel_jenis_pendidikan SET
                    jenis_pend = %s
                WHERE no = %s
            """, (jenis_pend, no), commit=True)
            
            if success:
                flash('Data jenis pendidikan berhasil diperbarui', 'success')
                return redirect(url_for('list_jenis_pendidikan'))
            else:
                flash('Gagal memperbarui data jenis pendidikan', 'error')
        except Exception as e:
            logger.error(f"Error updating jenis pendidikan: {str(e)}")
            flash('Terjadi kesalahan saat memperbarui data', 'error')

    return render_template('jenis_pendidikan/edit.html', 
                         jenis_pendidikan=jenis_pendidikan,
                         pilihan_jenis=pilihan_jenis)

@app.route('/jenis-pendidikan/delete/<int:no>', methods=['POST'])
def delete_jenis_pendidikan(no):
    if 'user_id' not in session:
        return redirect(url_for('login'))

    try:
        success = execute_query(
            "DELETE FROM tabel_jenis_pendidikan WHERE no = %s",
            (no,),
            commit=True
        )
        
        if success:
            flash('Data jenis pendidikan berhasil dihapus', 'success')
        else:
            flash('Gagal menghapus data jenis pendidikan', 'error')
    except Exception as e:
        logger.error(f"Error deleting jenis pendidikan: {str(e)}")
        flash('Terjadi kesalahan saat menghapus data', 'error')
    
    return redirect(url_for('list_jenis_pendidikan'))

# ==============================================
# PENDIDIKAN CRUD ROUTES
# ==============================================

@app.route('/pendidikan/export/pdf')
def export_pendidikan_pdf():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    try:
        # Get all parameters from the request
        search = request.args.get('search', '').strip()
        sort_column = request.args.get('sort', 'no')
        sort_direction = request.args.get('dir', 'asc')

        # Validate sort column
        valid_columns = {
            'no': 'p.no',
            'personel': 'per.nama',
            'jabatan': 'j.nama',
            'jenis': 'jp.jenis_pend',
            'tahun': 'p.tahun',
            'nama_pend': 'p.nama_pend'
        }
        sort_column = valid_columns.get(sort_column, 'p.no')
        sort_direction = 'DESC' if sort_direction.lower() == 'desc' else 'ASC'

        # Base query with joins
        query = """
            SELECT p.no, p.tahun, p.nama_pend,
                   per.no as id_personel, per.nama as nama_personel,
                   j.nama as nama_jabatan,
                   jp.jenis_pend,
                   r.nama_perwakilan as pwk
            FROM tabel_pendidikan p
            LEFT JOIN tabel_personel per ON p.id_personel = per.no
            LEFT JOIN ref_perwakilan r ON per.id_pwk = r.TRIGRAM
            LEFT JOIN tabel_jabatan j ON p.id_jabatan = j.no
            LEFT JOIN tabel_jenis_pendidikan jp ON p.id_jenis_pendidikan = jp.no
            WHERE (per.nama ILIKE %s OR 
                  j.nama ILIKE %s OR
                  jp.jenis_pend ILIKE %s OR
                  p.tahun::text ILIKE %s OR
                  p.nama_pend ILIKE %s)
        """

        # Parameters for the query
        search_param = f'%{search}%'
        params = [search_param for _ in range(5)]

        # Add perwakilan filter for non-admin users
        if session.get('role') != 0:  # If not admin
            user_trigram = session.get('trigram')
            if user_trigram:
                query += " AND per.id_pwk = %s"
                params.append(user_trigram.strip().upper())

        # Sort by the selected column
        query += f" ORDER BY {sort_column} {sort_direction}"

        # Execute query to get all data
        pendidikan_list = execute_query(query, params, fetch=True) or []

        # Create PDF with landscape orientation
        buffer = BytesIO()
        doc_width, doc_height = landscape(letter)
        doc = SimpleDocTemplate(buffer, 
                              pagesize=landscape(letter),
                              leftMargin=15,
                              rightMargin=15,
                              topMargin=20,
                              bottomMargin=20)
        elements = []
        
        # Styles configuration
        styles = getSampleStyleSheet()
        
        # Title style
        title_style = ParagraphStyle(
            'Title',
            parent=styles['Title'],
            fontSize=14,
            alignment=TA_CENTER,
            spaceAfter=12
        )
        
        # Metadata style
        meta_style = ParagraphStyle(
            'Meta',
            parent=styles['Normal'],
            fontSize=9,
            spaceAfter=6
        )
        
        # Add title
        elements.append(Paragraph("LAPORAN DATA PENDIDIKAN", title_style))
        
        # Add metadata
        metadata = [
            f"Tanggal Export: {datetime.now().strftime('%d-%m-%Y %H:%M')}"
        ]
        if search:
            metadata.insert(0, f"Filter: {search}")
        
        for meta in metadata:
            elements.append(Paragraph(meta, meta_style))
        
        # Prepare table data
        data = []
        headers = [
            "No", 
            "Perwakilan", 
            "Personel", 
            "Jabatan", 
            "Jenis Pendidikan", 
            "Tahun", 
            "Nama Pendidikan"
        ]
        data.append(headers)
        
        # Group data by personel
        grouped_data = {}
        for item in pendidikan_list:
            personel_id = item[3]  # id_personel
            if personel_id not in grouped_data:
                grouped_data[personel_id] = {
                    'pwk': item[7],        # nama_perwakilan
                    'personel': item[4],   # nama_personel
                    'jabatan': item[5],    # nama_jabatan
                    'pendidikan': []
                }
            grouped_data[personel_id]['pendidikan'].append(item)
        
        # Add rows to table
        row_num = 1
        for personel_id, personel_data in grouped_data.items():
            first_row = True
            for pendidikan in personel_data['pendidikan']:
                row = [
                    str(row_num) if first_row else '',
                    personel_data['pwk'] if first_row else '',
                    personel_data['personel'] if first_row else '',
                    personel_data['jabatan'] if first_row else '',
                    pendidikan[6],  # jenis_pend
                    str(pendidikan[1]),  # tahun
                    pendidikan[2]   # nama_pend
                ]
                data.append(row)
                if first_row:
                    first_row = False
                    row_num += 1
        
        # Calculate available width (subtract margins)
        available_width = doc_width - doc.leftMargin - doc.rightMargin
        
        # Define column distribution
        col_ratios = [0.05, 0.15, 0.15, 0.15, 0.15, 0.1, 0.25]  # Sum should be 1.0
        col_widths = [available_width * ratio for ratio in col_ratios]
        
        # Ensure minimum widths
        min_widths = [30, 60, 80, 60, 60, 40, 80]  # Minimum widths in points
        col_widths = [max(w, min_w) for w, min_w in zip(col_widths, min_widths)]
        
        # Create table
        table = Table(data, 
                     colWidths=col_widths,
                     repeatRows=1,
                     hAlign='LEFT')
        
        # Table style
        style = TableStyle([
            # Header style
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#4472C4')),
            ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,0), 8),
            ('ALIGN', (0,0), (-1,0), 'CENTER'),
            ('BOTTOMPADDING', (0,0), (-1,0), 6),
            
            # Cell style
            ('FONTSIZE', (0,1), (-1,-1), 8),
            ('LEADING', (0,0), (-1,-1), 9),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey),
            ('WORDWRAP', (0,0), (-1,-1), True),
            
            # Alternating row colors
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.whitesmoke]),
            
            # Specific column alignments
            ('ALIGN', (0,0), (0,-1), 'CENTER'),    # No column centered
            ('ALIGN', (1,0), (1,-1), 'LEFT'),      # Perwakilan left-aligned
            ('ALIGN', (2,0), (2,-1), 'LEFT'),      # Personel left-aligned
            ('ALIGN', (3,0), (3,-1), 'LEFT'),      # Jabatan left-aligned
            ('ALIGN', (4,0), (4,-1), 'LEFT'),      # Jenis Pendidikan left-aligned
            ('ALIGN', (5,0), (5,-1), 'CENTER'),    # Tahun centered
            ('ALIGN', (6,0), (6,-1), 'LEFT')       # Nama Pendidikan left-aligned
        ])
        
        # Add span for grouped rows
        row_idx = 1  # Start after header
        for personel_id, personel_data in grouped_data.items():
            span_count = len(personel_data['pendidikan'])
            if span_count > 1:
                # Span No, Perwakilan, Personel, and Jabatan columns
                for col in [0, 1, 2, 3]:
                    style.add('SPAN', (col, row_idx), (col, row_idx + span_count - 1))
            row_idx += span_count
        
        table.setStyle(style)
        elements.append(table)
        
        # Build PDF
        doc.build(elements)
        buffer.seek(0)
        
        return send_file(
            buffer,
            as_attachment=True,
            download_name=f"data_pendidikan_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
            mimetype='application/pdf'
        )
        
    except Exception as e:
        print(f"PDF Export Error: {str(e)}")
        flash('Gagal membuat PDF: ' + str(e), 'error')
        return redirect(url_for('list_pendidikan'))

@app.route('/pendidikan/export/excel')
def export_pendidikan_excel():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    try:
        # Get all parameters from the request
        search = request.args.get('search', '').strip()
        sort_column = request.args.get('sort', 'no')
        sort_direction = request.args.get('dir', 'asc')

        # Validate sort column
        valid_columns = {
            'no': 'p.no',
            'personel': 'per.nama',
            'jabatan': 'j.nama',
            'jenis': 'jp.jenis_pend',
            'tahun': 'p.tahun',
            'nama_pend': 'p.nama_pend'
        }
        sort_column = valid_columns.get(sort_column, 'p.no')
        sort_direction = 'DESC' if sort_direction.lower() == 'desc' else 'ASC'

        # Base query with joins
        query = """
            SELECT p.no, p.tahun, p.nama_pend,
                   per.no as id_personel, per.nama as nama_personel,
                   j.nama as nama_jabatan,
                   jp.jenis_pend,
                   r.nama_perwakilan as pwk
            FROM tabel_pendidikan p
            LEFT JOIN tabel_personel per ON p.id_personel = per.no
            LEFT JOIN ref_perwakilan r ON per.id_pwk = r.TRIGRAM
            LEFT JOIN tabel_jabatan j ON p.id_jabatan = j.no
            LEFT JOIN tabel_jenis_pendidikan jp ON p.id_jenis_pendidikan = jp.no
            WHERE (per.nama ILIKE %s OR 
                  j.nama ILIKE %s OR
                  jp.jenis_pend ILIKE %s OR
                  p.tahun::text ILIKE %s OR
                  p.nama_pend ILIKE %s)
        """

        # Parameters for the query
        search_param = f'%{search}%'
        params = [search_param for _ in range(5)]

        # Add perwakilan filter for non-admin users
        if session.get('role') != 0:  # If not admin
            user_trigram = session.get('trigram')
            if user_trigram:
                query += " AND per.id_pwk = %s"
                params.append(user_trigram.strip().upper())

        # Sort by the selected column
        query += f" ORDER BY {sort_column} {sort_direction}"

        # Execute query to get all data
        pendidikan_list = execute_query(query, params, fetch=True) or []

        # Create Excel file in memory
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet("Data Pendidikan")
        
        # Add formats
        title_format = workbook.add_format({
            'bold': True,
            'font_size': 14,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        subtitle_format = workbook.add_format({
            'font_size': 10,
            'align': 'left',
            'valign': 'vcenter'
        })
        
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#4472C4',
            'font_color': '#FFFFFF',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True
        })
        
        # Data formats
        data_format = workbook.add_format({
            'border': 1,
            'text_wrap': True,
            'valign': 'top'
        })
        
        center_format = workbook.add_format({
            'border': 1,
            'text_wrap': True,
            'valign': 'top',
            'align': 'center'
        })
        
        # Write title and metadata
        worksheet.merge_range('A1:G1', 'LAPORAN DATA PENDIDIKAN', title_format)
        
        current_row = 1
        if search:
            worksheet.write(current_row, 0, f"Filter: {search}", subtitle_format)
            current_row += 1
            
        worksheet.write(current_row, 0, f"Tanggal Export: {datetime.now().strftime('%d-%m-%Y %H:%M')}", subtitle_format)
        current_row += 2
        
        # Write headers
        headers = [
            "No", 
            "Perwakilan", 
            "Personel", 
            "Jabatan", 
            "Jenis Pendidikan", 
            "Tahun", 
            "Nama Pendidikan"
        ]
        
        for col, header in enumerate(headers):
            worksheet.write(current_row, col, header, header_format)
        
        current_row += 1
        
        # Group data by personel
        grouped_data = {}
        for item in pendidikan_list:
            personel_id = item[3]  # id_personel
            if personel_id not in grouped_data:
                grouped_data[personel_id] = {
                    'pwk': item[7],        # nama_perwakilan
                    'personel': item[4],   # nama_personel
                    'jabatan': item[5],    # nama_jabatan
                    'pendidikan': []
                }
            grouped_data[personel_id]['pendidikan'].append(item)
        
        # Write data
        row_num = 1
        for personel_id, personel_data in grouped_data.items():
            first_row = True
            for pendidikan in personel_data['pendidikan']:
                # Apply different formats based on column
                worksheet.write(current_row, 0, row_num if first_row else '', center_format)  # No
                worksheet.write(current_row, 1, personel_data['pwk'] if first_row else '', data_format)  # Perwakilan
                worksheet.write(current_row, 2, personel_data['personel'] if first_row else '', data_format)  # Personel
                worksheet.write(current_row, 3, personel_data['jabatan'] if first_row else '', data_format)  # Jabatan
                worksheet.write(current_row, 4, pendidikan[6], data_format)  # Jenis Pendidikan
                worksheet.write(current_row, 5, str(pendidikan[1]), center_format)  # Tahun - centered
                worksheet.write(current_row, 6, pendidikan[2], data_format)  # Nama Pendidikan
                
                if first_row:
                    first_row = False
                    row_num += 1
                current_row += 1
        
        # Auto-adjust column widths based on content
        col_widths = [10, 20, 25, 20, 20, 10, 30]  # Initial widths
        for i, header in enumerate(headers):
            max_length = len(header)
            for personel_id, personel_data in grouped_data.items():
                if i == 0:  # No
                    val_length = len(str(row_num))
                elif i == 1:  # Perwakilan
                    val_length = len(personel_data['pwk'])
                elif i == 2:  # Personel
                    val_length = len(personel_data['personel'])
                elif i == 3:  # Jabatan
                    val_length = len(personel_data['jabatan'] or '')
                elif i == 4:  # Jenis Pendidikan
                    max_pend = max(len(p[6]) for p in personel_data['pendidikan'])
                    val_length = max_pend if 'max_pend' in locals() else 0
                elif i == 5:  # Tahun
                    val_length = 4  # Year length
                elif i == 6:  # Nama Pendidikan
                    max_nama = max(len(p[2]) for p in personel_data['pendidikan'])
                    val_length = max_nama if 'max_nama' in locals() else 0
                
                if val_length > max_length:
                    max_length = val_length
            
            # Set column width with some padding
            worksheet.set_column(i, i, min(max_length + 2, 50))  # Max width 50 chars
        
        workbook.close()
        output.seek(0)
        
        return send_file(
            output,
            as_attachment=True,
            download_name=f"data_pendidikan_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        print(f"Excel Export Error: {str(e)}")
        flash('Gagal membuat Excel: ' + str(e), 'error')
        return redirect(url_for('list_pendidikan'))

@app.route('/pendidikan')
def list_pendidikan():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    # Get search and pagination parameters
    search = request.args.get('search', '').strip()
    page = request.args.get('page', 1, type=int)
    per_page = 20
    sort_column = request.args.get('sort', 'no')
    sort_direction = request.args.get('dir', 'asc')

    # Validate sort column
    valid_columns = {
        'no': 'p.no',
        'personel': 'per.nama',
        'jabatan': 'j.nama',
        'jenis': 'jp.jenis_pend',
        'tahun': 'p.tahun',
        'nama_pend': 'p.nama_pend'
    }
    sort_column = valid_columns.get(sort_column, 'p.no')
    
    # Validate sort direction
    sort_direction = 'DESC' if sort_direction.lower() == 'desc' else 'ASC'

    # Base query with joins
    query = f"""
        SELECT p.no, p.tahun, p.nama_pend,
               per.no as id_personel, per.nama as nama_personel,
               j.no as id_jabatan, j.nama as nama_jabatan,
               jp.no as id_jenis_pendidikan, jp.jenis_pend,
               per.id_pwk as id_pwk, 
			   r.nama_perwakilan as pwk
        FROM tabel_pendidikan p
        LEFT JOIN tabel_personel per ON p.id_personel = per.no
        LEFT JOIN ref_perwakilan r ON id_pwk = r.TRIGRAM
        LEFT JOIN tabel_jabatan j ON p.id_jabatan = j.no
        LEFT JOIN tabel_jenis_pendidikan jp ON p.id_jenis_pendidikan = jp.no
        WHERE (per.nama ILIKE %s OR 
              j.nama ILIKE %s OR
              jp.jenis_pend ILIKE %s OR
              p.tahun::text ILIKE %s OR
              p.nama_pend ILIKE %s)
    """

    # Parameters for the query
    search_param = f'%{search}%'
    params = [search_param for _ in range(5)]

    # Add perwakilan filter for non-admin users
    if session.get('role') != 0:  # If not admin
        user_trigram = session.get('trigram')
        if user_trigram:
            query += " AND per.id_pwk = %s"
            params.append(user_trigram.strip().upper())

    # Add sorting
    query += f" ORDER BY {sort_column} {sort_direction}"

    # Count query for pagination
    count_query = query.replace(
        "SELECT p.no, p.tahun, p.nama_pend, per.no as id_personel, per.nama as nama_personel, j.no as id_jabatan, j.nama as nama_jabatan, jp.no as id_jenis_pendidikan, jp.jenis_pend",
        "SELECT COUNT(*)"
    ).split("ORDER BY")[0]  # Remove ORDER BY for count query

    # Get total count
    total_result = execute_query(count_query, params, fetch_one=True)
    total = total_result[0] if total_result else 0
    
    # Add pagination to main query
    paginated_query = query + f" LIMIT {per_page} OFFSET {(page - 1) * per_page}"
    
    # Execute query
    pendidikan_list = execute_query(paginated_query, params, fetch=True) or []
    
    total_pages = max(1, (total + per_page - 1) // per_page)
    
    return render_template('pendidikan/list.html',
                         pendidikan_list=pendidikan_list,
                         search=search,
                         page=page,
                         per_page=per_page,
                         total=total,
                         total_pages=total_pages,
                         sort_column=sort_column.replace('p.', ''),
                         sort_direction=sort_direction)

@app.route('/pendidikan/create', methods=['GET', 'POST'])
def create_pendidikan():
    if 'user_id' not in session:
        return redirect(url_for('login'))

    # Get dropdown data - MODIFIED to filter by perwakilan
    personel_query = "SELECT no, nama FROM tabel_personel"
    jabatan_query = "SELECT no, nama FROM tabel_jabatan ORDER BY nama"
    jenis_pend_query = "SELECT no, jenis_pend FROM tabel_jenis_pendidikan ORDER BY jenis_pend"
    
    params = []
    
    # Filter personel by perwakilan for non-admin users
    if session.get('role') != 0:  # If not admin
        user_trigram = session.get('trigram')
        if user_trigram:
            personel_query += " WHERE UPPER(TRIM(id_pwk)) = UPPER(TRIM(%s))"
            params.append(user_trigram.strip().upper())

    personel_query += " ORDER BY nama"
    
    personel_list = execute_query(personel_query, params, fetch=True) or []
    jabatan_list = execute_query(jabatan_query, fetch=True) or []
    jenis_pend_list = execute_query(jenis_pend_query, fetch=True) or []

    if request.method == 'POST':
        try:
            id_personel = request.form.get('id_personel')
            if not id_personel:
                flash('Personel harus dipilih', 'error')
                return render_template('pendidikan/create.html', 
                                    personel_list=personel_list,
                                    jabatan_list=jabatan_list,
                                    jenis_pend_list=jenis_pend_list,
                                    form_data=request.form)

            # Authorization check - ensure the selected personel belongs to user's perwakilan
            if session.get('role') != 0:  # If not admin
                user_trigram = session.get('trigram')
                if user_trigram:
                    personel_perwakilan = execute_query(
                        "SELECT id_pwk FROM tabel_personel WHERE no = %s",
                        (id_personel,),
                        fetch_one=True
                    )
                    if not personel_perwakilan or personel_perwakilan[0].strip().upper() != user_trigram.strip().upper():
                        flash('Anda hanya bisa menambahkan data untuk personel di perwakilan Anda', 'error')
                        return redirect(url_for('create_pendidikan'))

            # [Rest of your existing POST handling code...]

            # Get all pendidikan entries
            jenis_pendidikan_list = request.form.getlist('jenis_pendidikan[]')
            tahun_list = request.form.getlist('tahun[]')
            nama_pendidikan_list = request.form.getlist('nama_pendidikan[]')
            jabatan_list_form = request.form.getlist('jabatan[]')

            # Validate all entries
            valid_entries = 0
            for i in range(len(jenis_pendidikan_list)):
                if not all([jenis_pendidikan_list[i], tahun_list[i], nama_pendidikan_list[i]]):
                    flash(f'Data pendidikan #{i+1} tidak lengkap', 'error')
                    continue
                valid_entries += 1

            if valid_entries == 0:
                flash('Minimal satu data pendidikan yang valid harus diisi', 'error')
                return render_template('pendidikan/create.html', 
                                    personel_list=personel_list,
                                    jabatan_list=jabatan_list,
                                    jenis_pend_list=jenis_pend_list,
                                    form_data=request.form)

            # Process all valid entries
            success_count = 0
            for i in range(len(jenis_pendidikan_list)):
                if not all([jenis_pendidikan_list[i], tahun_list[i], nama_pendidikan_list[i]]):
                    continue

                id_jabatan = int(jabatan_list_form[i]) if jabatan_list_form[i] else None

                success = execute_query("""
                    INSERT INTO tabel_pendidikan 
                    (id_personel, id_jabatan, id_jenis_pendidikan, tahun, nama_pend)
                    VALUES (%s, %s, %s, %s, %s)
                """, (int(id_personel), id_jabatan, int(jenis_pendidikan_list[i]), 
                    tahun_list[i], nama_pendidikan_list[i]), 
                    commit=True)
                
                if success:
                    success_count += 1

            if success_count > 0:
                flash(f'Berhasil menambahkan {success_count} data pendidikan', 'success')
                return redirect(url_for('list_pendidikan'))
            else:
                flash('Gagal menambahkan data pendidikan', 'error')
        except Exception as e:
            logger.error(f"Error creating pendidikan: {str(e)}")
            flash('Terjadi kesalahan saat menyimpan data', 'error')

    return render_template('pendidikan/create.html', 
                         personel_list=personel_list,
                         jabatan_list=jabatan_list,
                         jenis_pend_list=jenis_pend_list)

@app.route('/pendidikan/edit/<int:personel_id>', methods=['GET', 'POST'])
def edit_pendidikan(personel_id):
    if 'user_id' not in session:
        return redirect(url_for('login'))

    # Authorization check - ensure the personel belongs to user's perwakilan
    if session.get('role') != 0:  # If not admin
        user_trigram = session.get('trigram')
        if user_trigram:
            personel_perwakilan = execute_query(
                "SELECT id_pwk FROM tabel_personel WHERE no = %s",
                (personel_id,),
                fetch_one=True
            )
            if not personel_perwakilan or personel_perwakilan[0].strip().upper() != user_trigram.strip().upper():
                flash('Anda tidak memiliki akses untuk mengedit data ini', 'error')
                return redirect(url_for('list_pendidikan'))

    # Get dropdown data - MODIFIED to filter by perwakilan
    personel_query = "SELECT no, nama FROM tabel_personel"
    jabatan_query = "SELECT no, nama FROM tabel_jabatan ORDER BY nama"
    jenis_pend_query = "SELECT no, jenis_pend FROM tabel_jenis_pendidikan ORDER BY jenis_pend"
    
    params = []
    
    # Filter personel by perwakilan for non-admin users
    if session.get('role') != 0:  # If not admin
        user_trigram = session.get('trigram')
        if user_trigram:
            personel_query += " WHERE UPPER(TRIM(id_pwk)) = UPPER(TRIM(%s))"
            params.append(user_trigram.strip().upper())

    personel_query += " ORDER BY nama"
    
    personel_list = execute_query(personel_query, params, fetch=True) or []
    jabatan_list = execute_query(jabatan_query, fetch=True) or []
    jenis_pend_list = execute_query(jenis_pend_query, fetch=True) or []


    # Get all pendidikan entries for this personel
    pendidikan_list = execute_query(
        """SELECT p.no, p.id_personel, p.id_jabatan, p.id_jenis_pendidikan, 
                  p.tahun, p.nama_pend,
                  per.nama as nama_personel,
                  j.nama as nama_jabatan,
                  jp.jenis_pend
           FROM tabel_pendidikan p
           LEFT JOIN tabel_personel per ON p.id_personel = per.no
           LEFT JOIN tabel_jabatan j ON p.id_jabatan = j.no
           LEFT JOIN tabel_jenis_pendidikan jp ON p.id_jenis_pendidikan = jp.no
           WHERE p.id_personel = %s
           ORDER BY p.tahun""",
        (personel_id,),
        fetch=True
    )
    
    if not pendidikan_list:
        flash('Data pendidikan tidak ditemukan', 'error')
        return redirect(url_for('list_pendidikan'))

    if request.method == 'POST':
        try:
            # Process updates and deletions
            updated = 0
            deleted = 0
            
            # Get all existing IDs for this personel
            existing_ids = [str(p[0]) for p in pendidikan_list]
            
            # Process updates for existing entries
            for pendidikan in pendidikan_list:
                pid = str(pendidikan[0])
                if f"pendidikan_{pid}_update" in request.form:
                    data = {
                        'id_jabatan': request.form.get(f"pendidikan_{pid}_jabatan"),
                        'id_jenis_pendidikan': request.form.get(f"pendidikan_{pid}_jenis"),
                        'tahun': request.form.get(f"pendidikan_{pid}_tahun"),
                        'nama_pend': request.form.get(f"pendidikan_{pid}_nama"),
                        'no': pid
                    }

                    # Validate required fields
                    if not all([data['id_jenis_pendidikan'], data['tahun'], data['nama_pend']]):
                        flash(f'Data tahun {data["tahun"]} tidak lengkap', 'error')
                        continue

                    # Convert empty jabatan to None
                    id_jabatan = int(data['id_jabatan']) if data['id_jabatan'] else None

                    success = execute_query("""
                        UPDATE tabel_pendidikan SET
                            id_jabatan = %s,
                            id_jenis_pendidikan = %s,
                            tahun = %s,
                            nama_pend = %s
                        WHERE no = %s
                    """, (id_jabatan, int(data['id_jenis_pendidikan']), 
                         data['tahun'], data['nama_pend'], pid),
                        commit=True)
                    
                    if success:
                        updated += 1

            # Process deletions
            for pid in existing_ids:
                if f"pendidikan_{pid}_delete" in request.form:
                    success = execute_query(
                        "DELETE FROM tabel_pendidikan WHERE no = %s",
                        (pid,),
                        commit=True
                    )
                    if success:
                        deleted += 1

            # Process new entries
            new_entries = int(request.form.get('new_entry_count', 0))
            added = 0
            for i in range(1, new_entries + 1):
                if f"new_{i}_tahun" in request.form:
                    data = {
                        'id_personel': personel_id,
                        'id_jabatan': request.form.get(f"new_{i}_jabatan"),
                        'id_jenis_pendidikan': request.form.get(f"new_{i}_jenis"),
                        'tahun': request.form.get(f"new_{i}_tahun"),
                        'nama_pend': request.form.get(f"new_{i}_nama")
                    }

                    if not all([data['id_jenis_pendidikan'], data['tahun'], data['nama_pend']]):
                        continue

                    id_jabatan = int(data['id_jabatan']) if data['id_jabatan'] else None

                    success = execute_query("""
                        INSERT INTO tabel_pendidikan 
                        (id_personel, id_jabatan, id_jenis_pendidikan, tahun, nama_pend)
                        VALUES (%s, %s, %s, %s, %s)
                    """, (personel_id, id_jabatan, int(data['id_jenis_pendidikan']), 
                        data['tahun'], data['nama_pend']), 
                        commit=True)
                    
                    if success:
                        added += 1

            flash(f'Berhasil: {updated} data diperbarui, {deleted} data dihapus, {added} data baru ditambahkan', 'success')
            return redirect(url_for('list_pendidikan'))

        except Exception as e:
            logger.error(f"Error updating pendidikan: {str(e)}")
            flash('Terjadi kesalahan saat memperbarui data', 'error')

    return render_template('pendidikan/edit.html', 
                        personel_id=personel_id,
                        pendidikan_list=pendidikan_list,
                        personel_list=personel_list,
                        jabatan_list=jabatan_list,
                        jenis_pend_list=jenis_pend_list)

@app.route('/pendidikan/delete/<int:no>', methods=['POST'])
def delete_pendidikan(no):
    if 'user_id' not in session:
        return redirect(url_for('login'))

    try:
        success = execute_query(
            "DELETE FROM tabel_pendidikan WHERE no = %s",
            (no,),
            commit=True
        )
        
        if success:
            flash('Data pendidikan berhasil dihapus', 'success')
        else:
            flash('Gagal menghapus data pendidikan', 'error')
    except Exception as e:
        logger.error(f"Error deleting pendidikan: {str(e)}")
        flash('Terjadi kesalahan saat menghapus data', 'error')
    
    return redirect(url_for('list_pendidikan'))

# ==============================================
# JENIS FUNGSIONAL CRUD ROUTES
# ==============================================

@app.route('/jenis-fungsional')
def list_jenis_fungsional():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    # Get search and pagination parameters
    search = request.args.get('search', '').strip()
    page = request.args.get('page', 1, type=int)
    per_page = 20
    sort_column = request.args.get('sort', 'no')
    sort_direction = request.args.get('dir', 'asc')

    # Validate sort column
    valid_columns = {
        'no': 'no',
        'nama_fungsional': 'nama_fungsional'
    }
    sort_column = valid_columns.get(sort_column, 'no')
    
    # Validate sort direction
    sort_direction = 'DESC' if sort_direction.lower() == 'desc' else 'ASC'

    # Base query
    query = f"""
        SELECT no, nama_fungsional
        FROM tabel_jenis_fungsional
        WHERE nama_fungsional ILIKE %s
        ORDER BY {sort_column} {sort_direction}
    """

    # Count query for pagination
    count_query = """
        SELECT COUNT(*)
        FROM tabel_jenis_fungsional
        WHERE nama_fungsional ILIKE %s
    """

    search_param = f'%{search}%'
    
    # Get total count
    total = execute_query(count_query, (search_param,), fetch_one=True)[0] or 0
    
    # Add pagination to main query
    paginated_query = query + f" LIMIT {per_page} OFFSET {(page - 1) * per_page}"
    
    # Execute query
    jenis_fungsional_list = execute_query(paginated_query, (search_param,), fetch=True) or []
    
    total_pages = (total + per_page - 1) // per_page
    
    return render_template('jenis_fungsional/list.html',
                         jenis_fungsional_list=jenis_fungsional_list,
                         search=search,
                         page=page,
                         per_page=per_page,
                         total=total,
                         total_pages=total_pages,
                         sort_column=sort_column,
                         sort_direction=sort_direction)

@app.route('/jenis-fungsional/create', methods=['GET', 'POST'])
def create_jenis_fungsional():
    if 'user_id' not in session:
        return redirect(url_for('login'))

    if request.method == 'POST':
        try:
            nama_fungsional = request.form.get('nama_fungsional', '').strip()

            # Validate required fields
            if not nama_fungsional:
                flash('Nama fungsional wajib diisi', 'error')
                return render_template('jenis_fungsional/create.html',
                                     form_data=request.form)

            # Get next auto increment number
            next_no = execute_query("""
                SELECT COALESCE(MAX(no), 0) + 1 
                FROM tabel_jenis_fungsional
            """, fetch_one=True)[0]

            success = execute_query("""
                INSERT INTO tabel_jenis_fungsional 
                (no, nama_fungsional)
                VALUES (%s, %s)
            """, (next_no, nama_fungsional), commit=True)
            
            if success:
                flash('Data jenis fungsional berhasil ditambahkan', 'success')
                return redirect(url_for('list_jenis_fungsional'))
            else:
                flash('Gagal menambahkan data jenis fungsional', 'error')
        except Exception as e:
            logger.error(f"Error creating jenis fungsional: {str(e)}")
            flash('Terjadi kesalahan saat menyimpan data', 'error')

    return render_template('jenis_fungsional/create.html')

@app.route('/jenis-fungsional/edit/<int:no>', methods=['GET', 'POST'])
def edit_jenis_fungsional(no):
    if 'user_id' not in session:
        return redirect(url_for('login'))

    # Get existing data
    jenis_fungsional = execute_query(
        "SELECT no, nama_fungsional FROM tabel_jenis_fungsional WHERE no = %s",
        (no,),
        fetch_one=True
    )
    
    if not jenis_fungsional:
        flash('Data jenis fungsional tidak ditemukan', 'error')
        return redirect(url_for('list_jenis_fungsional'))

    if request.method == 'POST':
        try:
            nama_fungsional = request.form.get('nama_fungsional', '').strip()

            # Validate required fields
            if not nama_fungsional:
                flash('Nama fungsional wajib diisi', 'error')
                return redirect(url_for('edit_jenis_fungsional', no=no))

            success = execute_query("""
                UPDATE tabel_jenis_fungsional SET
                    nama_fungsional = %s
                WHERE no = %s
            """, (nama_fungsional, no), commit=True)
            
            if success:
                flash('Data jenis fungsional berhasil diperbarui', 'success')
                return redirect(url_for('list_jenis_fungsional'))
            else:
                flash('Gagal memperbarui data jenis fungsional', 'error')
        except Exception as e:
            logger.error(f"Error updating jenis fungsional: {str(e)}")
            flash('Terjadi kesalahan saat memperbarui data', 'error')

    return render_template('jenis_fungsional/edit.html', 
                         jenis_fungsional=jenis_fungsional)

@app.route('/jenis-fungsional/delete/<int:no>', methods=['POST'])
def delete_jenis_fungsional(no):
    if 'user_id' not in session:
        return redirect(url_for('login'))

    try:
        success = execute_query(
            "DELETE FROM tabel_jenis_fungsional WHERE no = %s",
            (no,),
            commit=True
        )
        
        if success:
            flash('Data jenis fungsional berhasil dihapus', 'success')
        else:
            flash('Gagal menghapus data jenis fungsional', 'error')
    except Exception as e:
        logger.error(f"Error deleting jenis fungsional: {str(e)}")
        flash('Terjadi kesalahan saat menghapus data', 'error')
    
    return redirect(url_for('list_jenis_fungsional'))

# ==============================================
# FUNGSIONAL CRUD ROUTES
# ==============================================

@app.route('/fungsional/export/pdf')
def export_fungsional_pdf():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    try:
        # Get all parameters from the request
        search = request.args.get('search', '').strip()
        sort_column = request.args.get('sort', 'no')
        sort_direction = request.args.get('dir', 'asc')

        # Validate sort column
        valid_columns = {
            'no': 'f.no',
            'nama_fungsional': 'jf.nama_fungsional',
            'nama_pendidikan': 'p.nama_pend',
            'tahun_pendidikan': 'p.tahun',
            'jenjang': 'f.jenjang',
            'tmt_jenjang': 'f.tmt_jenjang'
        }
        sort_column = valid_columns.get(sort_column, 'f.no')
        sort_direction = 'DESC' if sort_direction.lower() == 'desc' else 'ASC'

        # Base query with joins
        query = """
            SELECT f.no, jf.nama_fungsional, p.nama_pend, p.tahun, 
                   f.jenjang, f.tmt_jenjang, f.no_sk, per.nama as nama_personel, 
                   per.id_pwk as id_pwk, pwk.nama_perwakilan
            FROM tabel_fungsional f
            LEFT JOIN tabel_jenis_fungsional jf ON f.nama_fungsional = jf.no
            LEFT JOIN tabel_pendidikan p ON f.nama_pendidikan = p.no
            LEFT JOIN tabel_personel per ON f.id_personel = per.no
            LEFT JOIN ref_perwakilan pwk ON pwk.trigram = per.id_pwk
            WHERE (jf.nama_fungsional ILIKE %s OR p.nama_pend ILIKE %s OR per.nama ILIKE %s)
        """

        # Parameters for the query
        search_param = f'%{search}%'
        params = [search_param for _ in range(3)]

        # Add perwakilan filter for non-admin users
        if session.get('role') != 0:  # If not admin
            user_trigram = session.get('trigram')
            if user_trigram:
                query += " AND UPPER(TRIM(per.id_pwk)) = UPPER(TRIM(%s))"
                params.append(user_trigram.strip().upper())

        # Sort by the selected column
        query += f" ORDER BY {sort_column} {sort_direction}"

        # Execute query to get all data
        fungsional_list = execute_query(query, params, fetch=True) or []

        # Create PDF with A4 size (portrait orientation)
        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, 
                              pagesize=A4,
                              leftMargin=15,
                              rightMargin=15,
                              topMargin=20,
                              bottomMargin=20)
        elements = []
        
        # Styles configuration
        styles = getSampleStyleSheet()
        
        # Title style
        title_style = ParagraphStyle(
            'Title',
            parent=styles['Title'],
            fontSize=14,
            alignment=TA_CENTER,
            spaceAfter=12
        )
        
        # Metadata style
        meta_style = ParagraphStyle(
            'Meta',
            parent=styles['Normal'],
            fontSize=9,
            spaceAfter=6
        )
        
        # Add title
        elements.append(Paragraph("LAPORAN DATA FUNGSIONAL", title_style))
        
        # Add metadata
        metadata = [
            f"Tanggal Export: {datetime.now().strftime('%d-%m-%Y %H:%M')}"
        ]
        if search:
            metadata.insert(0, f"Filter: {search}")
        
        for meta in metadata:
            elements.append(Paragraph(meta, meta_style))
        
        # Prepare table data
        data = []
        headers = [
            "No", 
            "Perwakilan" if session.get('role') == 0 else None,  # Only show for admin
            "Personel", 
            "Fungsional", 
            "Pendidikan", 
            "Tahun", 
            "Jenjang",
            "No. SK"
        ]
        # Filter out None values (for non-admin users)
        headers = [h for h in headers if h is not None]
        data.append(headers)
        
        for idx, item in enumerate(fungsional_list, 1):
            tmt_jenjang = item[5].strftime('%d-%m-%Y') if item[5] else '-'
            row = [
                str(idx),
                item[9] or '-' if session.get('role') == 0 else None,  # Perwakilan (only for admin)
                item[7] or '-',  # Personel
                item[1] or '-',  # Fungsional
                item[2] or '-',  # Pendidikan
                str(item[3]) if item[3] else '-',  # Tahun
                item[4] or '-',  # Jenjang
                item[6] or '-'   # No. SK
            ]
            # Filter out None values (for non-admin users)
            row = [r for r in row if r is not None]
            data.append(row)
        
        # Calculate available width using A4 width
        available_width = A4[0] - doc.leftMargin - doc.rightMargin
        
        # Define column distribution - adjusted for A4 portrait
        if session.get('role') == 0:  # Admin
            col_ratios = [0.05, 0.15, 0.15, 0.18, 0.15, 0.08, 0.1, 0.15]
        else:  # Non-admin

            col_ratios = [0.05, 0.25, 0.2, 0.18, 0.08, 0.1, 0.15]
        
        col_widths = [available_width * ratio for ratio in col_ratios]
        
        # Ensure minimum widths
        if session.get('role') == 0:  # Admin
            min_widths = [25, 50, 60, 70, 40, 40, 50, 40]
        else:  # Non-admin
            min_widths = [25, 80, 70, 40, 40, 50, 40]
            
        col_widths = [max(w, min_w) for w, min_w in zip(col_widths, min_widths)]
        
        # Create table with adjusted column widths
        table = Table(data, 
                     colWidths=col_widths,
                     repeatRows=1,
                     hAlign='LEFT')
        
        # Table style with adjusted font size and padding
        style = TableStyle([
            # Header style
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#4472C4')),
            ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,0), 7),  # Reduced font size
            ('ALIGN', (0,0), (-1,0), 'CENTER'),
            ('BOTTOMPADDING', (0,0), (-1,0), 4),  # Reduced padding
            
            # Cell style
            ('FONTSIZE', (0,1), (-1,-1), 7),  # Reduced font size
            ('LEADING', (0,0), (-1,-1), 8),   # Reduced leading
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey),
            ('WORDWRAP', (0,0), (-1,-1), True),
            
            # Alternating row colors
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.whitesmoke]),
            
            # Specific column alignments
            ('ALIGN', (0,0), (0,-1), 'CENTER'),    # No column centered
            ('ALIGN', (1,0), (1,-1), 'LEFT'),      # Perwakilan left-aligned
            ('ALIGN', (2,0), (2,-1), 'LEFT'),      # Personel left-aligned
            ('ALIGN', (3,0), (3,-1), 'LEFT'),      # Fungsional left-aligned
            ('ALIGN', (4,0), (4,-1), 'LEFT'),      # Pendidikan left-aligned
            ('ALIGN', (5,0), (5,-1), 'CENTER'),    # Tahun centered
            ('ALIGN', (6,0), (6,-1), 'CENTER'),    # Jenjang centered
            ('ALIGN', (7,0), (7,-1), 'CENTER'),    # TMT Jenjang centered
            ('ALIGN', (8,0), (8,-1), 'CENTER')     # No. SK centered
        ])
        
        table.setStyle(style)
        elements.append(table)
        
        # Build PDF
        doc.build(elements)
        buffer.seek(0)
        
        return send_file(
            buffer,
            as_attachment=True,
            download_name=f"data_fungsional_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
            mimetype='application/pdf'
        )
        
    except Exception as e:
        print(f"PDF Export Error: {str(e)}")
        flash('Gagal membuat PDF: ' + str(e), 'error')
        return redirect(url_for('list_fungsional'))

@app.route('/fungsional/export/excel')
def export_fungsional_excel():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    try:
        # Get all parameters from the request
        search = request.args.get('search', '').strip()
        sort_column = request.args.get('sort', 'no')
        sort_direction = request.args.get('dir', 'asc')

        # Validate sort column
        valid_columns = {
            'no': 'f.no',
            'nama_fungsional': 'jf.nama_fungsional',
            'nama_pendidikan': 'p.nama_pend',
            'tahun_pendidikan': 'p.tahun',
            'jenjang': 'f.jenjang',
            'tmt_jenjang': 'f.tmt_jenjang'
        }
        sort_column = valid_columns.get(sort_column, 'f.no')
        sort_direction = 'DESC' if sort_direction.lower() == 'desc' else 'ASC'

        # Base query with joins
        query = """
            SELECT f.no, jf.nama_fungsional, p.nama_pend, p.tahun, 
                   f.jenjang, f.tmt_jenjang, f.no_sk, per.nama as nama_personel, 
                   per.id_pwk as id_pwk, pwk.nama_perwakilan
            FROM tabel_fungsional f
            LEFT JOIN tabel_jenis_fungsional jf ON f.nama_fungsional = jf.no
            LEFT JOIN tabel_pendidikan p ON f.nama_pendidikan = p.no
            LEFT JOIN tabel_personel per ON f.id_personel = per.no
            LEFT JOIN ref_perwakilan pwk ON pwk.trigram = per.id_pwk
            WHERE (jf.nama_fungsional ILIKE %s OR p.nama_pend ILIKE %s OR per.nama ILIKE %s)
        """

        # Parameters for the query
        search_param = f'%{search}%'
        params = [search_param for _ in range(3)]

        # Add perwakilan filter for non-admin users
        if session.get('role') != 0:  # If not admin
            user_trigram = session.get('trigram')
            if user_trigram:
                query += " AND UPPER(TRIM(per.id_pwk)) = UPPER(TRIM(%s))"
                params.append(user_trigram.strip().upper())

        # Sort by the selected column
        query += f" ORDER BY {sort_column} {sort_direction}"

        # Execute query to get all data
        fungsional_list = execute_query(query, params, fetch=True) or []

        # Create Excel file in memory
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet("Data Fungsional")
        
        # Add formats
        title_format = workbook.add_format({
            'bold': True,
            'font_size': 14,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        subtitle_format = workbook.add_format({
            'font_size': 10,
            'align': 'left',
            'valign': 'vcenter'
        })
        
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#4472C4',
            'font_color': '#FFFFFF',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True
        })
        
        # Data formats
        data_format = workbook.add_format({
            'border': 1,
            'text_wrap': True,
            'valign': 'top'
        })
        
        center_format = workbook.add_format({
            'border': 1,
            'text_wrap': True,
            'valign': 'top',
            'align': 'center'
        })
        
        date_format = workbook.add_format({
            'border': 1,
            'text_wrap': True,
            'valign': 'top',
            'align': 'center',
            'num_format': 'dd-mm-yyyy'
        })
        
        # Write title and metadata
        worksheet.merge_range('A1:I1', 'LAPORAN DATA FUNGSIONAL', title_format)
        
        current_row = 1
        if search:
            worksheet.write(current_row, 0, f"Filter: {search}", subtitle_format)
            current_row += 1
            
        worksheet.write(current_row, 0, f"Tanggal Export: {datetime.now().strftime('%d-%m-%Y %H:%M')}", subtitle_format)
        current_row += 2
        
        # Write headers
        headers = [
            "No", 
            "Perwakilan", 
            "Personel", 
            "Fungsional", 
            "Pendidikan", 
            "Tahun", 
            "Jenjang",
            "TMT Jenjang",
            "No. SK"
        ]
        
        for col, header in enumerate(headers):
            worksheet.write(current_row, col, header, header_format)
        
        current_row += 1
        
        # Write data
        for idx, item in enumerate(fungsional_list, 1):
            tmt_jenjang = item[5] if item[5] else None
            row = [
                idx,
                item[9] or '-',  # perwakilan
                item[7] or '-',  # personel
                item[1] or '-',  # fungsional
                item[2] or '-',  # pendidikan
                item[3] or '-',  # tahun
                item[4] or '-',  # jenjang
                tmt_jenjang,     # tmt_jenjang
                item[6] or '-'   # no_sk
            ]
            
            # Apply different formats based on column
            worksheet.write(current_row, 0, row[0], center_format)  # No - centered
            worksheet.write(current_row, 1, row[1], data_format)    # Perwakilan
            worksheet.write(current_row, 2, row[2], data_format)    # Personel
            worksheet.write(current_row, 3, row[3], data_format)    # Fungsional
            worksheet.write(current_row, 4, row[4], data_format)    # Pendidikan
            worksheet.write(current_row, 5, row[5], center_format)  # Tahun - centered
            worksheet.write(current_row, 6, row[6], center_format)  # Jenjang - centered
            worksheet.write(current_row, 7, row[7], date_format if row[7] else center_format)  # TMT Jenjang - date format
            worksheet.write(current_row, 8, row[8], center_format)  # No. SK - centered
            
            current_row += 1
        
        # Auto-adjust column widths based on content
        col_widths = [5, 15, 20, 20, 20, 8, 10, 12, 10]  # Initial widths
        for i, header in enumerate(headers):
            max_length = len(header)
            for row in fungsional_list:
                if i == 0:  # No
                    val_length = len(str(row[0]))
                elif i == 1:  # Perwakilan
                    val_length = len(str(row[9] or ''))
                elif i == 2:  # Personel
                    val_length = len(str(row[7] or ''))
                elif i == 3:  # Fungsional
                    val_length = len(str(row[1] or ''))
                elif i == 4:  # Pendidikan
                    val_length = len(str(row[2] or ''))
                elif i == 5:  # Tahun
                    val_length = len(str(row[3] or ''))
                elif i == 6:  # Jenjang
                    val_length = len(str(row[4] or ''))
                elif i == 7:  # TMT Jenjang
                    val_length = 10  # Date format length
                elif i == 8:  # No. SK
                    val_length = len(str(row[6] or ''))
                
                if val_length > max_length:
                    max_length = val_length
            
            # Set column width with some padding
            worksheet.set_column(i, i, min(max_length + 2, 50))  # Max width 50 chars
        
        workbook.close()
        output.seek(0)
        
        return send_file(
            output,
            as_attachment=True,
            download_name=f"data_fungsional_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        print(f"Excel Export Error: {str(e)}")
        flash('Gagal membuat Excel: ' + str(e), 'error')
        return redirect(url_for('list_fungsional'))

@app.route('/fungsional')
def list_fungsional():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    # Get search and pagination parameters
    search = request.args.get('search', '').strip()
    page = request.args.get('page', 1, type=int)
    per_page = 20
    sort_column = request.args.get('sort', 'no')
    sort_direction = request.args.get('dir', 'asc')

    # Validate sort column
    valid_columns = {
        'no': 'f.no',
        'nama_fungsional': 'jf.nama_fungsional',
        'nama_pendidikan': 'p.nama_pend',
        'tahun_pendidikan': 'p.tahun',
        'jenjang': 'f.jenjang',
        'tmt_jenjang': 'f.tmt_jenjang'
    }
    sort_column = valid_columns.get(sort_column, 'f.no')
    
    # Validate sort direction
    sort_direction = 'DESC' if sort_direction.lower() == 'desc' else 'ASC'

    # Base query with joins
    query = f"""
        SELECT f.no, jf.nama_fungsional, p.nama_pend, p.tahun, 
               f.jenjang, f.tmt_jenjang, f.no_sk, per.nama as nama_personel, 
               per.id_pwk as id_pwk, pwk.nama_perwakilan
        FROM tabel_fungsional f
        LEFT JOIN tabel_jenis_fungsional jf ON f.nama_fungsional = jf.no
        LEFT JOIN tabel_pendidikan p ON f.nama_pendidikan = p.no
        LEFT JOIN tabel_personel per ON f.id_personel = per.no
        LEFT JOIN ref_perwakilan pwk ON pwk.trigram = id_pwk
        WHERE (jf.nama_fungsional ILIKE %s OR p.nama_pend ILIKE %s OR per.nama ILIKE %s)
    """

    # Count query for pagination
    count_query = """
        SELECT COUNT(*)
        FROM tabel_fungsional f
        LEFT JOIN tabel_jenis_fungsional jf ON f.nama_fungsional = jf.no
        LEFT JOIN tabel_pendidikan p ON f.nama_pendidikan = p.no
        LEFT JOIN tabel_personel per ON f.id_personel = per.no
        WHERE (jf.nama_fungsional ILIKE %s OR p.nama_pend ILIKE %s OR per.nama ILIKE %s)
    """

    # Add perwakilan filter for non-admin users
    if session.get('role') != 0:  # If not admin
        user_trigram = session.get('trigram')
        if user_trigram:
            query += " AND UPPER(TRIM(per.id_pwk)) = UPPER(TRIM(%s))"
            count_query += " AND UPPER(TRIM(per.id_pwk)) = UPPER(TRIM(%s))"

    # Complete the query with sorting
    query += f" ORDER BY {sort_column} {sort_direction}"
    
    search_param = f'%{search}%'
    
    # Prepare parameters for query
    query_params = [search_param, search_param, search_param]
    count_params = [search_param, search_param, search_param]
    
    # Add trigram parameter if needed
    if session.get('role') != 0 and session.get('trigram'):
        query_params.append(user_trigram.strip().upper())
        count_params.append(user_trigram.strip().upper())
    
    # Get total count
    total = execute_query(count_query, count_params, fetch_one=True)[0] or 0
    
    # Add pagination to main query
    paginated_query = query + f" LIMIT {per_page} OFFSET {(page - 1) * per_page}"
    
    # Execute query
    fungsional_list = execute_query(paginated_query, query_params, fetch=True) or []
    
    total_pages = (total + per_page - 1) // per_page
    
    return render_template('fungsional/list.html',
                         fungsional_list=fungsional_list,
                         search=search,
                         page=page,
                         per_page=per_page,
                         total=total,
                         total_pages=total_pages,
                         sort_column=sort_column,
                         sort_direction=sort_direction)

@app.route('/fungsional/create', methods=['GET', 'POST'])
def create_fungsional():
    if 'user_id' not in session:
        return redirect(url_for('login'))

    # Get dropdown data - MODIFIED to filter by perwakilan
    jenis_fungsional_query = "SELECT no, nama_fungsional FROM tabel_jenis_fungsional ORDER BY nama_fungsional"
    pendidikan_query = "SELECT no, nama_pend, tahun FROM tabel_pendidikan ORDER BY nama_pend"
    personel_query = "SELECT no, nama FROM tabel_personel"
    
    params = []
    
    # Filter personel by perwakilan for non-admin users
    if session.get('role') != 0:  # If not admin
        user_trigram = session.get('trigram')
        if user_trigram:
            personel_query += " WHERE UPPER(TRIM(id_pwk)) = UPPER(TRIM(%s))"
            params.append(user_trigram.strip().upper())

    personel_query += " ORDER BY nama"
    
    jenis_fungsional_list = execute_query(jenis_fungsional_query, fetch=True) or []
    pendidikan_list = execute_query(pendidikan_query, fetch=True) or []
    personel_list = execute_query(personel_query, params, fetch=True) or []

    if request.method == 'POST':
        try:
            id_personel = request.form.get('id_personel')
            if not id_personel:
                flash('Personel harus dipilih', 'error')
                return render_template('fungsional/create.html', 
                                    jenis_fungsional_list=jenis_fungsional_list,
                                    pendidikan_list=pendidikan_list,
                                    personel_list=personel_list,
                                    form_data=request.form)

            # Authorization check - ensure the selected personel belongs to user's perwakilan
            if session.get('role') != 0:  # If not admin
                user_trigram = session.get('trigram')
                if user_trigram:
                    personel_perwakilan = execute_query(
                        "SELECT id_pwk FROM tabel_personel WHERE no = %s",
                        (id_personel,),
                        fetch_one=True
                    )
                    if not personel_perwakilan or personel_perwakilan[0].strip().upper() != user_trigram.strip().upper():
                        flash('Anda hanya bisa menambahkan data untuk personel di perwakilan Anda', 'error')
                        return redirect(url_for('create_fungsional'))

            # Get form data
            form_data = {
                'nama_fungsional': request.form.get('nama_fungsional', '').strip(),
                'nama_pendidikan': request.form.get('nama_pendidikan', '').strip(),
                'jenjang': request.form.get('jenjang', '').strip(),
                'tmt_jenjang': request.form.get('tmt_jenjang', '').strip(),
                'no_sk': request.form.get('no_sk', '').strip(),
                'id_personel': id_personel
            }

            # Validate required fields
            if not all([form_data['nama_fungsional'], form_data['nama_pendidikan'], form_data['id_personel']]):
                flash('Field yang bertanda bintang (*) wajib diisi', 'error')
                return render_template('fungsional/create.html',
                                    form_data=form_data,
                                    jenis_fungsional_list=jenis_fungsional_list,
                                    pendidikan_list=pendidikan_list,
                                    personel_list=personel_list)

            # Validate date format if provided
            if form_data['tmt_jenjang']:
                try:
                    datetime.strptime(form_data['tmt_jenjang'], '%Y-%m-%d')
                except ValueError:
                    flash('Format tanggal TMT Jenjang tidak valid', 'error')
                    return render_template('fungsional/create.html',
                                        form_data=form_data,
                                        jenis_fungsional_list=jenis_fungsional_list,
                                        pendidikan_list=pendidikan_list,
                                        personel_list=personel_list)

            # Get next auto increment number
            next_no_result = execute_query(
                "SELECT COALESCE(MAX(no), 0) + 1 FROM tabel_fungsional",
                fetch_one=True
            )
            next_no = next_no_result[0] if next_no_result else 1

            # Insert new record
            success = execute_query(
                """INSERT INTO tabel_fungsional 
                (no, nama_fungsional, nama_pendidikan, tahun_pendidikan, 
                jenjang, tmt_jenjang, no_sk, id_personel)
                VALUES (%s, %s, %s, (SELECT tahun FROM tabel_pendidikan WHERE no = %s), 
                        %s, %s, %s, %s)""",
                (next_no, 
                form_data['nama_fungsional'],
                form_data['nama_pendidikan'],
                form_data['nama_pendidikan'],
                form_data['jenjang'],
                form_data['tmt_jenjang'] or None,
                form_data['no_sk'],
                form_data['id_personel']),
                commit=True
            )
            
            if success:
                flash('Data fungsional berhasil ditambahkan', 'success')
                return redirect(url_for('list_fungsional'))
            else:
                flash('Gagal menambahkan data fungsional', 'error')

        except Exception as e:
            logger.error(f"Error creating fungsional: {str(e)}", exc_info=True)
            flash('Terjadi kesalahan saat menyimpan data', 'error')

    return render_template('fungsional/create.html',
                         jenis_fungsional_list=jenis_fungsional_list,
                         pendidikan_list=pendidikan_list,
                         personel_list=personel_list)

@app.route('/fungsional/edit/<int:no>', methods=['GET', 'POST'])
def edit_fungsional(no):
    if 'user_id' not in session:
        return redirect(url_for('login'))

    # Authorization check - ensure the fungsional belongs to user's perwakilan
    if session.get('role') != 0:  # If not admin
        user_trigram = session.get('trigram')
        if user_trigram:
            fungsional_perwakilan = execute_query(
                """SELECT p.id_pwk 
                   FROM tabel_fungsional f
                   JOIN tabel_personel p ON f.id_personel = p.no
                   WHERE f.no = %s""",
                (no,),
                fetch_one=True
            )
            if not fungsional_perwakilan or fungsional_perwakilan[0].strip().upper() != user_trigram.strip().upper():
                flash('Anda tidak memiliki akses untuk mengedit data ini', 'error')
                return redirect(url_for('list_fungsional'))

    # Get dropdown data - MODIFIED to filter by perwakilan
    jenis_fungsional_query = "SELECT no, nama_fungsional FROM tabel_jenis_fungsional ORDER BY nama_fungsional"
    pendidikan_query = "SELECT no, nama_pend, tahun FROM tabel_pendidikan ORDER BY nama_pend"
    personel_query = "SELECT no, nama FROM tabel_personel"
    
    params = []
    
    # Filter personel by perwakilan for non-admin users
    if session.get('role') != 0:  # If not admin
        user_trigram = session.get('trigram')
        if user_trigram:
            personel_query += " WHERE UPPER(TRIM(id_pwk)) = UPPER(TRIM(%s))"
            params.append(user_trigram.strip().upper())

    personel_query += " ORDER BY nama"
    
    jenis_fungsional_list = execute_query(jenis_fungsional_query, fetch=True) or []
    pendidikan_list = execute_query(pendidikan_query, fetch=True) or []
    personel_list = execute_query(personel_query, params, fetch=True) or []

    # Get existing data
    fungsional = execute_query(
        """SELECT f.no, f.nama_fungsional, f.nama_pendidikan, f.tahun_pendidikan, 
                  f.jenjang, f.tmt_jenjang, f.no_sk, f.id_personel
           FROM tabel_fungsional f
           WHERE f.no = %s""",
        (no,),
        fetch_one=True
    )
    
    if not fungsional:
        flash('Data fungsional tidak ditemukan', 'error')
        return redirect(url_for('list_fungsional'))

    if request.method == 'POST':
        try:
            nama_fungsional = request.form.get('nama_fungsional', '').strip()
            nama_pendidikan = request.form.get('nama_pendidikan', '').strip()
            tahun_pendidikan = request.form.get('tahun_pendidikan', '').strip()
            jenjang = request.form.get('jenjang', '').strip()
            tmt_jenjang = request.form.get('tmt_jenjang', '').strip()
            no_sk = request.form.get('no_sk', '').strip()
            id_personel = request.form.get('id_personel', '').strip()

            # Additional authorization check for the new personel selection
            if session.get('role') != 0:  # If not admin
                user_trigram = session.get('trigram')
                if user_trigram:
                    personel_perwakilan = execute_query(
                        "SELECT id_pwk FROM tabel_personel WHERE no = %s",
                        (id_personel,),
                        fetch_one=True
                    )
                    if not personel_perwakilan or personel_perwakilan[0].strip().upper() != user_trigram.strip().upper():
                        flash('Anda hanya bisa memilih personel di perwakilan Anda', 'error')
                        return redirect(url_for('edit_fungsional', no=no))

            # Validate required fields
            if not all([nama_fungsional, nama_pendidikan, id_personel]):
                flash('Field yang bertanda bintang (*) wajib diisi', 'error')
                return redirect(url_for('edit_fungsional', no=no))

            success = execute_query("""
                UPDATE tabel_fungsional SET
                    nama_fungsional = %s,
                    nama_pendidikan = %s,
                    tahun_pendidikan = %s,
                    jenjang = %s,
                    tmt_jenjang = %s,
                    no_sk = %s,
                    id_personel = %s
                WHERE no = %s
            """, (nama_fungsional, nama_pendidikan, tahun_pendidikan,
                  jenjang, tmt_jenjang, no_sk, id_personel, no), commit=True)
            
            if success:
                flash('Data fungsional berhasil diperbarui', 'success')
                return redirect(url_for('list_fungsional'))
            else:
                flash('Gagal memperbarui data fungsional', 'error')
        except Exception as e:
            logger.error(f"Error updating fungsional: {str(e)}")
            flash('Terjadi kesalahan saat memperbarui data', 'error')

    return render_template('fungsional/edit.html', 
                         fungsional=fungsional,
                         jenis_fungsional_list=jenis_fungsional_list,
                         pendidikan_list=pendidikan_list,
                         personel_list=personel_list)

@app.route('/fungsional/delete/<int:no>', methods=['POST'])
def delete_fungsional(no):
    if 'user_id' not in session:
        return redirect(url_for('login'))

    # Authorization check - ensure the fungsional belongs to user's perwakilan
    if session.get('role') != 0:  # If not admin
        user_trigram = session.get('trigram')
        if user_trigram:
            fungsional_perwakilan = execute_query(
                """SELECT p.id_pwk 
                   FROM tabel_fungsional f
                   JOIN tabel_personel p ON f.id_personel = p.no
                   WHERE f.no = %s""",
                (no,),
                fetch_one=True
            )
            if not fungsional_perwakilan or fungsional_perwakilan[0].strip().upper() != user_trigram.strip().upper():
                flash('Anda tidak memiliki akses untuk menghapus data ini', 'error')
                return redirect(url_for('list_fungsional'))

    try:
        success = execute_query(
            "DELETE FROM tabel_fungsional WHERE no = %s",
            (no,),
            commit=True
        )
        
        if success:
            flash('Data fungsional berhasil dihapus', 'success')
        else:
            flash('Gagal menghapus data fungsional', 'error')
    except Exception as e:
        logger.error(f"Error deleting fungsional: {str(e)}")
        flash('Terjadi kesalahan saat menghapus data', 'error')
    
    return redirect(url_for('list_fungsional'))
# ==============================================
# AKS CRUD ROUTES
# ==============================================

@app.route('/aks/export/pdf')
def export_aks_pdf():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    try:
        # Get all parameters from the request
        search = request.args.get('search', '').strip()
        sort_column = request.args.get('sort', 'no')
        sort_direction = request.args.get('dir', 'asc')

        # Validate sort column
        valid_columns = {
            'no': 'a.no',
            'nama_personel': 'p.nama',
            'perwakilan': 'r.nama_perwakilan',
            'aks': 'a.aks',
            'tgl_penggantian': 'a.tgl_penggantian',
            'status': 'a.status',
            'no_berita': 'a.no_berita'
        }
        sort_column = valid_columns.get(sort_column, 'a.no')
        sort_direction = 'DESC' if sort_direction.lower() == 'desc' else 'ASC'

        # Base query with joins
        query = """
            SELECT a.no, a.aks, a.tgl_penggantian, a.status, a.no_berita,
                   p.no as id_personel, p.nama as nama_personel,
                   r.trigram as id_pwk, r.nama_perwakilan
            FROM tabel_aks a
            LEFT JOIN tabel_personel p ON a.id_personel = p.no
            LEFT JOIN ref_perwakilan r ON a.id_pwk = r.trigram
            WHERE (a.aks ILIKE %s OR
                  p.nama ILIKE %s OR
                  r.nama_perwakilan ILIKE %s OR
                  a.no_berita ILIKE %s OR
                  a.tgl_penggantian::text ILIKE %s)
        """

        # Parameters for the query
        search_param = f'%{search}%'
        params = [search_param for _ in range(5)]

        # Add perwakilan filter for non-admin users
        if session.get('role') != 0:  # If not admin
            user_trigram = session.get('trigram')
            if user_trigram:
                query += " AND a.id_pwk = %s"
                params.append(user_trigram)

        # Sort by the selected column
        query += f" ORDER BY {sort_column} {sort_direction}"

        # Execute query to get all data
        aks_list = execute_query(query, params, fetch=True) or []

        # Create PDF with A4 size (portrait orientation)
        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, 
                              pagesize=A4,
                              leftMargin=15,
                              rightMargin=15,
                              topMargin=20,
                              bottomMargin=20)
        elements = []
        
        # Styles configuration
        styles = getSampleStyleSheet()
        
        # Title style
        title_style = ParagraphStyle(
            'Title',
            parent=styles['Title'],
            fontSize=14,
            alignment=TA_CENTER,
            spaceAfter=12
        )
        
        # Metadata style
        meta_style = ParagraphStyle(
            'Meta',
            parent=styles['Normal'],
            fontSize=9,
            spaceAfter=6
        )
        
        # Add title
        elements.append(Paragraph("LAPORAN DATA AKS", title_style))
        
        # Add metadata
        metadata = [
            f"Tanggal Export: {datetime.now().strftime('%d-%m-%Y %H:%M')}"
        ]
        if search:
            metadata.insert(0, f"Filter: {search}")
        
        for meta in metadata:
            elements.append(Paragraph(meta, meta_style))
        
        # Prepare table data
        data = []
        headers = [
            "No", 
            "Perwakilan" if session.get('role') == 0 else None,  # Only show for admin
            "Personel", 
            "Jenis AKS", 
            "Tgl Penggantian", 
            "Status", 
            "No. Berita"
        ]
        # Filter out None values (for non-admin users)
        headers = [h for h in headers if h is not None]
        data.append(headers)
        
        for idx, item in enumerate(aks_list, 1):
            row = [
                str(idx),
                item[8] or '-' if session.get('role') == 0 else None,  # Perwakilan (only for admin)
                item[6] or '-',  # Personel
                item[1] or '-',   # AKS
                item[2].strftime('%d-%m-%Y') if item[2] else '-',  # Tgl Penggantian
                'Aktif' if item[3] else 'Non-Aktif',  # Status
                item[4] or '-'    # No. Berita
            ]
            # Filter out None values (for non-admin users)
            row = [r for r in row if r is not None]
            data.append(row)
        
        # Calculate available width using A4 width
        available_width = A4[0] - doc.leftMargin - doc.rightMargin
        
        # Define column distribution - adjusted for A4 portrait
        if session.get('role') == 0:  # Admin
            col_ratios = [0.05, 0.15, 0.2, 0.15, 0.15, 0.1, 0.2]
        else:  # Non-admin
            col_ratios = [0.05, 0.25, 0.15, 0.15, 0.1, 0.3]
        
        col_widths = [available_width * ratio for ratio in col_ratios]
        
        # Ensure minimum widths
        if session.get('role') == 0:  # Admin
            min_widths = [25, 50, 70, 50, 50, 40, 70]
        else:  # Non-admin
            min_widths = [25, 80, 50, 50, 40, 90]
            
        col_widths = [max(w, min_w) for w, min_w in zip(col_widths, min_widths)]
        
        # Create table with adjusted column widths
        table = Table(data, 
                     colWidths=col_widths,
                     repeatRows=1,
                     hAlign='LEFT')
        
        # Table style with adjusted font size and padding
        style = TableStyle([
            # Header style
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#4472C4')),
            ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,0), 7),  # Reduced font size
            ('ALIGN', (0,0), (-1,0), 'CENTER'),
            ('BOTTOMPADDING', (0,0), (-1,0), 4),  # Reduced padding
            
            # Cell style
            ('FONTSIZE', (0,1), (-1,-1), 7),  # Reduced font size
            ('LEADING', (0,0), (-1,-1), 8),   # Reduced leading
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey),
            ('WORDWRAP', (0,0), (-1,-1), True),
            
            # Alternating row colors
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.whitesmoke]),
            
            # Specific column alignments
            ('ALIGN', (0,0), (0,-1), 'CENTER'),    # No column centered
            ('ALIGN', (1,0), (1,-1), 'LEFT'),      # Perwakilan left-aligned
            ('ALIGN', (2,0), (2,-1), 'LEFT'),      # Personel left-aligned
            ('ALIGN', (3,0), (3,-1), 'LEFT'),      # Jenis AKS left-aligned
            ('ALIGN', (4,0), (4,-1), 'CENTER'),    # Tgl Penggantian centered
            ('ALIGN', (5,0), (5,-1), 'CENTER'),    # Status centered
            ('ALIGN', (6,0), (6,-1), 'LEFT')       # No. Berita left-aligned
        ])
        
        table.setStyle(style)
        elements.append(table)
        
        # Build PDF
        doc.build(elements)
        buffer.seek(0)
        
        return send_file(
            buffer,
            as_attachment=True,
            download_name=f"data_aks_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
            mimetype='application/pdf'
        )
        
    except Exception as e:
        print(f"PDF Export Error: {str(e)}")
        flash('Gagal membuat PDF: ' + str(e), 'error')
        return redirect(url_for('list_aks'))
    
@app.route('/aks/export/excel')
def export_aks_excel():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    try:
        # Get all parameters from the request
        search = request.args.get('search', '').strip()
        sort_column = request.args.get('sort', 'no')
        sort_direction = request.args.get('dir', 'asc')

        # Validate sort column
        valid_columns = {
            'no': 'a.no',
            'nama_personel': 'p.nama',
            'perwakilan': 'r.nama_perwakilan',
            'aks': 'a.aks',
            'tgl_penggantian': 'a.tgl_penggantian',
            'status': 'a.status',
            'no_berita': 'a.no_berita'
        }
        sort_column = valid_columns.get(sort_column, 'a.no')
        sort_direction = 'DESC' if sort_direction.lower() == 'desc' else 'ASC'

        # Base query with joins
        query = """
            SELECT a.no, a.aks, a.tgl_penggantian, a.status, a.no_berita,
                   p.no as id_personel, p.nama as nama_personel,
                   r.trigram as id_pwk, r.nama_perwakilan
            FROM tabel_aks a
            LEFT JOIN tabel_personel p ON a.id_personel = p.no
            LEFT JOIN ref_perwakilan r ON a.id_pwk = r.trigram
            WHERE (a.aks ILIKE %s OR
                  p.nama ILIKE %s OR
                  r.nama_perwakilan ILIKE %s OR
                  a.no_berita ILIKE %s OR
                  a.tgl_penggantian::text ILIKE %s)
        """

        # Parameters for the query
        search_param = f'%{search}%'
        params = [search_param for _ in range(5)]

        # Add perwakilan filter for non-admin users
        if session.get('role') != 0:  # If not admin
            user_trigram = session.get('trigram')
            if user_trigram:
                query += " AND a.id_pwk = %s"
                params.append(user_trigram)

        # Sort by the selected column
        query += f" ORDER BY {sort_column} {sort_direction}"

        # Execute query to get all data
        aks_list = execute_query(query, params, fetch=True) or []

        # Create Excel file in memory
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet("Data AKS")
        
        # Add formats
        title_format = workbook.add_format({
            'bold': True,
            'font_size': 14,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        subtitle_format = workbook.add_format({
            'font_size': 10,
            'align': 'left',
            'valign': 'vcenter'
        })
        
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#4472C4',
            'font_color': '#FFFFFF',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True
        })
        
        # Data formats
        data_format = workbook.add_format({
            'border': 1,
            'text_wrap': True,
            'valign': 'top'
        })
        
        status_active_format = workbook.add_format({
            'border': 1,
            'text_wrap': True,
            'valign': 'top',
            'bg_color': '#C6EFCE',  # Light green
            'font_color': '#006100'  # Dark green
        })
        
        status_inactive_format = workbook.add_format({
            'border': 1,
            'text_wrap': True,
            'valign': 'top',
            'bg_color': '#FFC7CE',  # Light red
            'font_color': '#9C0006'  # Dark red
        })
        
        # Write title and metadata
        worksheet.merge_range('A1:G1', 'LAPORAN DATA AKS', title_format)
        
        current_row = 1
        if search:
            worksheet.write(current_row, 0, f"Filter: {search}", subtitle_format)
            current_row += 1
            
        worksheet.write(current_row, 0, f"Tanggal Export: {datetime.now().strftime('%d-%m-%Y %H:%M')}", subtitle_format)
        current_row += 2
        
        # Write headers
        headers = [
            "No", 
            "Perwakilan", 
            "Personel", 
            "Jenis AKS", 
            "Tgl Penggantian", 
            "Status", 
            "No. Berita"
        ]
        
        for col, header in enumerate(headers):
            worksheet.write(current_row, col, header, header_format)
        
        current_row += 1
        
        # Write data
        for idx, item in enumerate(aks_list, 1):
            row = [
                idx,
                item[8] or '-',  # Perwakilan
                item[6] or '-',  # Personel
                item[1] or '-',  # Jenis AKS
                item[2].strftime('%d-%m-%Y') if item[2] else '-',  # Tgl Penggantian
                'Aktif' if item[3] else 'Non-Aktif',  # Status
                item[4] or '-'    # No. Berita
            ]
            
            for col, value in enumerate(row):
                if col == 5:  # Status column
                    if item[3]:  # Active
                        worksheet.write(current_row, col, value, status_active_format)
                    else:  # Inactive
                        worksheet.write(current_row, col, value, status_inactive_format)
                else:
                    worksheet.write(current_row, col, value, data_format)
            
            current_row += 1
        
        # Auto-adjust column widths based on content
        for i, header in enumerate(headers):
            max_length = len(header)
            for row in aks_list:
                if i == 0:  # No column
                    val_length = len(str(row[0]))
                elif i == 1:  # Perwakilan
                    val_length = len(str(row[8] or ''))
                elif i == 2:  # Personel
                    val_length = len(str(row[6] or ''))
                elif i == 3:  # Jenis AKS
                    val_length = len(str(row[1] or ''))
                elif i == 4:  # Tgl Penggantian
                    val_length = 10  # Fixed length for date format
                elif i == 5:  # Status
                    val_length = 6 if row[3] else 9  # "Aktif" or "Non-Aktif"
                elif i == 6:  # No. Berita
                    val_length = len(str(row[4] or ''))
                
                if val_length > max_length:
                    max_length = val_length
            
            # Set column width with some padding
            worksheet.set_column(i, i, min(max_length + 2, 50))  # Max width 50 chars
        
        # Freeze header row
        worksheet.freeze_panes(1, 0)
        
        workbook.close()
        output.seek(0)
        
        return send_file(
            output,
            as_attachment=True,
            download_name=f"data_aks_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        print(f"Excel Export Error: {str(e)}")
        flash('Gagal membuat Excel: ' + str(e), 'error')
        return redirect(url_for('list_aks'))

@app.route('/aks')
def list_aks():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    # Get search and pagination parameters
    search = request.args.get('search', '').strip()
    page = request.args.get('page', 1, type=int)
    per_page = 20
    sort_column = request.args.get('sort', 'no')
    sort_direction = request.args.get('dir', 'asc')

    # Validate sort column
    valid_columns = {
        'no': 'a.no',
        'nama_personel': 'p.nama',
        'perwakilan': 'r.nama_perwakilan',
        'aks': 'a.aks',
        'tgl_penggantian': 'a.tgl_penggantian',
        'status': 'a.status',
        'no_berita': 'a.no_berita'
    }
    sort_column = valid_columns.get(sort_column, 'a.no')
    
    # Validate sort direction
    sort_direction = 'DESC' if sort_direction.lower() == 'desc' else 'ASC'

    # Base query with joins
    query = """
        SELECT a.no, a.aks, a.tgl_penggantian, a.status, a.no_berita,
               p.no as id_personel, p.nama as nama_personel,
               r.trigram as id_pwk, r.nama_perwakilan
        FROM tabel_aks a
        LEFT JOIN tabel_personel p ON a.id_personel = p.no
        LEFT JOIN ref_perwakilan r ON a.id_pwk = r.trigram
        WHERE (a.aks ILIKE %s OR
              p.nama ILIKE %s OR
              r.nama_perwakilan ILIKE %s OR
              a.no_berita ILIKE %s OR
              a.tgl_penggantian::text ILIKE %s)
    """

    # Parameters for the query
    search_param = f'%{search}%'
    params = [search_param for _ in range(5)]

    # Add perwakilan filter for non-admin users
    if session.get('role') != 0:  # If not admin
        user_trigram = session.get('trigram')
        if user_trigram:
            query += " AND a.id_pwk = %s"
            params.append(user_trigram)

    # Add sorting
    query += f" ORDER BY {sort_column} {sort_direction}"

    # Count query for pagination
    count_query = query.replace(
        "SELECT a.no, a.aks, a.tgl_penggantian, a.status, a.no_berita, " +
        "p.no as id_personel, p.nama as nama_personel, " +
        "r.trigram as id_pwk, r.nama_perwakilan", 
        "SELECT COUNT(*)"
    ).split("ORDER BY")[0]  # Remove ORDER BY for count query

    # Get total count - PERBAIKAN DI SINI
    count_result = execute_query(count_query, params, fetch_one=True)
    total = count_result[0] if count_result else 0
    
    # Add pagination to main query
    paginated_query = query + f" LIMIT {per_page} OFFSET {(page - 1) * per_page}"
    
    # Execute query
    aks_list = execute_query(paginated_query, params, fetch=True) or []
    
    total_pages = (total + per_page - 1) // per_page if total > 0 else 1
    
    return render_template('aks/list.html',
                         aks_list=aks_list,
                         search=search,
                         page=page,
                         per_page=per_page,
                         total=total,
                         total_pages=total_pages,
                         sort_column=sort_column.replace('a.', ''),
                         sort_direction=sort_direction)

@app.route('/aks/create', methods=['GET', 'POST'])
def create_aks():
    if 'user_id' not in session:
        return redirect(url_for('login'))

    # Get personel list based on user role
    if session.get('role') == 0:  # Admin
        personel_list = execute_query(
            """SELECT p.no, p.nama, r.nama_perwakilan 
               FROM tabel_personel p
               LEFT JOIN ref_perwakilan r ON p.id_pwk = r.trigram
               ORDER BY p.nama""",
            fetch=True
        ) or []
    else:  # Regular user
        user_trigram = session.get('trigram')
        personel_list = execute_query(
            """SELECT p.no, p.nama, r.nama_perwakilan 
               FROM tabel_personel p
               LEFT JOIN ref_perwakilan r ON p.id_pwk = r.trigram
               WHERE p.id_pwk = %s
               ORDER BY p.nama""",
            (user_trigram,),
            fetch=True
        ) or []

    # Get perwakilan list based on user role
    if session.get('role') == 0:  # Admin
        perwakilan_list = execute_query(
            "SELECT trigram, nama_perwakilan FROM ref_perwakilan ORDER BY nama_perwakilan",
            fetch=True
        ) or []
    else:  # Regular user
        perwakilan_list = execute_query(
            "SELECT trigram, nama_perwakilan FROM ref_perwakilan WHERE trigram = %s",
            (session.get('trigram'),),
            fetch=True
        ) or []

    if request.method == 'POST':
        try:
            # Get form data
            form_data = request.form
            
            # Prepare data tuple with proper null handling
            data = (
                form_data.get('id_pwk', '').strip().upper() or None,
                int(form_data.get('id_personel')) if form_data.get('id_personel') else None,
                form_data.get('aks', '').strip() or None,
                form_data.get('tgl_penggantian') or None,
                bool(int(form_data.get('status', 1))),  # Default to True if not provided
                form_data.get('no_berita', '').strip() or None
            )

            # Validate required fields
            if not data[2]:  # aks
                flash('Jenis Akses wajib diisi', 'error')
                return render_template('aks/create.html', 
                                    personel_list=personel_list,
                                    perwakilan_list=perwakilan_list,
                                    form_data=request.form)

            if not data[3]:  # tgl_penggantian
                flash('Tanggal penggantian wajib diisi', 'error')
                return render_template('aks/create.html', 
                                    personel_list=personel_list,
                                    perwakilan_list=perwakilan_list,
                                    form_data=request.form)

            # For non-admin users, verify they're not trying to create for another perwakilan
            if session.get('role') != 0 and data[0] != session.get('trigram'):
                flash('Anda hanya bisa membuat data untuk perwakilan Anda sendiri', 'error')
                return render_template('aks/create.html', 
                                    personel_list=personel_list,
                                    perwakilan_list=perwakilan_list,
                                    form_data=request.form)

            # Check if personel exists and belongs to the same perwakilan
            if data[1]:  # id_personel
                personel = execute_query(
                    """SELECT p.no, p.id_pwk 
                       FROM tabel_personel p 
                       WHERE p.no = %s""",
                    (data[1],),
                    fetch_one=True
                )
                
                if not personel:
                    flash('Personel tidak valid', 'error')
                    return render_template('aks/create.html', 
                                        personel_list=personel_list,
                                        perwakilan_list=perwakilan_list,
                                        form_data=request.form)
                
                if session.get('role') != 0 and personel[1] != session.get('trigram'):
                    flash('Personel tidak valid untuk perwakilan Anda', 'error')
                    return render_template('aks/create.html', 
                                        personel_list=personel_list,
                                        perwakilan_list=perwakilan_list,
                                        form_data=request.form)

            # Execute query
            success = execute_query("""
                INSERT INTO tabel_aks 
                (id_pwk, id_personel, aks, tgl_penggantian, status, no_berita)
                VALUES (%s, %s, %s, %s, %s, %s)
            """, data, commit=True)
            
            if success:
                flash('Data Akses berhasil ditambahkan', 'success')
                return redirect(url_for('list_aks'))
            else:
                flash('Gagal menambahkan data Akses', 'error')
        except ValueError as e:
            logger.error(f"ValueError creating aks: {str(e)}")
            flash('Format data tidak valid. Pastikan semua field diisi dengan benar', 'error')
        except Exception as e:
            logger.error(f"Error creating aks: {str(e)}")
            flash('Terjadi kesalahan saat menyimpan data', 'error')

    return render_template('aks/create.html', 
                         personel_list=personel_list,
                         perwakilan_list=perwakilan_list)

@app.route('/aks/edit/<int:no>', methods=['GET', 'POST'])
def edit_aks(no):
    if 'user_id' not in session:
        return redirect(url_for('login'))

    # Get existing data with joins
    aks = execute_query(
        """SELECT a.no, a.aks, a.tgl_penggantian, a.status, a.no_berita,
                  a.id_personel, p.nama as nama_personel,
                  a.id_pwk, r.nama_perwakilan
           FROM tabel_aks a
           LEFT JOIN tabel_personel p ON a.id_personel = p.no
           LEFT JOIN ref_perwakilan r ON a.id_pwk = r.trigram
           WHERE a.no = %s""",
        (no,),
        fetch_one=True
    )
    
    if not aks:
        flash('Data Akses tidak ditemukan', 'error')
        return redirect(url_for('list_aks'))

    # Authorization check for non-admin users
    if session.get('role') != 0:  # If not admin
        user_trigram = session.get('trigram')
        if user_trigram != aks[7]:  # aks[7] is id_pwk
            flash('Anda tidak memiliki akses untuk mengedit data ini', 'error')
            return redirect(url_for('list_aks'))

    # Get personel list based on user role
    if session.get('role') == 0:  # Admin
        personel_list = execute_query(
            """SELECT p.no, p.nama, r.nama_perwakilan 
               FROM tabel_personel p
               LEFT JOIN ref_perwakilan r ON p.id_pwk = r.trigram
               ORDER BY p.nama""",
            fetch=True
        ) or []
    else:  # Regular user
        personel_list = execute_query(
            """SELECT p.no, p.nama, r.nama_perwakilan 
               FROM tabel_personel p
               LEFT JOIN ref_perwakilan r ON p.id_pwk = r.trigram
               WHERE p.id_pwk = %s
               ORDER BY p.nama""",
            (session.get('trigram'),),
            fetch=True
        ) or []

    # Get perwakilan list based on user role
    if session.get('role') == 0:  # Admin
        perwakilan_list = execute_query(
            "SELECT trigram, nama_perwakilan FROM ref_perwakilan ORDER BY nama_perwakilan",
            fetch=True
        ) or []
    else:  # Regular user
        perwakilan_list = execute_query(
            "SELECT trigram, nama_perwakilan FROM ref_perwakilan WHERE trigram = %s",
            (session.get('trigram'),),
            fetch=True
        ) or []

    if request.method == 'POST':
        try:
            # Get form data
            form_data = request.form
            
            # Prepare data tuple with proper null handling
            data = (
                form_data.get('id_pwk', '').strip().upper() or None,
                int(form_data.get('id_personel')) if form_data.get('id_personel') else None,
                form_data.get('aks', '').strip() or None,
                form_data.get('tgl_penggantian') or None,
                bool(int(form_data.get('status', 1))),  # Default to True if not provided
                form_data.get('no_berita', '').strip() or None,
                no
            )

            # Validate required fields
            if not data[2]:  # aks
                flash('Jenis Akses wajib diisi', 'error')
                return redirect(url_for('edit_aks', no=no))

            if not data[3]:  # tgl_penggantian
                flash('Tanggal penggantian wajib diisi', 'error')
                return redirect(url_for('edit_aks', no=no))

            # For non-admin users, verify they're not trying to change to another perwakilan
            if session.get('role') != 0 and data[0] != session.get('trigram'):
                flash('Anda hanya bisa mengubah data untuk perwakilan Anda sendiri', 'error')
                return redirect(url_for('edit_aks', no=no))

            # Check if personel exists and belongs to the same perwakilan
            if data[1]:  # id_personel
                personel = execute_query(
                    """SELECT p.no, p.id_pwk 
                       FROM tabel_personel p 
                       WHERE p.no = %s""",
                    (data[1],),
                    fetch_one=True
                )
                
                if not personel:
                    flash('Personel tidak valid', 'error')
                    return redirect(url_for('edit_aks', no=no))
                
                if session.get('role') != 0 and personel[1] != session.get('trigram'):
                    flash('Personel tidak valid untuk perwakilan Anda', 'error')
                    return redirect(url_for('edit_aks', no=no))

            # Execute query
            success = execute_query("""
                UPDATE tabel_aks SET
                    id_pwk = %s,
                    id_personel = %s,
                    aks = %s,
                    tgl_penggantian = %s,
                    status = %s,
                    no_berita = %s
                WHERE no = %s
            """, data, commit=True)
            
            if success:
                flash('Data Akses berhasil diperbarui', 'success')
                return redirect(url_for('list_aks'))
            else:
                flash('Gagal memperbarui data Akses', 'error')
        except ValueError as e:
            logger.error(f"ValueError updating aks: {str(e)}")
            flash('Format data tidak valid. Pastikan semua field diisi dengan benar', 'error')
        except Exception as e:
            logger.error(f"Error updating aks: {str(e)}")
            flash('Terjadi kesalahan saat memperbarui data', 'error')

    return render_template('aks/edit.html', 
                         aks=aks,
                         personel_list=personel_list,
                         perwakilan_list=perwakilan_list)

@app.route('/aks/delete/<int:no>', methods=['POST'])
def delete_aks(no):
    if 'user_id' not in session:
        return redirect(url_for('login'))

    # First check if the user has permission to delete this record
    if session.get('role') != 0:  # If not admin
        aks = execute_query(
            "SELECT id_pwk FROM tabel_aks WHERE no = %s",
            (no,),
            fetch_one=True
        )
        
        if not aks or aks[0] != session.get('trigram'):
            flash('Anda tidak memiliki akses untuk menghapus data ini', 'error')
            return redirect(url_for('list_aks'))

    try:
        success = execute_query(
            "DELETE FROM tabel_aks WHERE no = %s",
            (no,),
            commit=True
        )
        
        if success:
            flash('Data Akses berhasil dihapus', 'success')
        else:
            flash('Gagal menghapus data Akses', 'error')
    except Exception as e:
        logger.error(f"Error deleting aks: {str(e)}")
        flash('Terjadi kesalahan saat menghapus data', 'error')
    
    return redirect(url_for('list_aks'))

# ==============================================
# REF_JENIS_SISTEM CRUD ROUTES
# ==============================================

@app.route('/jenis_sistem')
def list_jenis_sistem():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    # Get search and pagination parameters
    search = request.args.get('search', '').strip()
    page = request.args.get('page', 1, type=int)
    per_page = 20
    sort_column = request.args.get('sort', 'id_jenis')
    sort_direction = request.args.get('dir', 'asc')

    # Validate sort column
    valid_columns = {
        'id_jenis': 'j.ID_JENIS',
        'jenis': 'j.JENIS',
        'kategori': 'j.KATEGORI',
        'lembar': 'j.LEMBAR',
        'format_nomor': 'j.FORMAT_NOMOR',
        'perwakilan': 'COALESCE(r.nama_perwakilan, \'ALL PERWAKILAN\')'
    }
    sort_column = valid_columns.get(sort_column, 'j.ID_JENIS')
    
    # Validate sort direction
    sort_direction = 'DESC' if sort_direction.lower() == 'desc' else 'ASC'

    # Base query with joins and COALESCE for perwakilan name
    query = """
        SELECT j.ID_JENIS, j.JENIS, j.KATEGORI, j.LEMBAR, j.FORMAT_NOMOR,
               j.TRIGRAM_PWK, 
               COALESCE(r.nama_perwakilan, 'ALL PERWAKILAN') as nama_perwakilan
        FROM ref_jenis_sistem j
        LEFT JOIN ref_perwakilan r ON j.TRIGRAM_PWK = r.trigram
        WHERE (j.ID_JENIS ILIKE %s OR
              j.JENIS ILIKE %s OR
              j.KATEGORI ILIKE %s OR
              j.FORMAT_NOMOR ILIKE %s OR
              COALESCE(r.nama_perwakilan, 'ALL PERWAKILAN') ILIKE %s)
    """

    # Parameters for the query
    search_param = f'%{search}%'
    params = [search_param for _ in range(5)]

    # Add perwakilan filter for non-admin users
    if session.get('role') != 0:  # If not admin
        user_trigram = session.get('trigram')
        if user_trigram:
            query += " AND (j.TRIGRAM_PWK = %s OR j.TRIGRAM_PWK IS NULL)"
            params.append(user_trigram)

    # Add sorting
    query += f" ORDER BY {sort_column} {sort_direction}"

    # Count query for pagination
    count_query = query.replace(
        "SELECT j.ID_JENIS, j.JENIS, j.KATEGORI, j.LEMBAR, j.FORMAT_NOMOR, " +
        "j.TRIGRAM_PWK, COALESCE(r.nama_perwakilan, 'ALL PERWAKILAN') as nama_perwakilan", 
        "SELECT COUNT(*)"
    ).split("ORDER BY")[0]  # Remove ORDER BY for count query

    # Get total count - with proper type conversion
    try:
        count_result = execute_query(count_query, params, fetch_one=True)
        total = int(count_result[0]) if count_result and count_result[0] else 0
    except (ValueError, TypeError):
        total = 0
    
    # Add pagination to main query
    paginated_query = query + f" LIMIT {per_page} OFFSET {(page - 1) * per_page}"
    
    # Execute query
    jenis_sistem_list = execute_query(paginated_query, params, fetch=True) or []
    
    total_pages = (total + per_page - 1) // per_page if total > 0 else 1
    
    return render_template('jenis_sistem/list.html',
                         jenis_sistem_list=jenis_sistem_list,
                         search=search,
                         page=page,
                         per_page=per_page,
                         total=total,
                         total_pages=total_pages,
                         sort_column=sort_column.replace('j.', '').replace('COALESCE(r.nama_perwakilan, \'ALL PERWAKILAN\')', 'perwakilan'),
                         sort_direction=sort_direction)

@app.route('/jenis_sistem/create', methods=['GET', 'POST'])
def create_jenis_sistem():
    if 'user_id' not in session:
        return redirect(url_for('login'))

    # Get perwakilan list
    if session.get('role') == 0:  # Admin
        perwakilan_list = execute_query(
            "SELECT trigram, nama_perwakilan FROM ref_perwakilan ORDER BY nama_perwakilan",
            fetch=True
        ) or []
        # Tambahkan opsi ALL untuk admin
        perwakilan_list.append(('ALL', 'ALL PERWAKILAN'))
    else:  # Regular user
        perwakilan_list = execute_query(
            "SELECT trigram, nama_perwakilan FROM ref_perwakilan WHERE trigram = %s",
            (session.get('trigram'),),
            fetch=True
        ) or []

    # Get kategori list from database
    kategori_list = execute_query(
        "SELECT kategori FROM ref_kategori_sistem ORDER BY kategori",
        fetch=True
    ) or []
    kategori_options = [k[0] for k in kategori_list]  # Extract kategori values

    if request.method == 'POST':
        try:
            form_data = request.form
            
            # Debugging - print form data
            print("Form Data:", form_data)
            
            # Handle All Perwakilan checkbox
            use_all_perwakilan = 'all_perwakilan' in form_data
            trigram_pwk = 'ALL' if use_all_perwakilan else form_data.get('trigram_pwk', '').strip().upper()
            
            # Validasi untuk non-admin
            if session.get('role') != 0:
                if use_all_perwakilan:
                    flash('Anda tidak memiliki hak akses untuk membuat data ALL PERWAKILAN', 'error')
                    return render_template('jenis_sistem/create.html', 
                                        perwakilan_list=perwakilan_list,
                                        form_data=request.form,
                                        kategori_options=kategori_options)
                elif trigram_pwk and trigram_pwk != session.get('trigram'):
                    flash('Anda hanya bisa membuat data untuk perwakilan Anda sendiri', 'error')
                    return render_template('jenis_sistem/create.html', 
                                        perwakilan_list=perwakilan_list,
                                        form_data=request.form,
                                        kategori_options=kategori_options)

            # Generate ID_JENIS
            last_id = execute_query(
                "SELECT ID_JENIS FROM ref_jenis_sistem ORDER BY ID_JENIS DESC LIMIT 1",
                fetch_one=True
            )
            new_id = f"J{(int(last_id[0][1:]) + 1):04d}" if last_id else "J0001"

            # Pastikan 'ALL' ada di tabel ref_perwakilan
            if trigram_pwk == 'ALL':
                execute_query(
                    "INSERT INTO ref_perwakilan (trigram, nama_perwakilan) VALUES ('ALL', 'ALL PERWAKILAN') ON CONFLICT (trigram) DO NOTHING",
                    commit=True
                )

            # Prepare data
            data = (
                new_id,
                form_data.get('jenis', '').strip(),
                form_data.get('kategori', '').strip(),
                int(form_data.get('lembar')) if form_data.get('lembar') else None,
                form_data.get('format_nomor', '').strip(),
                trigram_pwk,
                session.get('user_id'),
                session.get('user_id')
            )
            
            # Debugging - print prepared data
            print("Prepared Data:", data)
            logger.debug(f"Data yang akan disimpan - ID: {new_id}, Trigram: {trigram_pwk}")
            logger.debug(f"Full data: {data}")

            # Validasi required fields
            if not data[1] or not data[2]:
                flash('Jenis Sistem dan Kategori wajib diisi', 'error')
                return render_template('jenis_sistem/create.html', 
                                    perwakilan_list=perwakilan_list,
                                    form_data=request.form,
                                    kategori_options=kategori_options)

            # Execute query yang benar ke tabel ref_jenis_sistem
            success = execute_query("""
                INSERT INTO ref_jenis_sistem 
                (ID_JENIS, JENIS, KATEGORI, LEMBAR, FORMAT_NOMOR, TRIGRAM_PWK, USER_INPUT, USER_UPDATE)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
            """, data, commit=True)
            
            if success:
                flash('Data berhasil ditambahkan', 'success')
                return redirect(url_for('list_jenis_sistem'))
            else:
                flash('Gagal menambahkan data', 'error')
        except Exception as e:
            logger.error(f"Error: {str(e)}")
            flash(f'Terjadi kesalahan saat menyimpan data: {str(e)}', 'error')

    return render_template('jenis_sistem/create.html', 
                         perwakilan_list=perwakilan_list,
                         kategori_options=kategori_options)

@app.route('/jenis_sistem/edit/<id_jenis>', methods=['GET', 'POST'])
def edit_jenis_sistem(id_jenis):
    if 'user_id' not in session:
        return redirect(url_for('login'))

    # Get existing data with joins
    jenis_sistem = execute_query(
        """SELECT j.ID_JENIS, j.JENIS, j.KATEGORI, j.LEMBAR, j.FORMAT_NOMOR,
                  j.TRIGRAM_PWK, r.nama_perwakilan,
                  j.USER_INPUT, j.DATE_INPUT, j.USER_UPDATE, j.DATE_UPDATE
           FROM ref_jenis_sistem j
           LEFT JOIN ref_perwakilan r ON j.TRIGRAM_PWK = r.trigram
           WHERE j.ID_JENIS = %s""",
        (id_jenis,),
        fetch_one=True
    )
    
    if not jenis_sistem:
        flash('Data Jenis Sistem tidak ditemukan', 'error')
        return redirect(url_for('list_jenis_sistem'))

    # Authorization check for non-admin users
    if session.get('role') != 0:  # If not admin
        user_trigram = session.get('trigram')
        if user_trigram != jenis_sistem[5] and jenis_sistem[5] != 'ALL':  # jenis_sistem[5] is TRIGRAM_PWK
            flash('Anda tidak memiliki akses untuk mengedit data ini', 'error')
            return redirect(url_for('list_jenis_sistem'))

    # Get perwakilan list based on user role
    if session.get('role') == 0:  # Admin
        perwakilan_list = execute_query(
            "SELECT trigram, nama_perwakilan FROM ref_perwakilan ORDER BY nama_perwakilan",
            fetch=True
        ) or []
        # Add ALL option for admin
        perwakilan_list.append(('ALL', 'ALL PERWAKILAN'))
    else:  # Regular user
        perwakilan_list = execute_query(
            "SELECT trigram, nama_perwakilan FROM ref_perwakilan WHERE trigram = %s",
            (session.get('trigram'),),
            fetch=True
        ) or []

    # Get kategori list from database
    kategori_list = execute_query(
        "SELECT kategori FROM ref_kategori_sistem ORDER BY kategori",
        fetch=True
    ) or []
    kategori_options = [k[0] for k in kategori_list]  # Extract kategori values

    if request.method == 'POST':
        try:
            # Get form data
            form_data = request.form
            
            # Prepare data tuple with proper null handling
            data = (
                form_data.get('jenis', '').strip(),
                form_data.get('kategori', '').strip(),
                int(form_data.get('lembar')) if form_data.get('lembar') else None,
                form_data.get('format_nomor', '').strip(),
                form_data.get('trigram_pwk', '').strip().upper(),
                session.get('user_id'),  # USER_UPDATE
                id_jenis
            )

            # Validate required fields
            if not data[0]:  # JENIS
                flash('Jenis Sistem wajib diisi', 'error')
                return redirect(url_for('edit_jenis_sistem', id_jenis=id_jenis))

            if not data[1]:  # KATEGORI
                flash('Kategori wajib diisi', 'error')
                return redirect(url_for('edit_jenis_sistem', id_jenis=id_jenis))

            # For non-admin users, verify they're not trying to change to another perwakilan
            if session.get('role') != 0 and data[4] != session.get('trigram') and data[4] != 'ALL':
                flash('Anda hanya bisa mengubah data untuk perwakilan Anda sendiri', 'error')
                return redirect(url_for('edit_jenis_sistem', id_jenis=id_jenis))

            # Execute query
            success = execute_query("""
                UPDATE ref_jenis_sistem SET
                    JENIS = %s,
                    KATEGORI = %s,
                    LEMBAR = %s,
                    FORMAT_NOMOR = %s,
                    TRIGRAM_PWK = %s,
                    USER_UPDATE = %s,
                    DATE_UPDATE = CURRENT_TIMESTAMP
                WHERE ID_JENIS = %s
            """, data, commit=True)
            
            if success:
                flash('Data Jenis Sistem berhasil diperbarui', 'success')
                return redirect(url_for('list_jenis_sistem'))
            else:
                flash('Gagal memperbarui data Jenis Sistem', 'error')
        except ValueError as e:
            logger.error(f"ValueError updating jenis_sistem: {str(e)}")
            flash('Format data tidak valid. Pastikan semua field diisi dengan benar', 'error')
        except Exception as e:
            logger.error(f"Error updating jenis_sistem: {str(e)}")
            flash('Terjadi kesalahan saat memperbarui data', 'error')

    return render_template('jenis_sistem/edit.html', 
                         jenis_sistem=jenis_sistem,
                         perwakilan_list=perwakilan_list,
                         kategori_options=kategori_options)

@app.route('/jenis_sistem/delete/<id_jenis>', methods=['POST'])
def delete_jenis_sistem(id_jenis):
    if 'user_id' not in session:
        return redirect(url_for('login'))

    # First check if the user has permission to delete this record
    if session.get('role') != 0:  # If not admin
        jenis_sistem = execute_query(
            "SELECT TRIGRAM_PWK FROM ref_jenis_sistem WHERE ID_JENIS = %s",
            (id_jenis,),
            fetch_one=True
        )
        
        if not jenis_sistem or (jenis_sistem[0] != session.get('trigram') and jenis_sistem[0] != 'ALL'):
            flash('Anda tidak memiliki akses untuk menghapus data ini', 'error')
            return redirect(url_for('list_jenis_sistem'))

    try:
        success = execute_query(
            "DELETE FROM ref_jenis_sistem WHERE ID_JENIS = %s",
            (id_jenis,),
            commit=True
        )
        
        if success:
            flash('Data Jenis Sistem berhasil dihapus', 'success')
        else:
            flash('Gagal menghapus data Jenis Sistem', 'error')
    except Exception as e:
        logger.error(f"Error deleting jenis_sistem: {str(e)}")
        flash('Terjadi kesalahan saat menghapus data', 'error')
    
    return redirect(url_for('list_jenis_sistem'))

# ==============================================
# TABEL_SISTEM CRUD ROUTES 
# ==============================================
@app.route('/sistem/export/pdf')
def export_sistem_pdf():
    if 'user_id' not in session or session.get('role') != 0:  # Admin only
        return redirect(url_for('login'))
    
    try:
        # Get all parameters from the request
        search = request.args.get('search', '').strip()
        sort_column = request.args.get('sort', 'id_sistem')
        sort_direction = request.args.get('dir', 'asc')

        # Validate sort column
        valid_columns = {
            'id_sistem': 's.id_sistem',
            'tahun': 's.tahun',
            'jenis': 'j.jenis',
            'no_sistem': 's.no_sistem',
            'nama_sistem': 's.nama_sistem',
            'jml_lembar': 's.jml_lembar',
            'status': 's.status',
            'no_urut': 's.no_urut'
        }
        sort_column = valid_columns.get(sort_column, 's.id_sistem')
        sort_direction = 'DESC' if sort_direction.lower() == 'desc' else 'ASC'

        # Base query with joins
        query = """
            SELECT s.id_sistem, s.tahun, j.jenis, s.no_sistem, s.nama_sistem,
                   s.jml_lembar, s.status, s.no_urut
            FROM tabel_sistem s
            JOIN ref_jenis_sistem j ON s.id_jenis = j.id_jenis
            WHERE (s.id_sistem ILIKE %s OR
                  s.tahun::text ILIKE %s OR
                  j.jenis ILIKE %s OR
                  s.no_sistem ILIKE %s OR
                  s.nama_sistem ILIKE %s)
        """

        # Parameters for the query
        search_param = f'%{search}%'
        params = [search_param for _ in range(5)]

        # Sort by the selected column
        query += f" ORDER BY {sort_column} {sort_direction}"

        # Execute query to get all data
        sistem_list = execute_query(query, params, fetch=True) or []

        # Create PDF with A4 orientation (portrait)
        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, 
                              pagesize=A4,  # Changed from landscape(letter) to A4
                              leftMargin=15,
                              rightMargin=15,
                              topMargin=20,
                              bottomMargin=20)
        elements = []
        
        # Styles configuration
        styles = getSampleStyleSheet()
        
        # Title style
        title_style = ParagraphStyle(
            'Title',
            parent=styles['Title'],
            fontSize=14,
            alignment=TA_CENTER,
            spaceAfter=12
        )
        
        # Metadata style
        meta_style = ParagraphStyle(
            'Meta',
            parent=styles['Normal'],
            fontSize=9,
            spaceAfter=6
        )
        
        # Add title
        elements.append(Paragraph("LAPORAN DATA SISTEM", title_style))
        
        # Add metadata
        metadata = [
            f"Tanggal Export: {datetime.now().strftime('%d-%m-%Y %H:%M')}"
        ]
        if search:
            metadata.insert(0, f"Filter: {search}")
        
        for meta in metadata:
            elements.append(Paragraph(meta, meta_style))
        
        # Prepare table data
        data = []
        headers = [
            "No", 
            "ID Sistem", 
            "Tahun", 
            "Jenis Sistem", 
            "Nomor Sistem", 
            "Nama Sistem",
            "Jml Lembar",
            "Status",
            "No Urut"
        ]
        data.append(headers)
        
        # Status mapping for display
        status_map = {
            0: "Belum Berlaku",
            1: "Sedang Berlaku",
            2: "Tidak Berlaku"
        }
        
        for idx, item in enumerate(sistem_list, 1):
            row = [
                str(idx),
                item[0] or '-',
                str(item[1]) if item[1] is not None else '-',
                item[2] or '-',
                item[3] or '-',
                item[4] or '-',
                str(item[5]) if item[5] is not None else '-',
                status_map.get(item[6], str(item[6])) if item[6] is not None else '-',
                str(item[7]) if item[7] is not None else '-'
            ]
            data.append(row)
        
        # Calculate available width (subtract margins)
        available_width = A4[0] - doc.leftMargin - doc.rightMargin
        
        # Define column distribution (adjust these ratios as needed)
        col_ratios = [0.05, 0.1, 0.07, 0.15, 0.15, 0.2, 0.08, 0.1, 0.1]  # Sum should be 1.0
        col_widths = [available_width * ratio for ratio in col_ratios]
        
        # Ensure minimum widths
        min_widths = [30, 50, 40, 60, 70, 80, 40, 50, 40]  # Minimum widths in points
        col_widths = [max(w, min_w) for w, min_w in zip(col_widths, min_widths)]
        
        # Create table
        table = Table(data, 
                     colWidths=col_widths,
                     repeatRows=1,
                     hAlign='LEFT')
        
        # Table style
        style = TableStyle([
            # Header style
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#4472C4')),
            ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,0), 8),
            ('ALIGN', (0,0), (-1,0), 'CENTER'),
            ('BOTTOMPADDING', (0,0), (-1,0), 6),
            
            # Cell style
            ('FONTSIZE', (0,1), (-1,-1), 8),
            ('LEADING', (0,0), (-1,-1), 9),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey),
            ('WORDWRAP', (0,0), (-1,-1), True),
            
            # Alternating row colors
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.whitesmoke]),
            
            # Specific column alignments
            ('ALIGN', (0,0), (0,-1), 'CENTER'),    # No column centered
            ('ALIGN', (1,0), (1,-1), 'CENTER'),    # ID Sistem centered
            ('ALIGN', (2,0), (2,-1), 'CENTER'),    # Tahun centered
            ('ALIGN', (3,0), (3,-1), 'LEFT'),      # Jenis Sistem left-aligned
            ('ALIGN', (4,0), (4,-1), 'LEFT'),      # Nomor Sistem left-aligned
            ('ALIGN', (5,0), (5,-1), 'LEFT'),      # Nama Sistem left-aligned
            ('ALIGN', (6,0), (6,-1), 'CENTER'),    # Jml Lembar centered
            ('ALIGN', (7,0), (7,-1), 'CENTER'),    # Status centered
            ('ALIGN', (8,0), (8,-1), 'CENTER')     # No Urut centered
        ])
        
        table.setStyle(style)
        elements.append(table)
        
        # Build PDF
        doc.build(elements)
        buffer.seek(0)
        
        return send_file(
            buffer,
            as_attachment=True,
            download_name=f"data_sistem_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
            mimetype='application/pdf'
        )
        
    except Exception as e:
        print(f"PDF Export Error: {str(e)}")
        flash('Gagal membuat PDF: ' + str(e), 'error')
        return redirect(url_for('list_sistem'))

@app.route('/sistem/export/excel')
def export_sistem_excel():
    if 'user_id' not in session or session.get('role') != 0:  # Admin only
        return redirect(url_for('login'))
    
    try:
        # Get all parameters from the request
        search = request.args.get('search', '').strip()
        sort_column = request.args.get('sort', 'id_sistem')
        sort_direction = request.args.get('dir', 'asc')

        # Validate sort column
        valid_columns = {
            'id_sistem': 's.id_sistem',
            'tahun': 's.tahun',
            'jenis': 'j.jenis',
            'no_sistem': 's.no_sistem',
            'nama_sistem': 's.nama_sistem',
            'jml_lembar': 's.jml_lembar',
            'status': 's.status',
            'no_urut': 's.no_urut'
        }
        sort_column = valid_columns.get(sort_column, 's.id_sistem')
        sort_direction = 'DESC' if sort_direction.lower() == 'desc' else 'ASC'

        # Base query with joins
        query = """
            SELECT s.id_sistem, s.tahun, j.jenis, s.no_sistem, s.nama_sistem,
                   s.jml_lembar, s.status, s.no_urut
            FROM tabel_sistem s
            JOIN ref_jenis_sistem j ON s.id_jenis = j.id_jenis
            WHERE (s.id_sistem ILIKE %s OR
                  s.tahun::text ILIKE %s OR
                  j.jenis ILIKE %s OR
                  s.no_sistem ILIKE %s OR
                  s.nama_sistem ILIKE %s)
        """

        # Parameters for the query
        search_param = f'%{search}%'
        params = [search_param for _ in range(5)]

        # Sort by the selected column
        query += f" ORDER BY {sort_column} {sort_direction}"

        # Execute query to get all data
        sistem_list = execute_query(query, params, fetch=True) or []

        # Create Excel file in memory
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet("Data Sistem")
        
        # Add formats
        title_format = workbook.add_format({
            'bold': True,
            'font_size': 14,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        subtitle_format = workbook.add_format({
            'font_size': 10,
            'align': 'left',
            'valign': 'vcenter'
        })
        
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#4472C4',
            'font_color': '#FFFFFF',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True
        })
        
        # Data formats
        data_format = workbook.add_format({
            'border': 1,
            'text_wrap': True,
            'valign': 'top'
        })
        
        center_format = workbook.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        # Status format colors
        status_format0 = workbook.add_format({
            'border': 1,
            'bg_color': '#CCCCCC',  # Gray for Belum Berlaku
            'align': 'center'
        })
        
        status_format1 = workbook.add_format({
            'border': 1,
            'bg_color': '#C6EFCE',  # Green for Sedang Berlaku
            'align': 'center'
        })
        
        status_format2 = workbook.add_format({
            'border': 1,
            'bg_color': '#FFC7CE',  # Red for Tidak Berlaku
            'align': 'center'
        })
        
        # Write title and metadata
        worksheet.merge_range('A1:I1', 'LAPORAN DATA SISTEM', title_format)
        
        current_row = 1
        if search:
            worksheet.write(current_row, 0, f"Filter: {search}", subtitle_format)
            current_row += 1
            
        worksheet.write(current_row, 0, f"Tanggal Export: {datetime.now().strftime('%d-%m-%Y %H:%M')}", subtitle_format)
        current_row += 2
        
        # Write headers
        headers = [
            "No", 
            "ID Sistem", 
            "Tahun", 
            "Jenis Sistem", 
            "Nomor Sistem", 
            "Nama Sistem",
            "Jml Lembar",
            "Status",
            "No Urut"
        ]
        
        for col, header in enumerate(headers):
            worksheet.write(current_row, col, header, header_format)
        
        current_row += 1
        
        # Status mapping for display
        status_map = {
            0: "Belum Berlaku",
            1: "Sedang Berlaku",
            2: "Tidak Berlaku"
        }
        
        # Write data
        for idx, item in enumerate(sistem_list, 1):
            # Write each cell with appropriate format
            worksheet.write(current_row, 0, idx, center_format)  # No
            
            # ID Sistem
            worksheet.write(current_row, 1, item[0] or '-', center_format)
            
            # Tahun
            worksheet.write(current_row, 2, str(item[1]) if item[1] is not None else '-', center_format)
            
            # Jenis Sistem
            worksheet.write(current_row, 3, item[2] or '-', data_format)
            
            # Nomor Sistem
            worksheet.write(current_row, 4, item[3] or '-', data_format)
            
            # Nama Sistem
            worksheet.write(current_row, 5, item[4] or '-', data_format)
            
            # Jml Lembar
            worksheet.write(current_row, 6, str(item[5]) if item[5] is not None else '-', center_format)
            
            # Status with conditional formatting
            status_value = item[6]
            status_text = status_map.get(status_value, str(status_value)) if status_value is not None else '-'
            
            if status_value == 0:
                worksheet.write(current_row, 7, status_text, status_format0)
            elif status_value == 1:
                worksheet.write(current_row, 7, status_text, status_format1)
            elif status_value == 2:
                worksheet.write(current_row, 7, status_text, status_format2)
            else:
                worksheet.write(current_row, 7, status_text, center_format)
            
            # No Urut
            worksheet.write(current_row, 8, str(item[7]) if item[7] is not None else '-', center_format)
            
            current_row += 1
        
        # Auto-adjust column widths based on content
        col_widths = [
            ("No", 5),
            ("ID Sistem", 10),
            ("Tahun", 8),
            ("Jenis Sistem", 20),
            ("Nomor Sistem", 15),
            ("Nama Sistem", 30),
            ("Jml Lembar", 10),
            ("Status", 15),
            ("No Urut", 10)
        ]
        
        for i, (header, default_width) in enumerate(col_widths):
            max_length = len(header)
            for row in sistem_list:
                if i == 0:  # No column
                    val_length = len(str(idx))
                elif i == 1:  # ID Sistem
                    val_length = len(str(row[0] or ''))
                elif i == 2:  # Tahun
                    val_length = len(str(row[1] or ''))
                elif i == 3:  # Jenis Sistem
                    val_length = len(str(row[2] or ''))
                elif i == 4:  # Nomor Sistem
                    val_length = len(str(row[3] or ''))
                elif i == 5:  # Nama Sistem
                    val_length = len(str(row[4] or ''))
                elif i == 6:  # Jml Lembar
                    val_length = len(str(row[5] or ''))
                elif i == 7:  # Status
                    status_text = status_map.get(row[6], str(row[6])) if row[6] is not None else '-'
                    val_length = len(status_text)
                elif i == 8:  # No Urut
                    val_length = len(str(row[7] or ''))
                
                if val_length > max_length:
                    max_length = val_length
            
            # Set column width with some padding
            worksheet.set_column(i, i, min(max_length + 2, 50))  # Max width 50 chars
        
        workbook.close()
        output.seek(0)
        
        return send_file(
            output,
            as_attachment=True,
            download_name=f"data_sistem_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        print(f"Excel Export Error: {str(e)}")
        flash('Gagal membuat Excel: ' + str(e), 'error')
        return redirect(url_for('list_sistem'))
        
@app.route('/sistem')
def list_sistem():
    if 'user_id' not in session or session.get('role') != 0:  # Admin only
        return redirect(url_for('login'))
    
    # Get search and pagination parameters
    search = request.args.get('search', '').strip()
    page = request.args.get('page', 1, type=int)
    per_page = 20
    sort_column = request.args.get('sort', 'id_sistem')
    sort_direction = request.args.get('dir', 'asc')

    # Validate sort column
    valid_columns = {
        'id_sistem': 's.id_sistem',
        'tahun': 's.tahun',
        'jenis': 'j.jenis',
        'no_sistem': 's.no_sistem',
        'nama_sistem': 's.nama_sistem',
        'jml_lembar': 's.jml_lembar',
        'status': 's.status',
        'no_urut': 's.no_urut'
    }
    sort_column = valid_columns.get(sort_column, 's.id_sistem')
    
    # Validate sort direction
    sort_direction = 'DESC' if sort_direction.lower() == 'desc' else 'ASC'

    # Base query with joins
    query = """
        SELECT s.id_sistem, s.tahun, j.jenis, s.no_sistem, s.nama_sistem,
               s.jml_lembar, s.status, s.no_urut,
               s.user_input, s.date_input, s.user_update, s.date_update
        FROM tabel_sistem s
        JOIN ref_jenis_sistem j ON s.id_jenis = j.id_jenis
        WHERE (s.id_sistem ILIKE %s OR
              s.tahun::text ILIKE %s OR
              j.jenis ILIKE %s OR
              s.no_sistem ILIKE %s OR
              s.nama_sistem ILIKE %s)
    """

    # Parameters for the query
    search_param = f'%{search}%'
    params = [search_param for _ in range(5)]

    # Add sorting
    query += f" ORDER BY {sort_column} {sort_direction}"

    # Count query for pagination
    count_query = query.replace(
        "SELECT s.id_sistem, s.tahun, j.jenis, s.no_sistem, s.nama_sistem, " +
        "s.jml_lembar, s.status, s.no_urut, s.user_input, s.date_input, " +
        "s.user_update, s.date_update", 
        "SELECT COUNT(*)"
    ).split("ORDER BY")[0]  # Remove ORDER BY for count query

    # Get total count
    try:
        count_result = execute_query(count_query, params, fetch_one=True)
        total = int(count_result[0]) if count_result and count_result[0] else 0
    except (ValueError, TypeError):
        total = 0
    
    # Add pagination to main query
    paginated_query = query + f" LIMIT {per_page} OFFSET {(page - 1) * per_page}"
    
    # Execute query
    sistem_list = execute_query(paginated_query, params, fetch=True) or []
    
    total_pages = (total + per_page - 1) // per_page if total > 0 else 1
    
    return render_template('sistem/list.html',
                         sistem_list=sistem_list,
                         search=search,
                         page=page,
                         per_page=per_page,
                         total=total,
                         total_pages=total_pages,
                         sort_column=sort_column.replace('s.', ''),
                         sort_direction=sort_direction)

@app.route('/sistem/create', methods=['GET', 'POST'])
def create_sistem():
    if 'user_id' not in session or session.get('role') != 0:  # Admin only
        return redirect(url_for('login'))

    # Get jenis sistem list
    jenis_list = execute_query(
        "SELECT id_jenis, jenis FROM ref_jenis_sistem ORDER BY jenis",
        fetch=True
    ) or []

    if request.method == 'POST':
        try:
            form_data = request.form
            
            # Generate ID_SISTEM
            last_id = execute_query(
                "SELECT id_sistem FROM tabel_sistem ORDER BY id_sistem DESC LIMIT 1",
                fetch_one=True
            )
            new_id = f"S{(int(last_id[0][1:]) + 1):04d}" if last_id else "S0001"

            # Prepare data
            data = (
                new_id,
                int(form_data.get('tahun')) if form_data.get('tahun') else None,
                form_data.get('id_jenis', '').strip(),
                form_data.get('no_sistem', '').strip(),
                form_data.get('nama_sistem', '').strip(),
                int(form_data.get('jml_lembar')) if form_data.get('jml_lembar') else None,
                int(form_data.get('status')) if form_data.get('status') else 0,
                int(form_data.get('no_urut')) if form_data.get('no_urut') else None,
                session.get('user_id'),
                session.get('user_id')
            )

            # Validasi required fields
            if not data[1] or not data[2] or not data[3]:
                flash('Tahun, Jenis Sistem, dan Nomor Sistem wajib diisi', 'error')
                return render_template('sistem/create.html', 
                                    jenis_list=jenis_list,
                                    form_data=request.form,
                                    status_options=[(0, 'Belum Berlaku'), (1, 'Sedang Berlaku'), (2, 'Tidak Berlaku')])

            # Execute query
            success = execute_query("""
                INSERT INTO tabel_sistem 
                (id_sistem, tahun, id_jenis, no_sistem, nama_sistem, jml_lembar, 
                 status, no_urut, user_input, user_update)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """, data, commit=True)
            
            if success:
                flash('Data berhasil ditambahkan', 'success')
                return redirect(url_for('list_sistem'))
            else:
                flash('Gagal menambahkan data', 'error')
        except Exception as e:
            logger.error(f"Error: {str(e)}")
            flash('Terjadi kesalahan saat menyimpan data', 'error')

    return render_template('sistem/create.html', 
                         jenis_list=jenis_list,
                         status_options=[(0, 'Belum Berlaku'), (1, 'Sedang Berlaku'), (2, 'Tidak Berlaku')])

@app.route('/sistem/edit/<id_sistem>', methods=['GET', 'POST'])
def edit_sistem(id_sistem):
    if 'user_id' not in session or session.get('role') != 0:  # Admin only
        return redirect(url_for('login'))

    # Get existing data with joins
    sistem = execute_query(
        """SELECT s.id_sistem, s.tahun, s.id_jenis, j.jenis, s.no_sistem, 
                  s.nama_sistem, s.jml_lembar, s.status, s.no_urut,
                  s.user_input, s.date_input, s.user_update, s.date_update
           FROM tabel_sistem s
           JOIN ref_jenis_sistem j ON s.id_jenis = j.id_jenis
           WHERE s.id_sistem = %s""",
        (id_sistem,),
        fetch_one=True
    )
    
    if not sistem:
        flash('Data Sistem tidak ditemukan', 'error')
        return redirect(url_for('list_sistem'))

    # Get jenis sistem list
    jenis_list = execute_query(
        "SELECT id_jenis, jenis FROM ref_jenis_sistem ORDER BY jenis",
        fetch=True
    ) or []

    if request.method == 'POST':
        try:
            # Get form data
            form_data = request.form
            
            # Prepare data tuple with proper null handling
            data = (
                int(form_data.get('tahun')) if form_data.get('tahun') else None,
                form_data.get('id_jenis', '').strip(),
                form_data.get('no_sistem', '').strip(),
                form_data.get('nama_sistem', '').strip(),
                int(form_data.get('jml_lembar')) if form_data.get('jml_lembar') else None,
                int(form_data.get('status')) if form_data.get('status') else 0,
                int(form_data.get('no_urut')) if form_data.get('no_urut') else None,
                session.get('user_id'),  # USER_UPDATE
                id_sistem
            )

            # Validate required fields
            if not data[0]:  # TAHUN
                flash('Tahun wajib diisi', 'error')
                return redirect(url_for('edit_sistem', id_sistem=id_sistem))

            if not data[1]:  # ID_JENIS
                flash('Jenis Sistem wajib diisi', 'error')
                return redirect(url_for('edit_sistem', id_sistem=id_sistem))

            if not data[2]:  # NO_SISTEM
                flash('Nomor Sistem wajib diisi', 'error')
                return redirect(url_for('edit_sistem', id_sistem=id_sistem))

            # Execute query
            success = execute_query("""
                UPDATE tabel_sistem SET
                    tahun = %s,
                    id_jenis = %s,
                    no_sistem = %s,
                    nama_sistem = %s,
                    jml_lembar = %s,
                    status = %s,
                    no_urut = %s,
                    user_update = %s,
                    date_update = CURRENT_TIMESTAMP
                WHERE id_sistem = %s
            """, data, commit=True)
            
            if success:
                flash('Data Sistem berhasil diperbarui', 'success')
                return redirect(url_for('list_sistem'))
            else:
                flash('Gagal memperbarui data Sistem', 'error')
        except ValueError as e:
            logger.error(f"ValueError updating sistem: {str(e)}")
            flash('Format data tidak valid. Pastikan semua field diisi dengan benar', 'error')
        except Exception as e:
            logger.error(f"Error updating sistem: {str(e)}")
            flash('Terjadi kesalahan saat memperbarui data', 'error')

    return render_template('sistem/edit.html', 
                         sistem=sistem,
                         jenis_list=jenis_list,
                         status_options=[(0, 'Belum Berlaku'), (1, 'Sedang Berlaku'), (2, 'Tidak Berlaku')])

@app.route('/sistem/delete/<id_sistem>', methods=['POST'])
def delete_sistem(id_sistem):
    if 'user_id' not in session or session.get('role') != 0:  # Admin only
        return redirect(url_for('login'))

    try:
        success = execute_query(
            "DELETE FROM tabel_sistem WHERE id_sistem = %s",
            (id_sistem,),
            commit=True
        )
        
        if success:
            flash('Data Sistem berhasil dihapus', 'success')
        else:
            flash('Gagal menghapus data Sistem', 'error')
    except Exception as e:
        logger.error(f"Error deleting sistem: {str(e)}")
        flash('Terjadi kesalahan saat menghapus data', 'error')
    
    return redirect(url_for('list_sistem'))

# ==============================================
# ALKOM CRUD ROUTES
# ==============================================

@app.route('/alkom/export/pdf')
def export_alkom_pdf():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    try:
        # Get all parameters from the request
        search = request.args.get('search', '').strip()
        sort_column = request.args.get('sort', 'no')
        sort_direction = request.args.get('dir', 'asc')

        # Validate sort column
        valid_columns = {
            'no': 'a.no',
            'perwakilan': 'r.nama_perwakilan',
            'provider': 'a.provider',
            'jenis_telpon_satelit': 'a.jenis_telpon_satelit',
            'status_alkom': 'a.status_alkom',
            'tahun_pengadaan': 'a.tahun_pengadaan'
        }
        sort_column = valid_columns.get(sort_column, 'a.no')
        sort_direction = 'DESC' if sort_direction.lower() == 'desc' else 'ASC'

        # Base query with joins
        query = """
            SELECT a.no, a.provider, a.jenis_telpon_satelit, a.nomor_telp_satelit,
                   a.status_alkom, a.fasilitas_internet, a.status_langganan,
                   a.pengadaan, a.tahun_pengadaan, a.pencatatan_BMN, a.tahun_pencatatan,
                   a.no_bmn, r.trigram as id_pwk, r.nama_perwakilan
            FROM tabel_Alkom a
            LEFT JOIN ref_perwakilan r ON a.perwakilan = r.trigram
            WHERE (a.provider ILIKE %s OR
                  a.jenis_telpon_satelit ILIKE %s OR
                  a.nomor_telp_satelit ILIKE %s OR
                  r.nama_perwakilan ILIKE %s OR
                  a.pengadaan ILIKE %s OR
                  a.no_bmn ILIKE %s)
        """

        # Parameters for the query
        search_param = f'%{search}%'
        params = [search_param for _ in range(6)]

        # Add perwakilan filter for non-admin users
        if session.get('role') != 0:  # If not admin
            user_trigram = session.get('trigram')
            if user_trigram:
                query += " AND a.perwakilan = %s"
                params.append(user_trigram)

        # Sort by the selected column
        query += f" ORDER BY {sort_column} {sort_direction}"

        # Execute query to get all data
        alkom_list = execute_query(query, params, fetch=True) or []

        # Create PDF with A4 size (portrait orientation)
        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, 
                              pagesize=A4,  # Changed to A4
                              leftMargin=15,
                              rightMargin=15,
                              topMargin=20,
                              bottomMargin=20)
        elements = []
        
        # Styles configuration
        styles = getSampleStyleSheet()
        
        # Title style
        title_style = ParagraphStyle(
            'Title',
            parent=styles['Title'],
            fontSize=14,
            alignment=TA_CENTER,
            spaceAfter=12
        )
        
        # Metadata style
        meta_style = ParagraphStyle(
            'Meta',
            parent=styles['Normal'],
            fontSize=9,
            spaceAfter=6
        )
        
        # Add title
        elements.append(Paragraph("LAPORAN DATA ALKOM", title_style))
        
        # Add metadata
        metadata = [
            f"Tanggal Export: {datetime.now().strftime('%d-%m-%Y %H:%M')}"
        ]
        if search:
            metadata.insert(0, f"Filter: {search}")
        
        for meta in metadata:
            elements.append(Paragraph(meta, meta_style))
        
        # Prepare table data
        data = []
        headers = [
            "No",
            "Perwakilan" if session.get('role') == 0 else None,
            "Provider",
            "Jenis Telpon Satelit",
            "Nomor Telp",
            "Status",
            "Fasilitas Internet",
            "Pengadaan",
            "Pencatatan BMN",
            "No. BMN"
        ]
        # Filter out None values (for non-admin users)
        headers = [h for h in headers if h is not None]
        data.append(headers)
        
        for idx, item in enumerate(alkom_list, 1):
            row = [
                str(idx),
                item[13] if session.get('role') == 0 else None,  # Perwakilan
                item[1] or '-',   # Provider
                item[2] or '-',   # Jenis Telpon Satelit
                item[3] or '-',   # Nomor Telp
                item[4] or '-',   # Status
                item[5] or '-',   # Fasilitas Internet
                f"{item[7] or '-'}\n{item[8] or ''}",  # Pengadaan + Tahun
                f"{item[9] or '-'}\n{item[10] or ''}",  # Pencatatan BMN + Tahun
                item[11] or '-'   # No BMN
            ]
            # Filter out None values (for non-admin users)
            row = [r for r in row if r is not None]
            data.append(row)
        
        # Calculate available width using A4 width
        available_width = A4[0] - doc.leftMargin - doc.rightMargin
        
        # Define column distribution - adjusted ratios for better fit
        if session.get('role') == 0:  # Admin
            col_ratios = [0.05, 0.12, 0.12, 0.12, 0.1, 0.08, 0.1, 0.12, 0.12, 0.07]
        else:  # Non-admin
            col_ratios = [0.05, 0.15, 0.15, 0.1, 0.08, 0.12, 0.15, 0.12, 0.08]
        
        col_widths = [available_width * ratio for ratio in col_ratios]
        
        # Ensure minimum widths - reduced to prevent overflow
        if session.get('role') == 0:  # Admin
            min_widths = [25, 40, 40, 40, 40, 30, 40, 50, 50, 30]
        else:  # Non-admin
            min_widths = [25, 40, 40, 40, 30, 40, 50, 50, 30]
            
        col_widths = [max(w, min_w) for w, min_w in zip(col_widths, min_widths)]
        
        # Create table with adjusted column widths
        table = Table(data, 
                     colWidths=col_widths,
                     repeatRows=1,
                     hAlign='LEFT')
        
        # Table style with adjusted font size and padding
        style = TableStyle([
            # Header style
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#4472C4')),
            ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,0), 7),  # Reduced font size
            ('ALIGN', (0,0), (-1,0), 'CENTER'),
            ('BOTTOMPADDING', (0,0), (-1,0), 4),  # Reduced padding
            
            # Cell style
            ('FONTSIZE', (0,1), (-1,-1), 7),  # Reduced font size
            ('LEADING', (0,0), (-1,-1), 8),   # Reduced leading
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey),
            ('WORDWRAP', (0,0), (-1,-1), True),
            
            # Alternating row colors
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.whitesmoke]),
            
            # Specific column alignments
            ('ALIGN', (0,0), (0,-1), 'CENTER'),
            ('ALIGN', (1,0), (1,-1), 'LEFT'),
            ('ALIGN', (-4,0), (-4,-1), 'CENTER'),
            ('ALIGN', (-3,0), (-3,-1), 'CENTER'),
            ('ALIGN', (-1,0), (-1,-1), 'LEFT'),
            
            # Cell padding
            ('TOPPADDING', (0,0), (-1,-1), 2),
            ('BOTTOMPADDING', (0,0), (-1,-1), 2),
            ('LEFTPADDING', (0,0), (-1,-1), 3),
            ('RIGHTPADDING', (0,0), (-1,-1), 3)
        ])
        
        table.setStyle(style)
        elements.append(table)
        
        # Build PDF
        doc.build(elements)
        buffer.seek(0)
        
        return send_file(
            buffer,
            as_attachment=True,
            download_name=f"data_alkom_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
            mimetype='application/pdf'
        )
        
    except Exception as e:
        print(f"PDF Export Error: {str(e)}")
        flash('Gagal membuat PDF: ' + str(e), 'error')
        return redirect(url_for('list_alkom'))

@app.route('/alkom/export/excel')
def export_alkom_excel():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    try:
        # Get all parameters from the request
        search = request.args.get('search', '').strip()
        sort_column = request.args.get('sort', 'no')
        sort_direction = request.args.get('dir', 'asc')

        # Validate sort column
        valid_columns = {
            'no': 'a.no',
            'perwakilan': 'r.nama_perwakilan',
            'provider': 'a.provider',
            'jenis_telpon_satelit': 'a.jenis_telpon_satelit',
            'status_alkom': 'a.status_alkom',
            'tahun_pengadaan': 'a.tahun_pengadaan'
        }
        sort_column = valid_columns.get(sort_column, 'a.no')
        sort_direction = 'DESC' if sort_direction.lower() == 'desc' else 'ASC'

        # Base query with joins
        query = """
            SELECT a.no, a.provider, a.jenis_telpon_satelit, a.nomor_telp_satelit,
                   a.status_alkom, a.fasilitas_internet, a.status_langganan,
                   a.pengadaan, a.tahun_pengadaan, a.pencatatan_BMN, a.tahun_pencatatan,
                   a.no_bmn, r.trigram as id_pwk, r.nama_perwakilan
            FROM tabel_Alkom a
            LEFT JOIN ref_perwakilan r ON a.perwakilan = r.trigram
            WHERE (a.provider ILIKE %s OR
                  a.jenis_telpon_satelit ILIKE %s OR
                  a.nomor_telp_satelit ILIKE %s OR
                  r.nama_perwakilan ILIKE %s OR
                  a.pengadaan ILIKE %s OR
                  a.no_bmn ILIKE %s)
        """

        # Parameters for the query
        search_param = f'%{search}%'
        params = [search_param for _ in range(6)]

        # Add perwakilan filter for non-admin users
        if session.get('role') != 0:  # If not admin
            user_trigram = session.get('trigram')
            if user_trigram:
                query += " AND a.perwakilan = %s"
                params.append(user_trigram)

        # Sort by the selected column
        query += f" ORDER BY {sort_column} {sort_direction}"

        # Execute query to get all data
        alkom_list = execute_query(query, params, fetch=True) or []

        # Create Excel file in memory
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet("Data Alkom")
        
        # Add formats
        title_format = workbook.add_format({
            'bold': True,
            'font_size': 14,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        subtitle_format = workbook.add_format({
            'font_size': 10,
            'align': 'left',
            'valign': 'vcenter'
        })
        
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#4472C4',
            'font_color': '#FFFFFF',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True
        })
        
        # Data formats
        data_format = workbook.add_format({
            'border': 1,
            'text_wrap': True,
            'valign': 'top'
        })
        
        status_active_format = workbook.add_format({
            'border': 1,
            'text_wrap': True,
            'valign': 'top',
            'bg_color': '#C6EFCE',  # Light green
            'font_color': '#006100'  # Dark green
        })
        
        status_inactive_format = workbook.add_format({
            'border': 1,
            'text_wrap': True,
            'valign': 'top',
            'bg_color': '#FFC7CE',  # Light red
            'font_color': '#9C0006'  # Dark red
        })
        
        # Write title and metadata
        if session.get('role') == 0:  # Admin
            worksheet.merge_range('A1:J1', 'LAPORAN DATA ALKOM', title_format)
        else:
            worksheet.merge_range('A1:I1', 'LAPORAN DATA ALKOM', title_format)
        
        current_row = 1
        if search:
            worksheet.write(current_row, 0, f"Filter: {search}", subtitle_format)
            current_row += 1
            
        worksheet.write(current_row, 0, f"Tanggal Export: {datetime.now().strftime('%d-%m-%Y %H:%M')}", subtitle_format)
        current_row += 2
        
        # Write headers
        headers = [
            "No",
            "Perwakilan" if session.get('role') == 0 else None,
            "Provider",
            "Jenis Telpon Satelit",
            "Nomor Telp",
            "Status",
            "Fasilitas Internet",
            "Pengadaan (Tahun)",
            "Pencatatan BMN (Tahun)",
            "No. BMN"
        ]
        # Filter out None values (for non-admin users)
        headers = [h for h in headers if h is not None]
        
        for col, header in enumerate(headers):
            worksheet.write(current_row, col, header, header_format)
        
        current_row += 1
        
        # Write data
        for idx, item in enumerate(alkom_list, 1):
            row = [
                idx,
                item[13] if session.get('role') == 0 else None,  # Perwakilan
                item[1] or '-',   # Provider
                item[2] or '-',   # Jenis Telpon Satelit
                item[3] or '-',   # Nomor Telp
                item[4] or '-',   # Status
                item[5] or '-',   # Fasilitas Internet
                f"{item[7] or '-'} ({item[8] or ''})",  # Pengadaan + Tahun
                f"{item[9] or '-'} ({item[10] or ''})",  # Pencatatan BMN + Tahun
                item[11] or '-'   # No BMN
            ]
            # Filter out None values (for non-admin users)
            row = [r for r in row if r is not None]
            
            for col, value in enumerate(row):
                if col == headers.index("Status") - (0 if session.get('role') == 0 else 1):
                    if item[4] == 'aktif':
                        worksheet.write(current_row, col, value, status_active_format)
                    else:
                        worksheet.write(current_row, col, value, status_inactive_format)
                else:
                    worksheet.write(current_row, col, value, data_format)
            
            current_row += 1
        
        # Auto-adjust column widths based on content
        for i, header in enumerate(headers):
            max_length = len(header)
            for row in alkom_list:
                if i == 0:  # No column
                    val_length = len(str(row[0]))
                elif i == 1 and session.get('role') == 0:  # Perwakilan
                    val_length = len(str(row[13] or ''))
                elif (i == 1 and session.get('role') != 0) or (i == 2 and session.get('role') == 0):
                    # Provider column (shifted for non-admin)
                    val_length = len(str(row[1] or ''))
                elif (i == 2 and session.get('role') != 0) or (i == 3 and session.get('role') == 0):
                    # Jenis Telpon Satelit column
                    val_length = len(str(row[2] or ''))
                elif (i == 3 and session.get('role') != 0) or (i == 4 and session.get('role') == 0):
                    # Nomor Telp column
                    val_length = len(str(row[3] or ''))
                elif (i == 4 and session.get('role') != 0) or (i == 5 and session.get('role') == 0):
                    # Status column
                    val_length = len(str(row[4] or ''))
                elif (i == 5 and session.get('role') != 0) or (i == 6 and session.get('role') == 0):
                    # Fasilitas Internet column
                    val_length = len(str(row[5] or ''))
                elif (i == 6 and session.get('role') != 0) or (i == 7 and session.get('role') == 0):
                    # Pengadaan column
                    val_length = len(f"{row[7] or ''} {row[8] or ''}")
                elif (i == 7 and session.get('role') != 0) or (i == 8 and session.get('role') == 0):
                    # Pencatatan BMN column
                    val_length = len(f"{row[9] or ''} {row[10] or ''}")
                elif (i == 8 and session.get('role') != 0) or (i == 9 and session.get('role') == 0):
                    # No BMN column
                    val_length = len(str(row[11] or ''))
                
                if val_length > max_length:
                    max_length = val_length
            
            # Set column width with some padding
            worksheet.set_column(i, i, min(max_length + 2, 50))  # Max width 50 chars
        
        # Freeze header row
        worksheet.freeze_panes(1, 0)
        
        workbook.close()
        output.seek(0)
        
        return send_file(
            output,
            as_attachment=True,
            download_name=f"data_alkom_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        print(f"Excel Export Error: {str(e)}")
        flash('Gagal membuat Excel: ' + str(e), 'error')
        return redirect(url_for('list_alkom'))

@app.route('/alkom')
def list_alkom():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    # Get search and pagination parameters
    search = request.args.get('search', '').strip()
    page = request.args.get('page', 1, type=int)
    per_page = 20
    sort_column = request.args.get('sort', 'no')
    sort_direction = request.args.get('dir', 'asc')

    # Validate sort column
    valid_columns = {
        'no': 'a.no',
        'perwakilan': 'r.nama_perwakilan',
        'provider': 'a.provider',
        'jenis_telpon_satelit': 'a.jenis_telpon_satelit',
        'status_alkom': 'a.status_alkom',
        'tahun_pengadaan': 'a.tahun_pengadaan'
    }
    sort_column = valid_columns.get(sort_column, 'a.no')
    
    # Validate sort direction
    sort_direction = 'DESC' if sort_direction.lower() == 'desc' else 'ASC'

    # Base query with joins
    query = """
    SELECT a.no, a.provider, a.jenis_telpon_satelit, a.nomor_telp_satelit,
           a.status_alkom, a.fasilitas_internet, a.status_langganan,
           a.pengadaan, a.tahun_pengadaan, a.pencatatan_BMN, a.tahun_pencatatan,
           a.no_bmn, r.trigram as id_pwk, r.nama_perwakilan
    FROM tabel_Alkom a
    LEFT JOIN ref_perwakilan r ON a.perwakilan = r.trigram
    WHERE (a.provider ILIKE %s OR
          a.jenis_telpon_satelit ILIKE %s OR
          a.nomor_telp_satelit ILIKE %s OR
          r.nama_perwakilan ILIKE %s OR
          a.pengadaan ILIKE %s OR
          a.no_bmn ILIKE %s)
    """

    # Parameters for the query
    search_param = f'%{search}%'
    params = [search_param for _ in range(5)]
    params = [search_param for _ in range(6)]

    # Add perwakilan filter for non-admin users
    if session.get('role') != 0:  # If not admin
        user_trigram = session.get('trigram')
        if user_trigram:
            query += " AND a.perwakilan = %s"
            params.append(user_trigram)

    # Add sorting
    query += f" ORDER BY {sort_column} {sort_direction}"

    # Count query for pagination
    count_query = query.replace(
        "SELECT a.no, a.provider, a.jenis_telpon_satelit, a.nomor_telp_satelit, " +
        "a.status_alkom, a.fasilitas_internet, a.status_langganan, " +
        "a.pengadaan, a.tahun_pengadaan, a.pencatatan_BMN, a.tahun_pencatatan, " +
        "r.trigram as id_pwk, r.nama_perwakilan", 
        "SELECT COUNT(*)"
    ).split("ORDER BY")[0]  # Remove ORDER BY for count query

    # Get total count
    count_result = execute_query(count_query, params, fetch_one=True)
    total = count_result[0] if count_result else 0
    
    # Add pagination to main query
    paginated_query = query + f" LIMIT {per_page} OFFSET {(page - 1) * per_page}"
    
    # Execute query
    alkom_list = execute_query(paginated_query, params, fetch=True) or []
    
    total_pages = (total + per_page - 1) // per_page if total > 0 else 1
    
    return render_template('alkom/list.html',
                         alkom_list=alkom_list,
                         search=search,
                         page=page,
                         per_page=per_page,
                         total=total,
                         total_pages=total_pages,
                         sort_column=sort_column.replace('a.', ''),
                         sort_direction=sort_direction)

@app.route('/alkom/create', methods=['GET', 'POST'])
def create_alkom():
    if 'user_id' not in session:
        return redirect(url_for('login'))

    # Get perwakilan list based on user role
    if session.get('role') == 0:  # Admin
        perwakilan_list = execute_query(
            "SELECT trigram, nama_perwakilan FROM ref_perwakilan ORDER BY nama_perwakilan",
            fetch=True
        ) or []
    else:  # Regular user
        perwakilan_list = execute_query(
            "SELECT trigram, nama_perwakilan FROM ref_perwakilan WHERE trigram = %s",
            (session.get('trigram'),),
            fetch=True
        ) or []

    if request.method == 'POST':
        try:
            form_data = request.form
            
            # Prepare data
            data = (
                form_data.get('perwakilan', '').strip().upper(),
                form_data.get('provider', '').strip(),
                form_data.get('jenis_telpon_satelit', '').strip(),
                form_data.get('nomor_telp_satelit', '').strip(),
                form_data.get('status_alkom', '').strip(),
                form_data.get('fasilitas_internet', '').strip(),
                form_data.get('status_langganan', '').strip(),
                form_data.get('pengadaan', '').strip(),
                int(form_data.get('tahun_pengadaan')) if form_data.get('tahun_pengadaan') else None,
                form_data.get('pencatatan_BMN', '').strip(),
                int(form_data.get('tahun_pencatatan')) if form_data.get('tahun_pencatatan') else None,
                form_data.get('no_bmn', '').strip(),  # Tambahkan no_bmn
                session.get('user_id')
            )

            # Validate required fields
            if not data[0]:  # perwakilan
                flash('Perwakilan wajib diisi', 'error')
                return render_template('alkom/create.html', 
                                    perwakilan_list=perwakilan_list,
                                    form_data=request.form,
                                    status_options=['aktif', 'tidak aktif'],
                                    fasilitas_options=['ya', 'tidak'],
                                    langganan_options=['pra bayar', 'pasca bayar'],
                                    pengadaan_options=['pusat', 'perwakilan'],
                                    bmn_options=['pusat', 'perwakilan'])

            # For non-admin users, verify they're not trying to create for another perwakilan
            if session.get('role') != 0 and data[0] != session.get('trigram'):
                flash('Anda hanya bisa membuat data untuk perwakilan Anda sendiri', 'error')
                return render_template('alkom/create.html', 
                                    perwakilan_list=perwakilan_list,
                                    form_data=request.form,
                                    status_options=['aktif', 'tidak aktif'],
                                    fasilitas_options=['ya', 'tidak'],
                                    langganan_options=['pra bayar', 'pasca bayar'],
                                    pengadaan_options=['pusat', 'perwakilan'],
                                    bmn_options=['pusat', 'perwakilan'])

            # Execute query
            success = execute_query("""
                INSERT INTO tabel_Alkom 
                (perwakilan, provider, jenis_telpon_satelit, nomor_telp_satelit,
                status_alkom, fasilitas_internet, status_langganan, pengadaan,
                tahun_pengadaan, pencatatan_BMN, tahun_pencatatan, no_bmn, user_input)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """, data, commit=True)
            
            if success:
                flash('Data Alkom berhasil ditambahkan', 'success')
                return redirect(url_for('list_alkom'))
            else:
                flash('Gagal menambahkan data Alkom', 'error')
        except Exception as e:
            logger.error(f"Error: {str(e)}")
            flash('Terjadi kesalahan saat menyimpan data', 'error')

    return render_template('alkom/create.html', 
                         perwakilan_list=perwakilan_list,
                         status_options=['aktif', 'tidak aktif'],
                         fasilitas_options=['ya', 'tidak'],
                         langganan_options=['pra bayar', 'pasca bayar'],
                         pengadaan_options=['pusat', 'perwakilan'],
                         bmn_options=['pusat', 'perwakilan'])

@app.route('/alkom/edit/<int:no>', methods=['GET', 'POST'])
def edit_alkom(no):
    if 'user_id' not in session:
        return redirect(url_for('login'))

    # Get existing data with joins
    alkom = execute_query(
        """SELECT a.no, a.perwakilan, a.provider, a.jenis_telpon_satelit,
                  a.nomor_telp_satelit, a.status_alkom, a.fasilitas_internet,
                  a.status_langganan, a.pengadaan, a.tahun_pengadaan,
                  a.pencatatan_BMN, a.tahun_pencatatan, r.nama_perwakilan
           FROM tabel_Alkom a
           LEFT JOIN ref_perwakilan r ON a.perwakilan = r.trigram
           WHERE a.no = %s""",
        (no,),
        fetch_one=True
    )
    
    if not alkom:
        flash('Data Alkom tidak ditemukan', 'error')
        return redirect(url_for('list_alkom'))

    # Authorization check for non-admin users
    if session.get('role') != 0:  # If not admin
        user_trigram = session.get('trigram')
        if user_trigram != alkom[1]:  # alkom[1] is perwakilan
            flash('Anda tidak memiliki akses untuk mengedit data ini', 'error')
            return redirect(url_for('list_alkom'))

    # Get perwakilan list based on user role
    if session.get('role') == 0:  # Admin
        perwakilan_list = execute_query(
            "SELECT trigram, nama_perwakilan FROM ref_perwakilan ORDER BY nama_perwakilan",
            fetch=True
        ) or []
    else:  # Regular user
        perwakilan_list = execute_query(
            "SELECT trigram, nama_perwakilan FROM ref_perwakilan WHERE trigram = %s",
            (session.get('trigram'),),
            fetch=True
        ) or []

    if request.method == 'POST':
        try:
            # Get form data
            form_data = request.form
            
            # Prepare data tuple with proper null handling
            data = (
                form_data.get('perwakilan', '').strip().upper(),
                form_data.get('provider', '').strip(),
                form_data.get('jenis_telpon_satelit', '').strip(),
                form_data.get('nomor_telp_satelit', '').strip(),
                form_data.get('status_alkom', '').strip(),
                form_data.get('fasilitas_internet', '').strip(),
                form_data.get('status_langganan', '').strip(),
                form_data.get('pengadaan', '').strip(),
                int(form_data.get('tahun_pengadaan')) if form_data.get('tahun_pengadaan') else None,
                form_data.get('pencatatan_BMN', '').strip(),
                int(form_data.get('tahun_pencatatan')) if form_data.get('tahun_pencatatan') else None,
                form_data.get('no_bmn', '').strip(),  # Tambahkan no_bmn
                session.get('user_id'),
                no
            )

            # Validate required fields
            if not data[0]:  # perwakilan
                flash('Perwakilan wajib diisi', 'error')
                return redirect(url_for('edit_alkom', no=no))

            # For non-admin users, verify they're not trying to change to another perwakilan
            if session.get('role') != 0 and data[0] != session.get('trigram'):
                flash('Anda hanya bisa mengubah data untuk perwakilan Anda sendiri', 'error')
                return redirect(url_for('edit_alkom', no=no))

            # Execute query
            success = execute_query("""
                UPDATE tabel_Alkom SET
                    perwakilan = %s,
                    provider = %s,
                    jenis_telpon_satelit = %s,
                    nomor_telp_satelit = %s,
                    status_alkom = %s,
                    fasilitas_internet = %s,
                    status_langganan = %s,
                    pengadaan = %s,
                    tahun_pengadaan = %s,
                    pencatatan_BMN = %s,
                    tahun_pencatatan = %s,
                    no_bmn = %s,
                    user_update = %s,
                    date_update = CURRENT_TIMESTAMP
                WHERE no = %s
            """, data, commit=True)
            
            if success:
                flash('Data Alkom berhasil diperbarui', 'success')
                return redirect(url_for('list_alkom'))
            else:
                flash('Gagal memperbarui data Alkom', 'error')
        except ValueError as e:
            logger.error(f"ValueError updating alkom: {str(e)}")
            flash('Format data tidak valid. Pastikan semua field diisi dengan benar', 'error')
        except Exception as e:
            logger.error(f"Error updating alkom: {str(e)}")
            flash('Terjadi kesalahan saat memperbarui data', 'error')

    return render_template('alkom/edit.html', 
                         alkom=alkom,
                         perwakilan_list=perwakilan_list,
                         status_options=['aktif', 'tidak aktif'],
                         fasilitas_options=['ya', 'tidak'],
                         langganan_options=['pra bayar', 'pasca bayar'],
                         pengadaan_options=['pusat', 'perwakilan'],
                         bmn_options=['pusat', 'perwakilan'])

@app.route('/alkom/delete/<int:no>', methods=['POST'])
def delete_alkom(no):
    if 'user_id' not in session:
        return redirect(url_for('login'))

    # First check if the user has permission to delete this record
    if session.get('role') != 0:  # If not admin
        alkom = execute_query(
            "SELECT perwakilan FROM tabel_Alkom WHERE no = %s",
            (no,),
            fetch_one=True
        )
        
        if not alkom or alkom[0] != session.get('trigram'):
            flash('Anda tidak memiliki akses untuk menghapus data ini', 'error')
            return redirect(url_for('list_alkom'))

    try:
        success = execute_query(
            "DELETE FROM tabel_Alkom WHERE no = %s",
            (no,),
            commit=True
        )
        
        if success:
            flash('Data Alkom berhasil dihapus', 'success')
        else:
            flash('Gagal menghapus data Alkom', 'error')
    except Exception as e:
        logger.error(f"Error deleting alkom: {str(e)}")
        flash('Terjadi kesalahan saat menghapus data', 'error')
    
    return redirect(url_for('list_alkom'))

# ===========================================
# TIPE PALSAN
# ===========================================
@app.route('/tipe_palsan')
def list_tipe_palsan():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    # Get search and pagination parameters
    search = request.args.get('search', '').strip()
    page = request.args.get('page', 1, type=int)
    per_page = 20
    sort_column = request.args.get('sort', 'id_tipe')
    sort_direction = request.args.get('dir', 'asc')

    # Validate sort column
    valid_columns = {
        'id_tipe': 'id_tipe',
        'nama': 'nama'
    }
    sort_column = valid_columns.get(sort_column, 'id_tipe')
    
    # Validate sort direction
    sort_direction = 'DESC' if sort_direction.lower() == 'desc' else 'ASC'

    # Base query
    query = """
        SELECT id_tipe, nama 
        FROM tipe_palsan
        WHERE (id_tipe ILIKE %s OR 
              nama ILIKE %s)
    """

    # Parameters for the query
    search_param = f'%{search}%'
    params = [search_param, search_param]

    # Add sorting
    query += f" ORDER BY {sort_column} {sort_direction}"

    # Count query for pagination
    count_query = query.replace(
        "SELECT id_tipe, nama", 
        "SELECT COUNT(*)"
    ).split("ORDER BY")[0]

    # Get total count
    count_result = execute_query(count_query, params, fetch_one=True)
    total = int(count_result[0]) if count_result and count_result[0] else 0
    
    # Add pagination to main query
    paginated_query = query + f" LIMIT {per_page} OFFSET {(page - 1) * per_page}"
    
    # Execute query
    tipe_palsan_list = execute_query(paginated_query, params, fetch=True) or []
    
    total_pages = (total + per_page - 1) // per_page if total > 0 else 1
    
    return render_template('tipe_palsan/list.html',
                         tipe_palsan_list=tipe_palsan_list,
                         search=search,
                         page=page,
                         per_page=per_page,
                         total=total,
                         total_pages=total_pages,
                         sort_column=sort_column,
                         sort_direction=sort_direction)

@app.route('/tipe_palsan/create', methods=['GET', 'POST'])
def create_tipe_palsan():
    if 'user_id' not in session:
        return redirect(url_for('login'))

    if request.method == 'POST':
        try:
            nama = request.form.get('nama', '').strip()
            
            # Validate required fields
            if not nama:
                flash('Nama wajib diisi', 'error')
                return render_template('tipe_palsan/create.html', 
                                    form_data=request.form)

            # Generate ID otomatis
            id_tipe = generate_tipe_palsan_id()

            # Insert with user tracking
            success = execute_query(
                """INSERT INTO tipe_palsan 
                   (id_tipe, nama, user_input, user_update) 
                   VALUES (%s, %s, %s, %s)""",
                (id_tipe, nama, session.get('user_id'), session.get('user_id')),
                commit=True
            )
            
            if success:
                flash('Tipe Palsan berhasil ditambahkan', 'success')
                return redirect(url_for('list_tipe_palsan'))
            else:
                flash('Gagal menambahkan Tipe Palsan', 'error')
        except Exception as e:
            logger.error(f"Error creating tipe palsan: {str(e)}")
            flash('Terjadi kesalahan saat menyimpan data', 'error')

    return render_template('tipe_palsan/create.html')

@app.route('/tipe_palsan/edit/<id_tipe>', methods=['GET', 'POST'])
def edit_tipe_palsan(id_tipe):
    if 'user_id' not in session:
        return redirect(url_for('login'))

    # Get existing data
    tipe = execute_query(
        "SELECT id_tipe, nama FROM tipe_palsan WHERE id_tipe = %s",
        (id_tipe,),
        fetch_one=True
    )
    
    if not tipe:
        flash('Tipe Palsan tidak ditemukan', 'error')
        return redirect(url_for('list_tipe_palsan'))

    if request.method == 'POST':
        try:
            nama = request.form.get('nama', '').strip()

            # Validate required fields
            if not nama:
                flash('Nama wajib diisi', 'error')
                return redirect(url_for('edit_tipe_palsan', id_tipe=id_tipe))

            # Update with user tracking
            success = execute_query(
                """UPDATE tipe_palsan SET 
                   nama = %s, 
                   user_update = %s,
                   date_update = CURRENT_TIMESTAMP 
                   WHERE id_tipe = %s""",
                (nama, session.get('user_id'), id_tipe),
                commit=True
            )
            
            if success:
                flash('Tipe Palsan berhasil diperbarui', 'success')
                return redirect(url_for('list_tipe_palsan'))
            else:
                flash('Gagal memperbarui Tipe Palsan', 'error')
        except Exception as e:
            logger.error(f"Error updating tipe palsan: {str(e)}")
            flash('Terjadi kesalahan saat memperbarui data', 'error')

    return render_template('tipe_palsan/edit.html', 
                         tipe=tipe)

@app.route('/tipe_palsan/delete/<id_tipe>', methods=['POST'])
def delete_tipe_palsan(id_tipe):
    if 'user_id' not in session:
        return redirect(url_for('login'))

    try:
        # Cek referential integrity
        used_in_palsan = execute_query(
            "SELECT 1 FROM palsan WHERE id_tipe = %s LIMIT 1",
            (id_tipe,),
            fetch_one=True
        )
        
        if used_in_palsan:
            flash('Tipe Palsan tidak dapat dihapus karena sudah digunakan', 'error')
            return redirect(url_for('list_tipe_palsan'))

        success = execute_query(
            "DELETE FROM tipe_palsan WHERE id_tipe = %s",
            (id_tipe,),
            commit=True
        )
        
        if success:
            flash('Tipe Palsan berhasil dihapus', 'success')
        else:
            flash('Gagal menghapus Tipe Palsan', 'error')
    except Exception as e:
        logger.error(f"Error deleting tipe palsan: {str(e)}")
        flash('Terjadi kesalahan saat menghapus data', 'error')
    
    return redirect(url_for('list_tipe_palsan'))

# ===========================================
# PALSAN CRUD ROUTES
# ===========================================
@app.route('/palsan/export/pdf')
def export_palsan_pdf():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    try:
        # Get all parameters from the request
        search = request.args.get('search', '').strip()
        sort_column = request.args.get('sort', 'id_palsan')
        sort_direction = request.args.get('dir', 'asc')

        # Validate sort column
        valid_columns = {
            'id_palsan': 'p.id_palsan',
            'serial_number': 'p.serial_number',
            'status': 'p.status',
            'perwakilan': 'r.nama_perwakilan',
            'tipe': 't.nama'
        }
        sort_column = valid_columns.get(sort_column, 'p.id_palsan')
        sort_direction = 'DESC' if sort_direction.lower() == 'desc' else 'ASC'

        # Base query with joins
        query = """
            SELECT p.id_palsan, p.serial_number, p.status, 
                   p.id_pwk, r.nama_perwakilan,
                   p.id_tipe, t.nama as tipe_palsan,
                   p.pengadaan, p.tahun_pengadaan,
                   p.pencatatan, p.tahun_pencatatan
            FROM tabel_palsan p
            LEFT JOIN ref_perwakilan r ON p.id_pwk = r.trigram
            LEFT JOIN tipe_palsan t ON p.id_tipe = t.id_tipe
            WHERE (p.id_palsan ILIKE %s OR 
                  p.serial_number ILIKE %s OR
                  p.status ILIKE %s OR
                  r.nama_perwakilan ILIKE %s OR
                  t.nama ILIKE %s)
        """

        # Parameters for the query
        search_param = f'%{search}%'
        params = [search_param for _ in range(5)]

        # Add perwakilan filter for non-admin users
        if session.get('role') != 0:  # If not admin
            user_trigram = session.get('trigram')
            if user_trigram:
                query += " AND p.id_pwk = %s"
                params.append(user_trigram)

        # Sort by the selected column
        query += f" ORDER BY {sort_column} {sort_direction}"

        # Execute query to get all data
        palsan_list = execute_query(query, params, fetch=True) or []

        # Create PDF with A4 size (portrait orientation)
        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, 
                              pagesize=A4,  # Changed to A4
                              leftMargin=15,
                              rightMargin=15,
                              topMargin=20,
                              bottomMargin=20)
        elements = []
        
        # Styles configuration
        styles = getSampleStyleSheet()
        
        # Title style
        title_style = ParagraphStyle(
            'Title',
            parent=styles['Title'],
            fontSize=14,
            alignment=TA_CENTER,
            spaceAfter=12
        )
        
        # Metadata style
        meta_style = ParagraphStyle(
            'Meta',
            parent=styles['Normal'],
            fontSize=9,
            spaceAfter=6
        )
        
        # Add title
        elements.append(Paragraph("LAPORAN DATA PALSAN", title_style))
        
        # Add metadata
        metadata = [
            f"Tanggal Export: {datetime.now().strftime('%d-%m-%Y %H:%M')}"
        ]
        if search:
            metadata.insert(0, f"Filter: {search}")
        
        for meta in metadata:
            elements.append(Paragraph(meta, meta_style))
        
        # Prepare table data
        data = []
        headers = [
            "No", 
            "ID Palsan", 
            "Perwakilan", 
            "Serial Number", 
            "Tipe Palsan", 
            "Status",
            "Tahun Pengadaan"
        ]
        data.append(headers)
        
        for idx, item in enumerate(palsan_list, 1):
            row = [
                str(idx),
                item[0] or '-',
                item[4] or '-',
                item[1] or '-',
                item[6] or '-',
                item[2] or '-',
                item[8] or '-'
            ]
            data.append(row)
        
        # Calculate available width (using A4 width)
        available_width = A4[0] - doc.leftMargin - doc.rightMargin
        
        # Define column distribution (adjust these ratios as needed)
        col_ratios = [0.05, 0.1, 0.2, 0.15, 0.2, 0.15, 0.15]  # Sum should be 1.0
        col_widths = [available_width * ratio for ratio in col_ratios]
        
        # Ensure minimum widths
        min_widths = [30, 50, 80, 70, 80, 60, 60]  # Minimum widths in points
        col_widths = [max(w, min_w) for w, min_w in zip(col_widths, min_widths)]
        
        # Create table
        table = Table(data, 
                     colWidths=col_widths,
                     repeatRows=1,
                     hAlign='LEFT')
        
        # Table style
        style = TableStyle([
            # Header style
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#4472C4')),
            ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,0), 8),
            ('ALIGN', (0,0), (-1,0), 'CENTER'),
            ('BOTTOMPADDING', (0,0), (-1,0), 6),
            
            # Cell style
            ('FONTSIZE', (0,1), (-1,-1), 8),
            ('LEADING', (0,0), (-1,-1), 9),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey),
            ('WORDWRAP', (0,0), (-1,-1), True),
            
            # Alternating row colors
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.whitesmoke]),
            
            # Specific column alignments
            ('ALIGN', (0,0), (0,-1), 'CENTER'),    # No column centered
            ('ALIGN', (1,0), (1,-1), 'CENTER'),    # ID Palsan centered
            ('ALIGN', (2,0), (2,-1), 'LEFT'),      # Perwakilan left-aligned
            ('ALIGN', (3,0), (3,-1), 'LEFT'),      # Serial Number left-aligned
            ('ALIGN', (4,0), (4,-1), 'LEFT'),      # Tipe Palsan left-aligned
            ('ALIGN', (5,0), (5,-1), 'CENTER'),    # Status centered
            ('ALIGN', (6,0), (6,-1), 'CENTER')     # Tahun Pengadaan centered
        ])
        
        table.setStyle(style)
        elements.append(table)
        
        # Build PDF
        doc.build(elements)
        buffer.seek(0)
        
        return send_file(
            buffer,
            as_attachment=True,
            download_name=f"data_palsan_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
            mimetype='application/pdf'
        )
        
    except Exception as e:
        print(f"PDF Export Error: {str(e)}")
        flash('Gagal membuat PDF: ' + str(e), 'error')
        return redirect(url_for('list_palsan'))

@app.route('/palsan/export/excel')
def export_palsan_excel():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    try:
        # Get all parameters from the request
        search = request.args.get('search', '').strip()
        sort_column = request.args.get('sort', 'id_palsan')
        sort_direction = request.args.get('dir', 'asc')

        # Validate sort column
        valid_columns = {
            'id_palsan': 'p.id_palsan',
            'serial_number': 'p.serial_number',
            'status': 'p.status',
            'perwakilan': 'r.nama_perwakilan',
            'tipe': 't.nama'
        }
        sort_column = valid_columns.get(sort_column, 'p.id_palsan')
        sort_direction = 'DESC' if sort_direction.lower() == 'desc' else 'ASC'

        # Base query with joins
        query = """
            SELECT p.id_palsan, p.serial_number, p.status, 
                   p.id_pwk, r.nama_perwakilan,
                   p.id_tipe, t.nama as tipe_palsan,
                   p.pengadaan, p.tahun_pengadaan,
                   p.pencatatan, p.tahun_pencatatan
            FROM tabel_palsan p
            LEFT JOIN ref_perwakilan r ON p.id_pwk = r.trigram
            LEFT JOIN tipe_palsan t ON p.id_tipe = t.id_tipe
            WHERE (p.id_palsan ILIKE %s OR 
                  p.serial_number ILIKE %s OR
                  p.status ILIKE %s OR
                  r.nama_perwakilan ILIKE %s OR
                  t.nama ILIKE %s)
        """

        # Parameters for the query
        search_param = f'%{search}%'
        params = [search_param for _ in range(5)]

        # Add perwakilan filter for non-admin users
        if session.get('role') != 0:  # If not admin
            user_trigram = session.get('trigram')
            if user_trigram:
                query += " AND p.id_pwk = %s"
                params.append(user_trigram)

        # Sort by the selected column
        query += f" ORDER BY {sort_column} {sort_direction}"

        # Execute query to get all data
        palsan_list = execute_query(query, params, fetch=True) or []

        # Create Excel file in memory
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet("Data Palsan")
        
        # Add formats
        title_format = workbook.add_format({
            'bold': True,
            'font_size': 14,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        subtitle_format = workbook.add_format({
            'font_size': 10,
            'align': 'left',
            'valign': 'vcenter'
        })
        
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#4472C4',
            'font_color': '#FFFFFF',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True
        })
        
        # Data formats
        data_format = workbook.add_format({
            'border': 1,
            'text_wrap': True,
            'valign': 'top'
        })
        
        center_format = workbook.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        # Write title and metadata
        worksheet.merge_range('A1:G1', 'LAPORAN DATA PALSAN', title_format)
        
        current_row = 1
        if search:
            worksheet.write(current_row, 0, f"Filter: {search}", subtitle_format)
            current_row += 1
            
        worksheet.write(current_row, 0, f"Tanggal Export: {datetime.now().strftime('%d-%m-%Y %H:%M')}", subtitle_format)
        current_row += 2
        
        # Write headers
        headers = [
            "No", 
            "ID Palsan", 
            "Perwakilan", 
            "Serial Number", 
            "Tipe Palsan", 
            "Status",
            "Tahun Pengadaan"
        ]
        
        for col, header in enumerate(headers):
            worksheet.write(current_row, col, header, header_format)
        
        current_row += 1
        
        # Write data
        for idx, item in enumerate(palsan_list, 1):
            row = [
                idx,
                item[0] or '-',
                item[4] or '-',
                item[1] or '-',
                item[6] or '-',
                item[2] or '-',
                item[8] or '-'
            ]
            
            # Write each cell with appropriate format
            worksheet.write(current_row, 0, row[0], center_format)  # No
            worksheet.write(current_row, 1, row[1], center_format)  # ID Palsan
            worksheet.write(current_row, 2, row[2], data_format)   # Perwakilan
            worksheet.write(current_row, 3, row[3], data_format)   # Serial Number
            worksheet.write(current_row, 4, row[4], data_format)   # Tipe Palsan
            worksheet.write(current_row, 5, row[5], center_format) # Status
            worksheet.write(current_row, 6, row[6], center_format) # Tahun Pengadaan
            
            current_row += 1
        
        # Auto-adjust column widths based on content
        col_widths = [
            ('No', 5),
            ('ID Palsan', 15),
            ('Perwakilan', 25),
            ('Serial Number', 20),
            ('Tipe Palsan', 25),
            ('Status', 15),
            ('Tahun Pengadaan', 15)
        ]
        
        for i, (header, default_width) in enumerate(col_widths):
            max_length = default_width
            for row in palsan_list:
                if i == 0:  # No column
                    val_length = len(str(idx))
                elif i == 1:  # ID Palsan
                    val_length = len(str(row[0] or ''))
                elif i == 2:  # Perwakilan
                    val_length = len(str(row[4] or ''))
                elif i == 3:  # Serial Number
                    val_length = len(str(row[1] or ''))
                elif i == 4:  # Tipe Palsan
                    val_length = len(str(row[6] or ''))
                elif i == 5:  # Status
                    val_length = len(str(row[2] or ''))
                elif i == 6:  # Tahun Pengadaan
                    val_length = len(str(row[8] or ''))
                
                if val_length > max_length:
                    max_length = val_length
            
            # Set column width with some padding
            worksheet.set_column(i, i, min(max_length + 2, 50))  # Max width 50 chars
        
        workbook.close()
        output.seek(0)
        
        return send_file(
            output,
            as_attachment=True,
            download_name=f"data_palsan_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        print(f"Excel Export Error: {str(e)}")
        flash('Gagal membuat Excel: ' + str(e), 'error')
        return redirect(url_for('list_palsan'))

@app.route('/palsan')
def list_palsan():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    # Get search and pagination parameters
    search = request.args.get('search', '').strip()
    page = request.args.get('page', 1, type=int)
    per_page = 20
    sort_column = request.args.get('sort', 'id_palsan')
    sort_direction = request.args.get('dir', 'asc')

    # Validate sort column
    valid_columns = {
        'id_palsan': 'p.id_palsan',
        'serial_number': 'p.serial_number',
        'status': 'p.status',
        'perwakilan': 'r.nama_perwakilan',
        'tipe': 't.nama',
        'dipinjamkan' : 't.dipinjamkan'
    }
    sort_column = valid_columns.get(sort_column, 'p.id_palsan')
    
    # Validate sort direction
    sort_direction = 'DESC' if sort_direction.lower() == 'desc' else 'ASC'

    # Base query with joins
    query = """
        SELECT p.id_palsan, p.serial_number, p.status, 
            p.id_pwk, r.nama_perwakilan,
            p.id_tipe, t.nama as tipe_palsan,
            p.pengadaan, p.tahun_pengadaan,
            p.pencatatan, p.tahun_pencatatan, p.dipinjamkan
        FROM tabel_palsan p
        LEFT JOIN ref_perwakilan r ON p.id_pwk = r.trigram
        LEFT JOIN tipe_palsan t ON p.id_tipe = t.id_tipe
        WHERE (p.id_palsan ILIKE %s OR 
            p.serial_number ILIKE %s OR
            p.status ILIKE %s OR
            r.nama_perwakilan ILIKE %s OR
            t.nama ILIKE %s)
    """

    # Parameters for the query
    search_param = f'%{search}%'
    params = [search_param for _ in range(5)]

    # Add perwakilan filter for non-admin users
    if session.get('role') != 0:  # If not admin
        user_trigram = session.get('trigram')
        if user_trigram:
            query += " AND p.id_pwk = %s"
            params.append(user_trigram)

    # Add sorting
    query += f" ORDER BY {sort_column} {sort_direction}"

    # Count query for pagination
    count_query = query.replace(
        "SELECT p.id_palsan, p.serial_number, p.status, p.id_pwk, r.nama_perwakilan, p.id_tipe, t.nama as tipe_palsan, p.pengadaan, p.tahun_pengadaan, p.pencatatan, p.tahun_pencatatan, p.dipinjamkan", 
        "SELECT COUNT(*)"
    ).split("ORDER BY")[0]

    # Get total count - handle both integer and string results
    count_result = execute_query(count_query, params, fetch_one=True)
    try:
        total = int(count_result[0]) if count_result and count_result[0] else 0
    except (ValueError, TypeError):
        logger.error(f"Invalid count value: {count_result[0] if count_result else 'None'}")
        total = 0
    
    # Add pagination to main query
    paginated_query = query + f" LIMIT {per_page} OFFSET {(page - 1) * per_page}"
    
    # Execute query
    palsan_list = execute_query(paginated_query, params, fetch=True) or []
    
    total_pages = (total + per_page - 1) // per_page if total > 0 else 1
    
    return render_template('palsan/list.html',
                         palsan_list=palsan_list,
                         search=search,
                         page=page,
                         per_page=per_page,
                         total=total,
                         total_pages=total_pages,
                         sort_column=sort_column.replace('p.', ''),
                         sort_direction=sort_direction)

@app.route('/palsan/create', methods=['GET', 'POST'])
def create_palsan():
    if 'user_id' not in session:
        return redirect(url_for('login'))

    # Get data untuk dropdown
    perwakilan_list = execute_query(
        "SELECT trigram, nama_perwakilan FROM ref_perwakilan ORDER BY nama_perwakilan",
        fetch=True
    ) or []
    
    tipe_palsan_list = execute_query(
        "SELECT id_tipe, nama FROM tipe_palsan ORDER BY nama",
        fetch=True
    ) or []

    if request.method == 'POST':
        try:
            # Ambil data dari form
            form_data = request.form
            
            data = (
                generate_palsan_id(),
                form_data.get('id_pwk', '').strip(),
                form_data.get('id_tipe', '').strip(),
                form_data.get('serial_number', '').strip(),
                form_data.get('pengadaan', '').strip() or None,
                form_data.get('tahun_pengadaan') or None,
                form_data.get('pencatatan', '').strip() or None,
                form_data.get('tahun_pencatatan') or None,
                form_data.get('status', 'Aktif').strip(),
                session.get('user_id'),
                session.get('user_id')
            )

            # Validasi required fields
            if not data[1] or not data[2] or not data[3]:
                flash('Perwakilan, Tipe Palsan, dan Serial Number wajib diisi', 'error')
                return render_template('palsan/create.html',
                                    form_data=request.form,
                                    perwakilan_list=perwakilan_list,
                                    tipe_palsan_list=tipe_palsan_list)

            # Insert data
            success = execute_query(
                """INSERT INTO tabel_palsan 
                   (id_palsan, id_pwk, id_tipe, serial_number, 
                    pengadaan, tahun_pengadaan, pencatatan, tahun_pencatatan,
                    status, user_input, user_update) 
                   VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)""",
                data,
                commit=True
            )
            
            if success:
                flash('Data Palsan berhasil ditambahkan', 'success')
                return redirect(url_for('list_palsan'))
            else:
                flash('Gagal menambahkan data Palsan', 'error')
        except Exception as e:
            logger.error(f"Error creating palsan: {str(e)}", exc_info=True)
            flash(f'Terjadi kesalahan saat menyimpan data: {str(e)}', 'error')

    return render_template('palsan/create.html',
                         perwakilan_list=perwakilan_list,
                         tipe_palsan_list=tipe_palsan_list)

@app.route('/palsan/edit/<id_palsan>', methods=['GET', 'POST'])
def edit_palsan(id_palsan):
    if 'user_id' not in session:
        return redirect(url_for('login'))

    # Get existing data
    palsan = execute_query(
        """SELECT p.id_palsan, p.id_pwk, p.id_tipe, p.serial_number, 
                  p.pengadaan, p.tahun_pengadaan, p.pencatatan, p.tahun_pencatatan,
                  p.status, r.nama_perwakilan, t.nama as tipe_palsan
           FROM tabel_palsan p
           LEFT JOIN ref_perwakilan r ON p.id_pwk = r.trigram
           LEFT JOIN tipe_palsan t ON p.id_tipe = t.id_tipe
           WHERE p.id_palsan = %s""",
        (id_palsan,),
        fetch_one=True
    )
    
    if not palsan:
        flash('Data Palsan tidak ditemukan', 'error')
        return redirect(url_for('list_palsan'))

    # Get data untuk dropdown
    perwakilan_list = execute_query(
        "SELECT trigram, nama_perwakilan FROM ref_perwakilan ORDER BY nama_perwakilan",
        fetch=True
    ) or []
    
    tipe_palsan_list = execute_query(
        "SELECT id_tipe, nama FROM tipe_palsan ORDER BY nama",
        fetch=True
    ) or []

    if request.method == 'POST':
        try:
            # Ambil data dari form
            form_data = request.form
            
            data = (
                form_data.get('id_pwk', '').strip(),
                form_data.get('id_tipe', '').strip(),
                form_data.get('serial_number', '').strip(),
                form_data.get('pengadaan', '').strip() or None,
                form_data.get('tahun_pengadaan') or None,
                form_data.get('pencatatan', '').strip() or None,
                form_data.get('tahun_pencatatan') or None,
                form_data.get('status', 'Aktif').strip(),
                session.get('user_id'),
                id_palsan
            )

            # Validasi required fields
            if not data[0] or not data[1] or not data[2]:
                flash('Perwakilan, Tipe Palsan, dan Serial Number wajib diisi', 'error')
                return redirect(url_for('edit_palsan', id_palsan=id_palsan))

            # Update data
            success = execute_query(
                """UPDATE tabel_palsan SET 
                   id_pwk = %s, 
                   id_tipe = %s,
                   serial_number = %s,
                   pengadaan = %s,
                   tahun_pengadaan = %s,
                   pencatatan = %s,
                   tahun_pencatatan = %s,
                   status = %s,
                   user_update = %s,
                   date_update = CURRENT_TIMESTAMP
                   WHERE id_palsan = %s""",
                data,
                commit=True
            )
            
            if success:
                flash('Data Palsan berhasil diperbarui', 'success')
                return redirect(url_for('list_palsan'))
            else:
                flash('Gagal memperbarui data Palsan', 'error')
        except Exception as e:
            logger.error(f"Error updating palsan: {str(e)}", exc_info=True)
            flash(f'Terjadi kesalahan saat memperbarui data: {str(e)}', 'error')

    return render_template('palsan/edit.html',
                         palsan=palsan,
                         perwakilan_list=perwakilan_list,
                         tipe_palsan_list=tipe_palsan_list)

@app.route('/palsan/delete/<id_palsan>', methods=['POST'])
def delete_palsan(id_palsan):
    if 'user_id' not in session:
        return redirect(url_for('login'))

    try:
        # Check if the palsan has been distributed
        distributed = execute_query(
            "SELECT COUNT(*) FROM distribusi_palsan WHERE id_palsan = %s",
            (id_palsan,),
            fetch_one=True
        )
        
        if distributed and distributed[0] > 0:
            # If distributed, first delete from distribusi_palsan
            execute_query(
                "DELETE FROM distribusi_palsan WHERE id_palsan = %s",
                (id_palsan,),
                commit=True
            )
        
        # Then delete from tabel_palsan
        success = execute_query(
            "DELETE FROM tabel_palsan WHERE id_palsan = %s",
            (id_palsan,),
            commit=True
        )
        
        if success:
            flash('Data Palsan berhasil dihapus', 'success')
        else:
            flash('Gagal menghapus data Palsan', 'error')
    except Exception as e:
        logger.error(f"Error deleting palsan: {str(e)}", exc_info=True)
        flash('Terjadi kesalahan saat menghapus data', 'error')
    
    return redirect(url_for('list_palsan'))

@app.route('/palsan/distribusi', methods=['GET', 'POST'])
def distribusi_palsan():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    # Hanya untuk Kantor Pusat (PJB)
    if session.get('trigram') != 'PJB':
        flash('Hanya admin Kantor Pusat (PJB) yang dapat mengakses distribusi palsan', 'error')
        return redirect(url_for('list_palsan'))

    # Daftar Satker
    satker_list = [
        "PK Satker Wamenlu",
        "PK Satker Sahli",
        "PK Satker Aspasaf",
        "PK Satker Amerop",
        "PK Satker KS Asean",
        "PK Satker Multilateral",
        "PK Satker HPI",
        "PK Satker IDP",
        "PK Satker Itjen",
        "PK Satker BSKLN",
        "PK Satker Sekjen",
        "PK Satker BDSP",
        "PK Satker Protkons",
        "PK Satker PWNI",
        "PK Satker BHAKP",
        "PK Satker BUM",
        "PK Satker Keuangan",
        "PK Satker BSDM",
        "PK Satker BPO",
        "PK Satker Pusdiklat"
    ]

    # Ambil tipe palsan yang tersedia (PJB dan belum didistribusikan)
    tipe_palsan_list = execute_query(
        """SELECT DISTINCT t.id_tipe, t.nama 
           FROM tipe_palsan t
           JOIN tabel_palsan p ON t.id_tipe = p.id_tipe
           WHERE p.id_pwk = 'PJB' 
           AND p.status = 'Aktif'
           AND (p.dipinjamkan = 0 OR p.dipinjamkan IS NULL)""",
        fetch=True
    ) or []

    # Inisialisasi serial numbers
    serial_numbers = []
    selected_tipe = None

    if request.method == 'GET':
        selected_tipe = request.args.get('tipe')
    elif request.method == 'POST':
        selected_tipe = request.form.get('id_tipe')

    if selected_tipe:
        # Ambil serial number yang tersedia untuk tipe yang dipilih
        serial_numbers = execute_query(
            """SELECT p.id_palsan, p.serial_number 
               FROM tabel_palsan p
               WHERE p.id_pwk = 'PJB' 
               AND p.id_tipe = %s
               AND p.status = 'Aktif'
               AND (p.dipinjamkan = 0 OR p.dipinjamkan IS NULL)""",
            (selected_tipe,),
            fetch=True
        ) or []

    if request.method == 'POST':
        try:
            # Ambil data dari form
            id_palsan = request.form.get('serial_number')
            satker_peminjam = request.form.get('satker_peminjam', '').strip()
            nama_peminjam = request.form.get('nama_peminjam', '').strip()
            nip_peminjam = request.form.get('nip_peminjam', '').strip()
            penyerah = request.form.get('penyerah', '').strip()
            nip_penyerah = request.form.get('nip_penyerah', '').strip()

            # Validasi field wajib
            if not id_palsan or not satker_peminjam or not nama_peminjam or not nip_peminjam:
                flash('Serial Number, Satker Peminjam, Nama Peminjam, dan NIP Peminjam wajib diisi', 'error')
                return render_template('palsan/distribusi.html',
                                    tipe_palsan_list=tipe_palsan_list,
                                    serial_numbers=serial_numbers,
                                    satker_list=satker_list,
                                    selected_tipe=selected_tipe,
                                    form_data=request.form)

            # Update status palsan menjadi dipinjamkan
            execute_query(
                """UPDATE tabel_palsan 
                   SET dipinjamkan = 1,
                       user_update = %s,
                       date_update = CURRENT_TIMESTAMP
                   WHERE id_palsan = %s""",
                (session['user_id'], id_palsan),
                commit=True
            )
            
            # Tambahkan record distribusi
            execute_query(
                """INSERT INTO distribusi_palsan 
                   (id_palsan, satker_peminjam, nama_peminjam, nip_peminjam,
                    penyerah, nip_penyerah, user_id) 
                   VALUES (%s, %s, %s, %s, %s, %s, %s)""",
                (id_palsan, satker_peminjam, nama_peminjam, nip_peminjam,
                 penyerah, nip_penyerah, session['user_id']),
                commit=True
            )
            
             # Check which button was clicked
            action = request.form.get('action', 'save')
            
            if action == 'print':
                tipe_palsan = execute_query(
                    "SELECT nama FROM tipe_palsan WHERE id_tipe = %s",
                    (request.form.get('id_tipe'),),
                    fetch_one=True
                )
                
                pdf_data = generate_distribution_pdf({
                    'id_palsan': id_palsan,
                    'tipe_palsan': tipe_palsan[0] if tipe_palsan else request.form.get('id_tipe'),
                    'serial_number': request.form.get('serial_number'),
                    'nama_peminjam': request.form.get('nama_peminjam'),
                    'nip_peminjam': request.form.get('nip_peminjam'),
                    'penyerah': request.form.get('penyerah'),
                    'nip_penyerah': request.form.get('nip_penyerah')
                })
                return send_file(pdf_data, as_attachment=True, download_name=f'tanda_terima_{id_palsan}.pdf')
            
            flash('Data distribusi berhasil disimpan', 'success')
            return redirect(url_for('list_palsan'))

        except Exception as e:
            logger.error(f"Error in distribution: {str(e)}", exc_info=True)
            flash(f'Terjadi kesalahan: {str(e)}', 'error')


    return render_template('palsan/distribusi.html',
                         tipe_palsan_list=tipe_palsan_list,
                         serial_numbers=serial_numbers,
                         satker_list=satker_list,
                         selected_tipe=selected_tipe)

@app.route('/palsan/distribusi/download/<id_palsan>')
def download_distribution_pdf(id_palsan):
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    try:
        # Ambil data distribusi
        distribusi = execute_query(
            """SELECT d.id_palsan, p.serial_number, t.nama as tipe_palsan,
                      d.satker_peminjam, d.nama_peminjam, d.nip_peminjam,
                      d.penyerah, d.nip_penyerah, d.created_at
               FROM distribusi_palsan d
               JOIN tabel_palsan p ON d.id_palsan = p.id_palsan
               JOIN tipe_palsan t ON p.id_tipe = t.id_tipe
               WHERE d.id_palsan = %s""",
            (id_palsan,),
            fetch_one=True
        )
        
        if not distribusi:
            flash('Data distribusi tidak ditemukan', 'error')
            return redirect(url_for('list_palsan'))
        
        # Format data untuk PDF
        pdf_data = {
            'id_palsan': distribusi[0],
            'serial_number': distribusi[1],
            'tipe_palsan': distribusi[2],
            'satker_peminjam': distribusi[3] or '-',
            'nama_peminjam': distribusi[4] or '-',
            'nip_peminjam': distribusi[5] or '-',
            'penyerah': distribusi[6] or '-',
            'nip_penyerah': distribusi[7] or '-',
            'tanggal': distribusi[8].strftime('%d/%m/%Y') if distribusi[8] else datetime.now().strftime('%d/%m/%Y')
        }
        
        # Generate dan kembalikan PDF
        pdf_buffer = generate_distribution_pdf(pdf_data)
        return send_file(
            pdf_buffer,
            as_attachment=True,
            download_name=f"distribusi_{id_palsan}.pdf",
            mimetype='application/pdf'
        )
        
    except Exception as e:
        logger.error(f"PDF download error: {str(e)}", exc_info=True)
        flash(f'Gagal menghasilkan PDF distribusi: {str(e)}', 'error')
        return redirect(url_for('list_palsan'))

# ==============================================
# KATEGORI SISTEM CRUD ROUTES
# ==============================================

@app.route('/kategori-sistem')
def list_kategori_sistem():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    # Get search and pagination parameters
    search = request.args.get('search', '').strip()
    page = request.args.get('page', 1, type=int)
    per_page = 20
    sort_column = request.args.get('sort', 'id')
    sort_direction = request.args.get('dir', 'asc')

    # Validate sort column
    valid_columns = {
        'id': 'id',
        'kategori': 'kategori',
        'keterangan': 'keterangan'
    }
    sort_column = valid_columns.get(sort_column, 'id')
    
    # Validate sort direction
    sort_direction = 'DESC' if sort_direction.lower() == 'desc' else 'ASC'

    # Base query
    query = f"""
        SELECT id, kategori, keterangan
        FROM ref_kategori_sistem
        WHERE kategori ILIKE %s OR keterangan ILIKE %s
        ORDER BY {sort_column} {sort_direction}
    """

    # Count query for pagination
    count_query = """
        SELECT COUNT(*)
        FROM ref_kategori_sistem
        WHERE kategori ILIKE %s OR keterangan ILIKE %s
    """

    search_param = f'%{search}%'
    
    # Get total count
    total = execute_query(count_query, (search_param, search_param), fetch_one=True)[0] or 0
    
    # Add pagination to main query
    paginated_query = query + f" LIMIT {per_page} OFFSET {(page - 1) * per_page}"
    
    # Execute query
    kategori_sistem_list = execute_query(paginated_query, (search_param, search_param), fetch=True) or []
    
    total_pages = (total + per_page - 1) // per_page
    
    return render_template('kategori_sistem/list.html',
                         kategori_sistem_list=kategori_sistem_list,
                         search=search,
                         page=page,
                         per_page=per_page,
                         total=total,
                         total_pages=total_pages,
                         sort_column=sort_column,
                         sort_direction=sort_direction)

@app.route('/kategori-sistem/create', methods=['GET', 'POST'])
def create_kategori_sistem():
    if 'user_id' not in session:
        return redirect(url_for('login'))

    if request.method == 'POST':
        try:
            kategori = request.form.get('kategori', '').strip()
            keterangan = request.form.get('keterangan', '').strip()

            if not kategori:
                flash('Kategori wajib diisi', 'error')
                return render_template('kategori_sistem/create.html', 
                                    form_data=request.form)

            new_id = generate_next_kategori_id()

            success = execute_query(
                "INSERT INTO ref_kategori_sistem (id, kategori, keterangan) VALUES (%s, %s, %s)",
                (new_id, kategori, keterangan),
                commit=True
            )
            
            if success:
                flash('Data berhasil ditambahkan', 'success')
                return redirect(url_for('list_kategori_sistem'))
            else:
                flash('Gagal menambahkan data', 'error')
        except Exception as e:
            logger.error(f"Error: {str(e)}")
            flash(f'Terjadi kesalahan sistem: {str(e)}', 'error')

    return render_template('kategori_sistem/create.html')

@app.route('/kategori-sistem/edit/<string:id>', methods=['GET', 'POST'])
def edit_kategori_sistem(id):
    if 'user_id' not in session:
        return redirect(url_for('login'))

    # Get existing data
    kategori_sistem = execute_query(
        "SELECT id, kategori, keterangan FROM ref_kategori_sistem WHERE id = %s",
        (str(id),),  # Ensure ID is treated as string
        fetch_one=True
    )
    
    if not kategori_sistem:
        flash('Data kategori sistem tidak ditemukan', 'error')
        return redirect(url_for('list_kategori_sistem'))

    if request.method == 'POST':
        try:
            kategori = request.form.get('kategori', '').strip()
            keterangan = request.form.get('keterangan', '').strip()

            # Validate required fields
            if not kategori:
                flash('Kategori wajib diisi', 'error')
                return redirect(url_for('edit_kategori_sistem', id=id))

            success = execute_query("""
                UPDATE ref_kategori_sistem SET
                    kategori = %s,
                    keterangan = %s
                WHERE id = %s
            """, (kategori, keterangan, id), commit=True)
            
            if success:
                flash('Data kategori sistem berhasil diperbarui', 'success')
                return redirect(url_for('list_kategori_sistem'))
            else:
                flash('Gagal memperbarui data kategori sistem', 'error')
        except Exception as e:
            logger.error(f"Error updating kategori sistem: {str(e)}")
            flash('Terjadi kesalahan saat memperbarui data', 'error')

    return render_template('kategori_sistem/edit.html', kategori_sistem=kategori_sistem)

@app.route('/kategori-sistem/delete/<string:id>', methods=['POST'])
def delete_kategori_sistem(id):
    if 'user_id' not in session:
        return redirect(url_for('login'))

    try:
        success = execute_query(
            "DELETE FROM ref_kategori_sistem WHERE id = %s",
            (str(id),),  # Ensure ID is treated as string
            commit=True
        )
        
        if success:
            flash('Data kategori sistem berhasil dihapus', 'success')
        else:
            flash('Gagal menghapus data kategori sistem', 'error')
    except Exception as e:
        logger.error(f"Error deleting kategori sistem: {str(e)}")
        flash('Terjadi kesalahan saat menghapus data', 'error')
    
    return redirect(url_for('list_kategori_sistem'))
# ==============================================
# RUN APPLICATION
# ==============================================

if __name__ == '__main__':
    app.run(debug=True)