from io import BytesIO
import pandas as pd
from werkzeug.utils import secure_filename
import psycopg2
from datetime import datetime
from flask import Flask, render_template, request, redirect, send_file, url_for, session, flash
from uuid import uuid4
import os
import logging

# Configure logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

app = Flask(__name__)
app.secret_key = 'your-secret-key-123'
app.config['UPLOAD_FOLDER'] = 'static/uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Valid jenis perwakilan options
VALID_JENIS_PWK = ['KBRI', 'KJRI', 'PTRI', 'KRI', 'KDEI']

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
    conn = get_db_connection()
    if not conn:
        return None
        
    try:
        with conn.cursor() as cur:
            logger.debug(f"Executing query: {query} with params: {params}")
            cur.execute(query, params or ())
            
            if fetch:
                result = cur.fetchall()
            elif fetch_one:
                result = cur.fetchone()
            elif commit:
                conn.commit()
                logger.debug("Changes committed to database")
                return True  # <--- Tambahkan ini!
            else:
                result = None
                
            return result
    except psycopg2.Error as e:
        conn.rollback()
        logger.error(f'Database error: {e}')
        flash(f'Database error: {e}', 'error')
        return None
    finally:
        if conn:
            conn.close()
            logger.debug("Database connection closed")

def get_next_urutan():
    """Get the next auto-increment value for no_urutan"""
    result = execute_query(
        "SELECT COALESCE(MAX(NO_URUTAN), 0) + 1 FROM REF_PERWAKILAN",
        fetch_one=True
    )
    return result[0] if result else 1

def get_next_no_perwakilan():
    """Get the next auto-increment value for no_perwakilan"""
    result = execute_query(
        "SELECT COALESCE(MAX(NO_PERWAKILAN), 0) + 1 FROM REF_PERWAKILAN",
        fetch_one=True
    )
    return result[0] if result else 1

# ==============================================
# AUTHENTICATION ROUTES
# ==============================================

@app.route('/')
def home():
    return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')

        user = execute_query(
            "SELECT * FROM TABEL_PENGGUNA WHERE USERNAME = %s AND PASSWORD = %s",
            (username, password),
            fetch_one=True
        )

        if user:
            session['user_id'] = user[0]
            session['username'] = user[2]
            session['role'] = user[4]
            return redirect(url_for('dashboard'))
        else:
            flash('Username atau password salah!', 'error')
    
    return render_template('login.html')

@app.route('/register', methods=['GET', 'POST'])
def register():
    if request.method == 'POST':
        # Generate user ID
        max_id = execute_query(
            "SELECT MAX(CAST(SUBSTRING(ID_PENGGUNA, 2) AS INTEGER)) FROM TABEL_PENGGUNA",
            fetch_one=True
        )
        user_id = f'U{(max_id[0] + 1) if max_id and max_id[0] else 1:04}'
        
        data = (
            user_id,
            request.form.get('nama_pengguna'),
            request.form.get('username'),
            request.form.get('password'),
            1,  # Default role
            request.form.get('username'),
            datetime.now()
        )
        
        success = execute_query("""
            INSERT INTO TABEL_PENGGUNA 
            (ID_PENGGUNA, NAMA_PENGGUNA, USERNAME, PASSWORD, ROLE, USER_INPUT, DATE_INPUT)
            VALUES (%s, %s, %s, %s, %s, %s, %s)
        """, data, commit=True)
        
        if success:
            flash('Registrasi berhasil! Silakan login', 'success')
            return redirect(url_for('login'))
    
    return render_template('register.html')

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
# PERWAKILAN CRUD (IMPROVED VERSION)
# ==============================================

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

@app.route('/kepri')
def list_kepri():
    # Logika untuk menampilkan tabel Kepri
    return render_template('kepri/list.html')

@app.route('/personel') 
def list_personel():
    # Logika untuk menampilkan tabel Personel
    return render_template('personel/list.html')

@app.route('/pegawai-setempat')
def list_pegawai_setempat():
    # Logika untuk menampilkan tabel Pegawai Setempat
    return render_template('pegawai/list.html')

# ==============================================
# RUN APPLICATION
# ==============================================

if __name__ == '__main__':
    app.run(debug=True)