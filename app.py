import os
import sqlite3
import zipfile
import io
import pandas as pd
import qrcode
from datetime import datetime
from flask import Flask, request, jsonify, send_file, render_template_string, redirect, render_template, flash
from werkzeug.utils import secure_filename

app = Flask(__name__)
# Secret key is required to use Flask's 'flash' message system for the state import alerts
app.secret_key = 'dmf_raigarh_secure_key' 

# Set up required folders
UPLOAD_FOLDER = 'project_documents'
QR_FOLDER = os.path.join('static', 'qrcodes')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(QR_FOLDER, exist_ok=True)

ALLOWED_EXTENSIONS = {'pdf', 'jpg', 'jpeg'}

# --- COMMON HEADER HTML TEMPLATE ---
HEADER_HTML = '''
<div style="background-color: #002147; border-bottom: 4px solid #FF9933; color: white; padding: 15px 0; margin-bottom: 30px; text-align: center;">
    <h3 style="margin: 0; font-weight: 600;">District Mineral Foundation (DMF)</h3>
    <p style="margin: 5px 0 0 0; font-weight: 300; font-size: 0.9rem;">District Administration, Raigarh, Government of Chhattisgarh</p>
</div>
'''

# --- DATA SOURCING: EXCEL REGISTER ---
try:
    xls = pd.ExcelFile('DMF Raigarh - Code Register.xlsx')
    
    # 1. Parse Block Codes
    df_blocks = pd.read_excel(xls, 'Block Codes')
    df_blocks.columns = df_blocks.iloc[2]
    df_blocks = df_blocks[3:].dropna(subset=['Block Code'])
    df_blocks['Block Code'] = df_blocks['Block Code'].astype(str).str.strip()
    block_dict = df_blocks.set_index('Block Code')['Development Block'].to_dict()

    # 2. Parse Village Codes and create mapping for dropdown logic
    df_villages = pd.read_excel(xls, 'Master Village Index')
    df_villages.columns = df_villages.iloc[2]
    df_villages = df_villages[3:].dropna(subset=['VIC Code'])
    
    village_dict = {}
    village_map = {}
    for idx, row in df_villages.iterrows():
        v_code = str(row['VIC Code']).strip()
        b_code = str(row['Block Code']).strip()
        if v_code == 'nan' or b_code == 'nan': continue
        
        village_dict[v_code] = row.to_dict()
        
        if b_code not in village_map: village_map[b_code] = []
        village_map[b_code].append({'code': v_code, 'label': f"{row['Village Name']} - {row['Gram Panchayat']} GP"})

    # 3. Parse Sector Codes
    df_sectors = pd.read_excel(xls, 'Sector Codes')
    df_sectors.columns = df_sectors.iloc[3]
    df_sectors = df_sectors[4:].dropna(subset=['Code'])
    df_sectors['Code'] = df_sectors['Code'].astype(str).str.strip()
    df_sectors['Priority Tier'] = df_sectors['Priority Tier'].ffill().str.replace('\n', ' ')
    sector_dict = df_sectors.set_index('Code').to_dict('index')

except Exception as e:
    print(f"Error loading Excel data: {e}")
    block_dict, village_dict, sector_dict, village_map = {}, {}, {}, {}


# --- DATABASE SETUP ---
def get_db():
    conn = sqlite3.connect('dmf_trial.db')
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    with get_db() as conn:
        conn.execute('''
            CREATE TABLE IF NOT EXISTS projects (
                id INTEGER PRIMARY KEY AUTOINCREMENT, district TEXT, block TEXT, village TEXT, sector TEXT,
                fy TEXT, seq INTEGER, status TEXT, work_name TEXT, department TEXT, objective TEXT, 
                justification TEXT, proposed_by TEXT, created_at TEXT, updated_at TEXT,
                wo_number TEXT, as_number TEXT
            )
        ''')
init_db()

def validate_file(file):
    filename = file.filename
    if '.' not in filename: return False
    ext = filename.rsplit('.', 1)[1].lower()
    if ext not in ALLOWED_EXTENSIONS: return False
    file.seek(0, os.SEEK_END)
    size = file.tell()
    file.seek(0)
    if ext == 'pdf' and size > 10 * 1024 * 1024: return False
    if ext in ['jpg', 'jpeg'] and size > 500 * 1024: return False
    return True


# --- ROUTES ---

@app.route('/')
def submission_form():
    return render_template('form.html', blocks=block_dict, village_map=village_map, sectors=sector_dict)


@app.route('/submit_project', methods=['POST'])
def submit_project():
    fields = ['block', 'village', 'sector', 'fy', 'work_name', 'department', 'objective', 'justification', 'proposed_by', 'wo_number']
    data = {f: request.form.get(f) for f in fields}
    
    has_as = request.form.get('has_as')
    as_number = request.form.get('as_number') if has_as == 'yes' else None
    
    if not all(data.values()): return jsonify({'error': 'Missing form data'}), 400

    # Duplicate Checks (WO & AS)
    with get_db() as conn:
        cur = conn.cursor()
        exists_wo = cur.execute('SELECT id FROM projects WHERE wo_number = ?', (data['wo_number'],)).fetchone()
        if exists_wo:
            return jsonify({'error': 'Duplicate Entry: This Work Order (WO) number already exists in the system.'}), 400
        
        if as_number:
            exists_as = cur.execute('SELECT id FROM projects WHERE as_number = ?', (as_number,)).fetchone()
            if exists_as:
                return jsonify({'error': 'Duplicate Entry: This Administrative Sanction (AS) number already exists in the system.'}), 400

    # File validation check
    doc_types = ['UCC', 'GEO', 'ADS', 'TES', 'PRP']
    for doc in doc_types:
        if doc not in request.files or not validate_file(request.files[doc]):
             return jsonify({'error': f'File {doc} is missing or violates size/type rules.'}), 400

    current_time = datetime.now().strftime("%Y-%m-%d %I:%M %p")

    # Generate Sequence and Save to DB
    with get_db() as conn:
        cur = conn.cursor()
        cur.execute('SELECT MAX(seq) FROM projects WHERE block=? AND village=? AND sector=? AND fy=?', 
                    (data['block'], data['village'], data['sector'], data['fy']))
        max_seq = cur.fetchone()[0]
        seq = 1 if max_seq is None else max_seq + 1
        seq_str = f"{seq:03d}"

        cur.execute('''INSERT INTO projects 
            (district, block, village, sector, fy, seq, status, work_name, department, objective, justification, proposed_by, created_at, updated_at, wo_number, as_number) 
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''', 
            ("Raigarh", data['block'], data['village'], data['sector'], data['fy'], seq, 'S', 
             data['work_name'], data['department'], data['objective'], data['justification'], data['proposed_by'], current_time, 'Never', data['wo_number'], as_number))
        project_id = cur.lastrowid

    # Save physical files with immutable base ID
    base_filename = f"{data['block']}_{data['village']}_{data['sector']}_{data['fy']}_{seq_str}"
    for doc in doc_types:
        file = request.files[doc]
        ext = file.filename.rsplit('.', 1)[1].lower()
        file.save(os.path.join(UPLOAD_FOLDER, f"{base_filename}_{doc}.{ext}"))

    # Generate QR Code for Social Audit
    public_url = f"{request.host_url}public/view/{base_filename}"
    qr = qrcode.QRCode(box_size=10, border=2)
    qr.add_data(public_url)
    qr.make(fit=True)
    qr_img = qr.make_image(fill_color="black", back_color="white")
    
    qr_filename = f"{base_filename}.png"
    qr_path = os.path.join(QR_FOLDER, qr_filename)
    qr_img.save(qr_path)

    return render_template('success.html', file_base=base_filename, time=current_time, project_id=project_id)


@app.route('/view/<int:project_id>')
def view_project(project_id):
    conn = get_db()
    project = conn.execute('SELECT * FROM projects WHERE id = ?', (project_id,)).fetchone()
    if not project: return "Project not found.", 404

    base_id = f"{project['block']}_{project['village']}_{project['sector']}_{project['fy']}_{project['seq']:03d}"
    
    html = f'''
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <title>View Application</title>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
        <style>body {{ background-color: #f4f6f9; }} .label {{ font-weight: 600; color: #495057; font-size: 0.9rem; }}</style>
    </head>
    <body>
        {HEADER_HTML}
        <div class="container" style="max-width: 800px;">
            <div class="card shadow-sm">
                <div class="card-header bg-white border-bottom border-3 border-primary py-3">
                    <h5 class="mb-0 text-primary">Application Details: {base_id}</h5>
                </div>
                <div class="card-body">
                    <div class="row g-3 mb-3">
                        <div class="col-md-12"><div class="p-3 bg-light rounded"><span class="label">Work Name</span><br>{{{{ p.work_name }}}}</div></div>
                        <div class="col-md-6"><div><span class="label">Work Order (WO)</span><br><span class="badge bg-secondary">{{{{ p.wo_number }}}}</span></div></div>
                        <div class="col-md-6"><div><span class="label">Administrative Sanction (AS)</span><br><span class="badge bg-info text-dark">{{{{ p.as_number or 'Pending' }}}}</span></div></div>
                        <div class="col-md-6"><div><span class="label">Department</span><br>{{{{ p.department }}}}</div></div>
                        <div class="col-md-6"><div><span class="label">Proposed By</span><br>{{{{ p.proposed_by }}}}</div></div>
                        <hr>
                        <div class="col-md-4"><div><span class="label">District</span><br>{{{{ p.district }}}}</div></div>
                        <div class="col-md-4"><div><span class="label">Block Code</span><br>{{{{ p.block }}}}</div></div>
                        <div class="col-md-4"><div><span class="label">Village Code</span><br>{{{{ p.village }}}}</div></div>
                        <div class="col-md-6"><div><span class="label">Sector Code</span><br>{{{{ p.sector }}}}</div></div>
                        <div class="col-md-6"><div><span class="label">Financial Year</span><br>{{{{ p.fy }}}}</div></div>
                        <hr>
                        <div class="col-md-12"><div><span class="label">Objective</span><br>{{{{ p.objective }}}}</div></div>
                        <div class="col-md-12"><div><span class="label">Justification</span><br>{{{{ p.justification }}}}</div></div>
                    </div>
                </div>
                <div class="card-footer text-muted text-end small">Submitted On: {{{{ p.created_at }}}}</div>
            </div>
            <div class="text-center mt-4 mb-5">
                <a href="/" class="btn btn-primary">Return to Form</a>
            </div>
        </div>
    </body>
    </html>
    '''
    return render_template_string(html, p=project)


# --- DASHBOARD AUTO-UPDATE API ---
@app.route('/api/latest_project_id')
def latest_project_id():
    # A lightweight endpoint that returns the ID of the newest project in the database.
    # The dashboard Javascript silently polls this every 10 seconds.
    with get_db() as conn:
        res = conn.execute('SELECT MAX(id) FROM projects').fetchone()[0]
        return jsonify({'max_id': res or 0})


@app.route('/dashboard')
def dashboard():
    with get_db() as conn:
        projects = conn.execute('SELECT * FROM projects ORDER BY id DESC').fetchall()
    
    dashboard_data = []
    max_id = 0
    for p in projects:
        if p['id'] > max_id:
            max_id = p['id']
            
        b_code, v_code, s_code = p['block'], p['village'], p['sector']
        b_name = block_dict.get(b_code, b_code)
        v_info = village_dict.get(v_code, {})
        v_name = v_info.get('Village Name', v_code)
        gp_name = v_info.get('Gram Panchayat', 'Unknown GP')
        v_cat = v_info.get('Category', 'Unknown Category')
        s_info = sector_dict.get(s_code, {})
        s_name = s_info.get('Sector', s_code)
        s_pri = s_info.get('Priority Tier', 'Unknown Priority')
        
        summary = (f"Work: {p['work_name']} ({s_name} - {s_pri}). "
                   f"Location: {v_name}, {gp_name} GP ({v_cat}) in {b_name}, {p['district']}. "
                   f"Dept: {p['department']}. Proposed by: {p['proposed_by']}.")
                   
        file_base = f"{b_code}_{v_code}_{s_code}_{p['fy']}_{p['seq']:03d}"
        
        dashboard_data.append({
            'id': p['id'], 'file_base': file_base, 'b_name': b_name, 'v_name': v_name, 
            'gp_name': gp_name, 's_name': s_name, 's_code': s_code, 's_pri': s_pri, 
            'v_cat': v_cat, 'status': p['status'], 'summary': summary,
            'created_at': p['created_at'], 'updated_at': p['updated_at'],
            'wo_number': p['wo_number'], 'as_number': p['as_number']
        })

    from flask import get_flashed_messages
    messages = get_flashed_messages(with_categories=True)
    msg_html = ""
    for category, message in messages:
        alert_type = 'success' if category == 'success' else 'danger'
        msg_html += f'<div class="alert alert-{alert_type} alert-dismissible fade show">{message}<button type="button" class="btn-close" data-bs-dismiss="alert"></button></div>'

    html = f'''
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <title>DMF Control Dashboard</title>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
        <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/1.13.6/css/dataTables.bootstrap5.min.css">
        <style>
            body {{ background-color: #f4f6f9; }}
            .table-container {{ background: white; padding: 25px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); margin: 0 20px; }}
            table.dataTable tbody tr {{ font-size: 0.85rem; }}
            .timestamp-box {{ font-size: 0.75rem; color: #6c757d; margin-top: 8px; border-top: 1px dashed #dee2e6; padding-top: 5px; }}
        </style>
    </head>
    <body>
        {HEADER_HTML}
        
        <!-- NEW: Hidden update banner triggered by Javascript polling -->
        <div id="updateBanner" class="alert alert-warning text-center mx-auto shadow-sm" style="display: none; max-width: 600px; position: sticky; top: 10px; z-index: 1000;">
            <strong>Notice:</strong> New project submissions detected.
            <button class="btn btn-sm btn-primary ms-3 fw-bold" onclick="location.reload()">Refresh Dashboard Now</button>
        </div>

        <div class="table-container mt-3">
            {msg_html}
            
            <div class="row align-items-center mb-4 pb-3 border-bottom">
                <div class="col-md-6">
                    <h4 class="text-primary mb-0">Project Control Dashboard</h4>
                </div>
                <div class="col-md-6 text-end">
                    <form action="/import_state_data" method="POST" enctype="multipart/form-data" class="d-inline-flex align-items-center bg-light p-2 rounded border">
                        <span class="me-2 fw-bold text-muted small">Bulk State Sync:</span>
                        <input type="file" name="state_file" accept=".csv, .xlsx" class="form-control form-control-sm w-auto me-2" required>
                        <button type="submit" class="btn btn-sm btn-dark">Import</button>
                    </form>
                    <div class="form-text text-muted small mt-1 text-end">Upload CSV/Excel with 'wo_number' & 'status' columns.</div>
                </div>
            </div>

            <table id="projectsTable" class="table table-bordered table-hover w-100">
                <thead class="table-light">
                    <tr>
                        <th>Base ID / Ref Numbers</th>
                        <th>Location</th>
                        <th>Project Type</th>
                        <th>Village</th>
                        <th>Priority</th>
                        <th width="25%">Summary & Trace</th>
                        <th>Status</th>
                        <th width="10%">Actions</th>
                    </tr>
                </thead>
                <tbody>
                    {{% for row in data %}}
                    <tr>
                        <td>
                            <span class="fw-bold">{{{{ row.file_base }}}}</span><br>
                            <small class="text-primary">WO: {{{{ row.wo_number }}}}</small><br>
                            <small class="text-info">AS: {{{{ row.as_number or 'N/A' }}}}</small>
                        </td>
                        <td>{{{{ row.b_name }}}}<br><small class="text-muted">{{{{ row.gp_name }}}} GP</small></td>
                        <td>{{{{ row.s_name }}}}<br><small class="text-muted">({{{{ row.s_code }}}})</small></td>
                        <td>{{{{ row.v_name }}}}</td>
                        <td><span class="badge bg-secondary">{{{{ row.s_pri }}}}</span></td>
                        <td>
                            {{{{ row.summary }}}}
                            <div class="timestamp-box">
                                Sub: {{{{ row.created_at }}}} <br> Edt: {{{{ row.updated_at }}}}
                            </div>
                        </td>
                        <td>
                            <form action="/update_status/{{{{ row.id }}}}" method="POST" class="m-0">
                                <select name="status" class="form-select form-select-sm" onchange="this.form.submit()">
                                    <option value="S" {{% if row.status == 'S' %}}selected{{% endif %}}>Submitted</option>
                                    <option value="A" {{% if row.status == 'A' %}}selected{{% endif %}}>Approved</option>
                                    <option value="T" {{% if row.status == 'T' %}}selected{{% endif %}}>Tendered</option>
                                    <option value="O" {{% if row.status == 'O' %}}selected{{% endif %}}>Ongoing</option>
                                    <option value="C" {{% if row.status == 'C' %}}selected{{% endif %}}>Completed</option>
                                    <option value="D" {{% if row.status == 'D' %}}selected{{% endif %}}>Dropped</option>
                                </select>
                            </form>
                        </td>
                        <td>
                            <div class="d-flex flex-column gap-1">
                                <div class="btn-group">
                                    <a href="/download_zip/{{{{ row.id }}}}" class="btn btn-sm btn-primary">ZIP</a>
                                    <!-- NEW: QR Code Direct Download Button -->
                                    <a href="/static/qrcodes/{{{{ row.file_base }}}}.png" download="{{{{ row.file_base }}}}_QR.png" class="btn btn-sm btn-info text-white fw-bold">QR</a>
                                </div>
                                <a href="/edit/{{{{ row.id }}}}" class="btn btn-sm btn-warning" onclick="return confirm('Edit base codes?');">Edit</a>
                                <form action="/delete/{{{{ row.id }}}}" method="POST" class="m-0" onsubmit="return confirm('Permanently delete project and files?');">
                                    <button type="submit" class="btn btn-sm btn-danger w-100">Del</button>
                                </form>
                            </div>
                        </td>
                    </tr>
                    {{% endfor %}}
                </tbody>
            </table>
        </div>
        <script src="https://code.jquery.com/jquery-3.7.0.min.js"></script>
        <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
        <script src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js"></script>
        <script src="https://cdn.datatables.net/1.13.6/js/dataTables.bootstrap5.min.js"></script>
        
        <script> 
            $(document).ready(function() {{ 
                $('#projectsTable').DataTable(); 
                
                // --- NEW LOGIC: Smart Auto-Update Polling ---
                // Stores the highest project ID currently loaded on the page
                let currentMaxId = {{{{ max_id }}}};
                
                // Pings the backend every 10 seconds silently
                setInterval(() => {{
                    fetch('/api/latest_project_id')
                        .then(response => response.json())
                        .then(data => {{
                            // If the backend has a higher ID, drop down the update banner
                            if (data.max_id > currentMaxId) {{
                                document.getElementById('updateBanner').style.display = 'block';
                            }}
                        }})
                        .catch(err => console.error("Polling error:", err));
                }}, 10000);
            }}); 
        </script>
    </body>
    </html>
    '''
    return render_template_string(html, data=dashboard_data, max_id=max_id)


@app.route('/import_state_data', methods=['POST'])
def import_state_data():
    file = request.files.get('state_file')
    if not file or file.filename == '':
        flash('No file selected.', 'error')
        return redirect('/dashboard')

    try:
        if file.filename.endswith('.csv'): df = pd.read_csv(file)
        else: df = pd.read_excel(file)

        if 'wo_number' not in df.columns or 'status' not in df.columns:
            flash("Upload failed: The file must contain exact columns named 'wo_number' and 'status'.", 'error')
            return redirect('/dashboard')

        update_count = 0
        edit_time = datetime.now().strftime("%Y-%m-%d %I:%M %p")
        valid_statuses = ['S', 'A', 'T', 'O', 'C', 'D']

        with get_db() as conn:
            cur = conn.cursor()
            for index, row in df.iterrows():
                wo = str(row['wo_number']).strip()
                new_status = str(row['status']).strip().upper()
                if wo != 'nan' and new_status in valid_statuses:
                    cur.execute('UPDATE projects SET status=?, updated_at=? WHERE wo_number=?', (new_status, edit_time, wo))
                    update_count += cur.rowcount
            conn.commit()

        flash(f'Success! {update_count} project statuses updated from State sync.', 'success')
    except Exception as e:
        flash(f"An error occurred while processing the file: {str(e)}", 'error')

    return redirect('/dashboard')


@app.route('/update_status/<int:project_id>', methods=['POST'])
def update_status(project_id):
    edit_time = datetime.now().strftime("%Y-%m-%d %I:%M %p")
    with get_db() as conn:
        conn.execute('UPDATE projects SET status = ?, updated_at = ? WHERE id = ?', (request.form.get('status'), edit_time, project_id))
        conn.commit()
    return redirect('/dashboard')


@app.route('/delete/<int:project_id>', methods=['POST'])
def delete_project(project_id):
    with get_db() as conn:
        project = conn.execute('SELECT * FROM projects WHERE id = ?', (project_id,)).fetchone()
        if project:
            base = f"{project['block']}_{project['village']}_{project['sector']}_{project['fy']}_{project['seq']:03d}"
            
            # Delete associated files (PDFs/Images)
            for f in os.listdir(UPLOAD_FOLDER):
                if f.startswith(base): os.remove(os.path.join(UPLOAD_FOLDER, f))
                
            # Delete associated QR Code
            qr_path = os.path.join(QR_FOLDER, f"{base}.png")
            if os.path.exists(qr_path): os.remove(qr_path)
            
            conn.execute('DELETE FROM projects WHERE id = ?', (project_id,))
            conn.commit()
    return redirect('/dashboard')


@app.route('/edit/<int:project_id>', methods=['GET', 'POST'])
def edit_project(project_id):
    conn = get_db()
    project = conn.execute('SELECT * FROM projects WHERE id = ?', (project_id,)).fetchone()
    if request.method == 'POST':
        new_block = request.form.get('block')
        new_village = request.form.get('village')
        new_sector = request.form.get('sector')
        new_fy = request.form.get('fy')
        edit_time = datetime.now().strftime("%Y-%m-%d %I:%M %p")
        
        old_base = f"{project['block']}_{project['village']}_{project['sector']}_{project['fy']}_{project['seq']:03d}"
        new_base = f"{new_block}_{new_village}_{new_sector}_{new_fy}_{project['seq']:03d}"
        
        if old_base != new_base:
            # Rename physical documents
            for f in os.listdir(UPLOAD_FOLDER):
                if f.startswith(old_base):
                    os.rename(os.path.join(UPLOAD_FOLDER, f), os.path.join(UPLOAD_FOLDER, f.replace(old_base, new_base, 1)))
            
            # Rename QR code
            old_qr = os.path.join(QR_FOLDER, f"{old_base}.png")
            new_qr = os.path.join(QR_FOLDER, f"{new_base}.png")
            if os.path.exists(old_qr): os.rename(old_qr, new_qr)
        
        conn.execute('UPDATE projects SET block=?, village=?, sector=?, fy=?, updated_at=? WHERE id=?', 
                     (new_block, new_village, new_sector, new_fy, edit_time, project_id))
        conn.commit()
        return redirect('/dashboard')

    html = f'''
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <title>Edit Project Metadata</title>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
        <style>body {{ background-color: #f4f6f9; }} label {{ font-weight: 500; font-size: 0.9rem; }}</style>
    </head>
    <body>
        {HEADER_HTML}
        <div class="container" style="max-width: 600px;">
            <div class="card shadow-sm border-warning">
                <div class="card-header bg-warning text-dark fw-bold">Admin Override: Edit Base Attributes</div>
                <div class="card-body">
                    <p class="small text-danger mb-4">Warning: Modifying these attributes will automatically recalculate the Base ID and rename the physical documents on the server.</p>
                    <form method="POST">
                        <div class="mb-3">
                            <label class="form-label">Block Code</label>
                            <input type="text" class="form-control" name="block" id="blockInput" list="blockOptions" value="{{{{ p.block }}}}" autocomplete="off" required>
                            <datalist id="blockOptions">{{% for code, name in blocks.items() %}}<option value="{{{{ code }}}}">{{{{ name }}}}</option>{{% endfor %}}</datalist>
                        </div>
                        <div class="mb-3">
                            <label class="form-label">Village Code</label>
                            <input type="text" class="form-control" name="village" id="villageInput" list="villageOptions" value="{{{{ p.village }}}}" autocomplete="off" required>
                            <datalist id="villageOptions"></datalist>
                        </div>
                        <div class="mb-3">
                            <label class="form-label">Sector Code</label>
                            <input type="text" class="form-control" name="sector" list="sectorOptions" value="{{{{ p.sector }}}}" autocomplete="off" required>
                            <datalist id="sectorOptions">{{% for code, info in sectors.items() %}}<option value="{{{{ code }}}}">{{{{ info['Sector'] }}}}</option>{{% endfor %}}</datalist>
                        </div>
                        <div class="mb-4">
                            <label class="form-label">Financial Year</label>
                            <input type="text" class="form-control" name="fy" value="{{{{ p.fy }}}}" required>
                        </div>
                        <div class="d-grid gap-2">
                            <button type="submit" class="btn btn-warning fw-bold">Save Administrative Changes</button>
                            <a href="/dashboard" class="btn btn-light border">Cancel</a>
                        </div>
                    </form>
                </div>
            </div>
        </div>
        <script>
            const villageMap = {{{{ village_map | tojson | safe }}}};
            const blockInput = document.getElementById('blockInput');
            const villageInput = document.getElementById('villageInput');
            const villageDatalist = document.getElementById('villageOptions');
            function populateVillages(blockCode) {{
                villageDatalist.innerHTML = ''; 
                if (villageMap[blockCode]) {{
                    villageMap[blockCode].forEach(function(v) {{
                        const opt = document.createElement('option');
                        opt.value = v.code; opt.innerText = v.label;
                        villageDatalist.appendChild(opt);
                    }});
                }}
            }}
            populateVillages(blockInput.value);
            blockInput.addEventListener('input', function() {{ populateVillages(this.value); villageInput.value = ''; }});
        </script>
    </body>
    </html>
    '''
    return render_template_string(html, p=project, blocks=block_dict, village_map=village_map, sectors=sector_dict)


@app.route('/download_zip/<int:project_id>', methods=['GET'])
def download_zip(project_id):
    with get_db() as conn: project = conn.execute('SELECT * FROM projects WHERE id = ?', (project_id,)).fetchone()
    
    disk_base = f"{project['block']}_{project['village']}_{project['sector']}_{project['fy']}_{project['seq']:03d}"
    zip_base = f"{project['block']}_{project['village']}_{project['sector']}_{project['fy']}_{project['status']}_{project['seq']:03d}"
    
    mem = io.BytesIO()
    with zipfile.ZipFile(mem, 'w') as zf:
        for f in os.listdir(UPLOAD_FOLDER):
            if f.startswith(disk_base): 
                zf.write(os.path.join(UPLOAD_FOLDER, f), arcname=f.replace(disk_base, zip_base, 1))
    mem.seek(0)
    return send_file(mem, download_name=f"{zip_base}.zip", as_attachment=True, mimetype='application/zip')


# --- SOCIAL AUDIT ROUTES ---

@app.route('/print_plaque/<int:project_id>')
def print_plaque(project_id):
    conn = get_db()
    project = conn.execute('SELECT * FROM projects WHERE id = ?', (project_id,)).fetchone()
    if not project: return "Project not found.", 404
    
    base_id = f"{project['block']}_{project['village']}_{project['sector']}_{project['fy']}_{project['seq']:03d}"
    qr_url = f"/static/qrcodes/{base_id}.png"
    
    return render_template('plaque.html', base_id=base_id, qr_url=qr_url)

@app.route('/public/view/<base_id>')
def public_view(base_id):
    parts = base_id.split('_')
    if len(parts) < 5: return "Invalid Project ID", 400
    
    block, village, sector, fy, seq_str = parts[0], parts[1], parts[2], parts[3], parts[4]
    
    conn = get_db()
    project = conn.execute('SELECT * FROM projects WHERE block=? AND village=? AND sector=? AND fy=? AND seq=?', 
                           (block, village, sector, fy, int(seq_str))).fetchone()
    
    if not project: return "Project not found.", 404

    documents = []
    doc_types = {'UCC': 'Utilization Certificate', 'GEO': 'Site Photograph', 
                 'ADS': 'Admin Sanction', 'TES': 'Technical Sanction', 'PRP': 'Proposal'}
                 
    for filename in os.listdir(UPLOAD_FOLDER):
        if filename.startswith(base_id):
            doc_code = filename.replace(f"{base_id}_", "").split('.')[0].upper()
            if doc_code in doc_types:
                documents.append({
                    'type': doc_types[doc_code],
                    'filename': filename
                })

    return render_template('public_view.html', p=project, base_id=base_id, documents=documents)

@app.route('/public/download/<filename>')
def public_download(filename):
    safe_filename = secure_filename(filename)
    return send_file(os.path.join(UPLOAD_FOLDER, safe_filename))

if __name__ == '__main__':
    app.run(debug=True, port=5000)