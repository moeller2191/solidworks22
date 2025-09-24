

# --- Excel Data Processing Function ---
def update_job_orders_from_excel():
    print("Starting Excel file update...")
    excel_file = 'JobOrders.xlsx'
    # Check if the Excel file exists
    if not os.path.exists(excel_file):
        print(f"Error: Excel file '{excel_file}' not found.")
        return
    try:
        df = pd.read_excel(excel_file, sheet_name='Sheet1', engine='openpyxl')
        df = df.where(pd.notnull(df), None)
        # Delete all old records
        num_deleted = db.session.query(JobOrder).delete()
        db.session.commit()
        print(f"Deleted {num_deleted} old records from the database.")
        # Load new records
        for index, row in df.iterrows():
            job = JobOrder(
                fjobno=str(row.get('fjobno', '')),
                fpartrev=str(row.get('fpartrev', '')) if 'fpartrev' in row else None,
                fquantity=int(row['fquantity']) if 'fquantity' in row and pd.notnull(row.get('fquantity')) and isinstance(row.get('fquantity'), (int, float)) else 0,
                fstatus=str(row.get('fstatus', '')) if 'fstatus' in row else None,
                fdesc=str(row.get('fdesc', '')) if 'fdesc' in row else None,
                fcudrev=str(row.get('fcudrev', '')) if 'fcudrev' in row else None,
                fdescmemo=str(row.get('fdescmemo', '')) if 'fdescmemo' in row else None,
                fpartnoOrginal=str(row.get('fpartnoOrginal', '')) if 'fpartnoOrginal' in row else None,
                find_rev=str(row.get('find_rev', '')) if 'find_rev' in row else None,
                find_rev2=str(row.get('find_rev2', '')) if 'find_rev2' in row else None,
                find_rev3=str(row.get('find_rev3', '')) if 'find_rev3' in row else None,
                find_rev4=str(row.get('find_rev4', '')) if 'find_rev4' in row else None,
                select=str(row.get('select', '')) if 'select' in row else None,
                kirby_check=str(row.get('kirby_check', '')) if 'kirby_check' in row else None,
                kirby_p_hash=str(row.get('kirby_p_hash', '')) if 'kirby_p_hash' in row else None,
                final_rev=str(row.get('final_rev', '')) if 'final_rev' in row else None,
                fpartno=str(row.get('fpartno', '')) if 'fpartno' in row else None,
                gaston=str(row.get('gaston', '')) if 'gaston' in row else None,
                final_rev_review=str(row.get('final_rev_review', '')) if 'final_rev_review' in row else None,
                combined_gaston=str(row.get('combined_gaston', '')) if 'combined_gaston' in row else None,
                combined_carrier=str(row.get('combined_carrier', '')) if 'combined_carrier' in row else None,
                combined_rev_wkyr=str(row.get('combined_rev_wkyr', '')) if 'combined_rev_wkyr' in row else None
            )
            db.session.add(job)
        db.session.commit()
        print(f"Inserted {len(df)} new records into the database.")
    except Exception as e:
        print(f"Error updating job orders from Excel: {e}")




# --- Error Handlers for JSON API ---
from werkzeug.exceptions import HTTPException

import os
import requests
import pandas as pd
from flask import Flask, request, jsonify, render_template, send_from_directory
from flask_sqlalchemy import SQLAlchemy
from openpyxl import load_workbook
from flask_apscheduler import APScheduler
import logging
from werkzeug.utils import secure_filename

# Allowed extensions for uploads
ALLOWED_EXTENSIONS = {'xlsx', 'dymo'}


# Initialize Flask App
app = Flask(__name__, static_folder='static', template_folder='templates')
logging.basicConfig(level=logging.INFO)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///dmg.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)
scheduler = APScheduler()

@app.errorhandler(Exception)
def handle_exception(e):
    # If the request is for an API endpoint, return JSON
    if request.path.startswith('/upload_') or request.path.startswith('/api/'):
        code = 500
        if isinstance(e, HTTPException):
            code = e.code
        return jsonify({'error': str(e)}), code
    # Otherwise, use default error handling
    raise e

# --- Utility Functions ---
def allowed_file(filename, allowed_exts):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed_exts
from werkzeug.exceptions import HTTPException

@app.errorhandler(Exception)
def handle_exception(e):
    # If the request is for an API endpoint, return JSON
    if request.path.startswith('/upload_') or request.path.startswith('/api/'):
        code = 500
        if isinstance(e, HTTPException):
            code = e.code
        return jsonify({'error': str(e)}), code
    # Otherwise, use default error handling
    raise e

            # --- Utility Functions ---
def allowed_file(filename, allowed_exts):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed_exts

# Initialize Flask App
app = Flask(__name__, static_folder='static', template_folder='templates')
logging.basicConfig(level=logging.INFO)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///dmg.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)
scheduler = APScheduler()


# --- Models ---
class PredefinedLocation(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(50), unique=True, nullable=False)

# --- Label Print Log Model ---
class LabelPrintLog(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    job_number = db.Column(db.String(50), nullable=False)
    quantity_printed = db.Column(db.Integer, nullable=False)
    timestamp = db.Column(db.DateTime, default=db.func.now())

# --- Label Print Status Endpoint ---
@app.route('/print_label_status', methods=['GET'])
def print_label_status():
    job_number = request.args.get('job_number')
    if not job_number:
        return jsonify({'error': 'Job number required'}), 400
    job = JobOrder.query.filter_by(fjobno=job_number).first()
    if not job:
        return jsonify({'error': 'Job not found'}), 404
    allowed = job.fquantity
    total_printed = db.session.query(db.func.sum(LabelPrintLog.quantity_printed)).filter_by(job_number=job_number).scalar() or 0
    return jsonify({'job_number': job_number, 'allowed': allowed, 'printed': total_printed})

# --- Label Print Endpoint with Supervisor Override ---
@app.route('/print_label', methods=['POST'])
def print_label():
    data = request.get_json()
    job_number = data.get('job_number')
    quantity = int(data.get('quantity', 1))
    if not job_number or quantity < 1:
        return jsonify({'error': 'Job number and valid quantity required'}), 400

    job = JobOrder.query.filter_by(fjobno=job_number).first()
    if not job:
        return jsonify({'error': 'Job not found'}), 404

    # Calculate total printed so far
    total_printed = db.session.query(db.func.sum(LabelPrintLog.quantity_printed)).filter_by(job_number=job_number).scalar() or 0
    allowed = job.fquantity
    override = data.get('override', False)
    password = data.get('password', '')
    if total_printed + quantity > allowed:
        if not (override and password == 'password123'):
            return jsonify({'error': f'Cannot print {quantity} labels. Already printed {total_printed} of {allowed}.'}), 400
    # Log the print event
    log = LabelPrintLog(job_number=job_number, quantity_printed=quantity)
    db.session.add(log)
    db.session.commit()
    # ...existing code to trigger actual label printing...
    return jsonify({'message': f'{quantity} labels printed for job {job_number}. Total printed: {total_printed + quantity} of {allowed}.'})


# --- Send Teams Message Endpoint for Main Screen ---
@app.route('/send_message', methods=['POST'])
def send_message():
    data = request.get_json()
    job_number = data.get('job_number')
    if not job_number:
        return jsonify({'error': 'Job number is required'}), 400

    hardware = HardwareLocation.query.filter_by(job_number=job_number).first()
    location = hardware.location if hardware else 'Unknown'

    # Use a default webhook name for hardware notifications
    webhook_name = 'hardware'
    webhook = TeamsWebhook.query.filter_by(webhook_name=webhook_name).first()
    if not webhook:
        return jsonify({'error': f'Webhook for hardware notifications not found'}), 404

    message = f"Assembly job {job_number} is not at its expected location. Last known location: {location}."
    try:
        headers = {'Content-Type': 'application/json'}
        payload = {'text': message}
        response = requests.post(webhook.webhook_url, json=payload, headers=headers)
        response.raise_for_status()
        return jsonify({'message': 'Teams notification sent successfully'})
    except requests.exceptions.RequestException as e:
        return jsonify({'error': f'Failed to send Teams notification: {e}'}), 500

# --- Lookup Hardware Location Endpoint ---
@app.route('/lookup/<string:job_number>', methods=['GET'])
def lookup_job_location(job_number):
    hardware = HardwareLocation.query.filter_by(job_number=job_number).first()
    if hardware:
        return jsonify({'location': hardware.location})
    return jsonify({'error': 'Location not found'}), 404

# --- Log Location Endpoint for Main Screen ---
@app.route('/log_location', methods=['POST'])
def log_location():
    data = request.get_json()
    job_number = data.get('job_number')
    location_name = data.get('location')

    if not job_number or not location_name:
        return jsonify({'error': 'Job number and location name are required'}), 400

    try:
        # Check if the location is predefined
        predefined_location = PredefinedLocation.query.filter_by(name=location_name).first()
        if not predefined_location:
            return jsonify({'error': 'Invalid location name'}), 400
        hardware = HardwareLocation.query.filter_by(job_number=job_number).first()
        if hardware:
            hardware.location = location_name
        else:
            hardware = HardwareLocation(job_number=job_number, location=location_name)
            db.session.add(hardware)
        db.session.commit()
        return jsonify({'message': f'Location for job {job_number} logged as {location_name}.'})
    except Exception as e:
        db.session.rollback()
        return jsonify({'error': str(e)}), 500

# --- Predefined Location API ---
@app.route('/get_predefined_locations', methods=['GET'])
def get_predefined_locations():
    locations = PredefinedLocation.query.order_by(PredefinedLocation.name).all()
    return jsonify([loc.name for loc in locations])

@app.route('/add_predefined_location', methods=['POST'])
def add_predefined_location():
    data = request.get_json()
    name = data.get('name', '').strip()
    if not name:
        return jsonify({'error': 'Location name required'}), 400
    if PredefinedLocation.query.filter_by(name=name).first():
        return jsonify({'error': 'Location already exists'}), 400
    loc = PredefinedLocation(name=name)
    db.session.add(loc)
    db.session.commit()
    return jsonify({'message': f'Location "{name}" added successfully.'})

@app.route('/delete_predefined_location', methods=['POST'])
def delete_predefined_location():
    data = request.get_json()
    name = data.get('name', '').strip()
    loc = PredefinedLocation.query.filter_by(name=name).first()
    if not loc:
        return jsonify({'error': 'Location not found'}), 404
    db.session.delete(loc)
    db.session.commit()
    return jsonify({'message': f'Location "{name}" deleted successfully.'})

class JobOrder(db.Model):
    __tablename__ = 'job_orders'
    fjobno = db.Column(db.String(50), primary_key=True)
    fpartrev = db.Column(db.String(50))
    fquantity = db.Column(db.Integer)
    fstatus = db.Column(db.String(50))
    fdesc = db.Column(db.String(255))
    fcudrev = db.Column(db.String(50))
    fdescmemo = db.Column(db.String(255))
    fpartnoOrginal = db.Column(db.String(50))
    find_rev = db.Column(db.String(50))
    find_rev2 = db.Column(db.String(50))
    find_rev3 = db.Column(db.String(50))
    find_rev4 = db.Column(db.String(50))
    select = db.Column(db.String(50))
    kirby_check = db.Column(db.String(50))
    kirby_p_hash = db.Column(db.String(50))
    final_rev = db.Column(db.String(50))
    fpartno = db.Column(db.String(50))
    gaston = db.Column(db.String(50))
    final_rev_review = db.Column(db.String(50))
    combined_gaston = db.Column(db.String(50))
    combined_carrier = db.Column(db.String(50))
    combined_rev_wkyr = db.Column(db.String(50))


    # --- Upload Endpoint for Job Order Label Template ---
    @app.route('/upload_label_template', methods=['POST'])
    def upload_label_template():
        return _upload_label_template('job_order')

    # --- Upload Endpoint for Rework Label Template ---
    @app.route('/upload_rework_label_template', methods=['POST'])
    def upload_rework_label_template():
        return _upload_label_template('rework')

    def _upload_label_template(template_name):
        if 'label_file' not in request.files:
            return jsonify({'error': 'No file part'}), 400
        file = request.files['label_file']
        if file.filename == '':
            return jsonify({'error': 'No selected file'}), 400
        if not allowed_file(file.filename, {'dymo'}):
            return jsonify({'error': 'Invalid file type'}), 400
        xml = file.read().decode('utf-8')
        template = LabelTemplate.query.filter_by(name=template_name).first()
        if template:
            template.xml = xml
        else:
            template = LabelTemplate(name=template_name, xml=xml)
            db.session.add(template)
        db.session.commit()
        return jsonify({'message': f'{template_name.replace("_", " ").title()} label template uploaded successfully.'})

    # --- Fetch Endpoint for Job Order Label Template ---
    @app.route('/api/get_label_format', methods=['GET'])
    def get_label_format():
        return _get_label_template('job_order')

    # --- Fetch Endpoint for Rework Label Template ---
    @app.route('/api/get_rework_label_format', methods=['GET'])
    def get_rework_label_format():
        return _get_label_template('rework')

    def _get_label_template(template_name):
        template = LabelTemplate.query.filter_by(name=template_name).first()
        if template:
            return template.xml, 200, {'Content-Type': 'text/xml'}
        else:
            return jsonify({'error': f'{template_name.replace("_", " ").title()} label template not found.'}), 404

    def as_dict(self):
        return {c.name: getattr(self, c.name) for c in self.__table__.columns}

class ExceptionString(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    exception_text = db.Column(db.String(50), unique=True, nullable=False)

class TeamsWebhook(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    webhook_name = db.Column(db.String(50), unique=True, nullable=False)
    webhook_url = db.Column(db.String(255), nullable=False)

class HardwareLocation(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    job_number = db.Column(db.String(20), unique=True, nullable=False)
    location = db.Column(db.String(50), nullable=False)
    timestamp = db.Column(db.DateTime, default=db.func.now())

# --- Label Template Model ---

# --- Label Template Model ---
class LabelTemplate(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(50), unique=True, nullable=False)
    xml = db.Column(db.Text, nullable=False)

# --- Gaston Label Template Model ---
class GastonLabelTemplate(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(50), unique=True, nullable=False, default='gaston')
    xml = db.Column(db.Text, nullable=False)

# --- Gaston Label Template Upload Endpoint ---
@app.route('/upload_gaston_label_template', methods=['POST'])
def upload_gaston_label_template():
    if 'label_file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    file = request.files['label_file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    if not allowed_file(file.filename, {'dymo'}):
        return jsonify({'error': 'Invalid file type. Only .dymo allowed.'}), 400
    xml = file.read().decode('utf-8')
    template = GastonLabelTemplate.query.filter_by(name='gaston').first()
    if template:
        template.xml = xml
    else:
        template = GastonLabelTemplate(name='gaston', xml=xml)
        db.session.add(template)
    db.session.commit()
    return jsonify({'message': 'Gaston label template uploaded successfully.'})

# --- Gaston Label Template Fetch Endpoint ---
@app.route('/api/get_gaston_label_format', methods=['GET'])
def get_gaston_label_format():
    template = GastonLabelTemplate.query.filter_by(name='gaston').first()
    if template:
        return template.xml, 200, {'Content-Type': 'text/xml'}
    else:
        return jsonify({'error': 'Gaston label template not found.'}), 404

# --- Label Template Helper Functions ---
def _upload_label_template(template_name):
    if 'label_file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    file = request.files['label_file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    if not allowed_file(file.filename, {'dymo'}):
        return jsonify({'error': 'Invalid file type'}), 400
    xml = file.read().decode('utf-8')
    template = LabelTemplate.query.filter_by(name=template_name).first()
    if template:
        template.xml = xml
    else:
        template = LabelTemplate(name=template_name, xml=xml)
        db.session.add(template)
    db.session.commit()
    return jsonify({'message': f'{template_name.replace("_", " ").title()} label template uploaded successfully.'})

def _get_label_template(template_name):
    template = LabelTemplate.query.filter_by(name=template_name).first()
    if template:
        return template.xml, 200, {'Content-Type': 'text/xml'}
    else:
        return jsonify({'error': f'{template_name.replace("_", " ").title()} label template not found.'}), 404

class InspectionCode(db.Model):
    id = db.Column(db.Integer, primary_key=True)

    code = db.Column(db.String(10), unique=True, nullable=False)
    description = db.Column(db.String(255), nullable=False)

# --- API Endpoint: Get Single Job Order by Job Number ---
@app.route('/api/job_order/<string:job_no>', methods=['GET'])
def get_job_order(job_no):
    job = JobOrder.query.filter_by(fjobno=job_no).first()
    if not job:
        return jsonify({'error': 'Job number not found.'}), 404
    return jsonify({c.name: getattr(job, c.name) for c in job.__table__.columns})

# --- Home Page Route (after all models, before utility functions) ---
@app.route('/')
def home():
    return render_template('Home.html')

# --- UI Page Routes for Home Buttons ---
@app.route('/dymo_rework_labels')
def dymo_rework_labels():
    return render_template('DYMO Rework Labels.html')

@app.route('/dymo_job_order')
def dymo_job_order():
    return render_template('DYMO Label Printer.html')

@app.route('/gaston_label_print')
def gaston_label_print():
    return render_template('GastonLabelPrint.html')

@app.route('/locate_my_hardware')
def locate_my_hardware():
    return render_template('LocateMyHardware.html')


# --- Utility Functions ---
def allowed_file(filename, allowed_exts):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed_exts



# --- File Upload Endpoints ---
from flask import current_app

# --- Inspection Codes API Endpoint ---
@app.route('/api/inspection_codes', methods=['GET'])
def api_inspection_codes():
    try:
        codes = InspectionCode.query.all()
        return jsonify([
            {'code': c.code, 'description': c.description}
            for c in codes
        ])
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/upload_inspcode', methods=['POST'])
def upload_inspcode():
    import pandas as pd
    if 'file' not in request.files:
        return jsonify({'error': 'No file part in request'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    if not allowed_file(file.filename, {'xlsx'}):
        return jsonify({'error': 'Invalid file type. Only .xlsx allowed.'}), 400
    try:
        save_path = os.path.join(current_app.root_path, 'inspectioncode.xlsx')
        file.save(save_path)
        # Parse the Excel and update InspectionCode table
        df = pd.read_excel(save_path, engine='openpyxl')
        # Expect columns: code, description (case-insensitive)
        df.columns = [c.lower() for c in df.columns]
        if 'code' not in df.columns or 'description' not in df.columns:
            return jsonify({'error': 'Excel must have columns: code, description'}), 400
        # Remove all old records
        InspectionCode.query.delete()
        db.session.commit()
        # Add new records
        for _, row in df.iterrows():
            code = str(row['code']).strip()
            desc = str(row['description']).strip()
            if code and desc:
                db.session.add(InspectionCode(code=code, description=desc))
        db.session.commit()
        return jsonify({'message': 'inspectioncode.xlsx uploaded and InspectionCode table updated.'})
    except Exception as e:
        db.session.rollback()
        return jsonify({'error': f'Failed to process file: {str(e)}'}), 500

@app.route('/upload_reworklabel', methods=['POST'])
def upload_reworklabel():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part in request'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    if not allowed_file(file.filename, {'dymo'}):
        return jsonify({'error': 'Invalid file type. Only .dymo allowed.'}), 400
    try:
        save_path = os.path.join(current_app.root_path, 'static', 'ReworkLabel.dymo')
        file.save(save_path)
        return jsonify({'message': 'ReworkLabel.dymo uploaded and replaced successfully.'})
    except Exception as e:
        return jsonify({'error': str(e)}), 500



# --- API Endpoints and Routes ---


@app.route('/api/exceptions', methods=['GET'])
def get_exceptions():
    """Returns a list of all exception strings."""
    try:
        exceptions = ExceptionString.query.all()
        return jsonify([{'id': e.id, 'exception_text': e.exception_text} for e in exceptions])
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/exceptions', methods=['POST'])
def add_exception():
    """Adds a new exception string to the database."""
    data = request.get_json()
    exception_text = data.get('exception_text')
    if not exception_text:
        return jsonify({'error': 'Exception text required'}), 400
    
    existing_exception = ExceptionString.query.filter_by(exception_text=exception_text).first()
    if existing_exception:
        return jsonify({'error': 'Exception string already exists'}), 409
        
    try:
        new_exception = ExceptionString(exception_text=exception_text)
        db.session.add(new_exception)
        db.session.commit()
        return jsonify({'message': 'Exception added successfully'}), 201
    except Exception as e:
        db.session.rollback()
        return jsonify({'error': str(e)}), 500

@app.route('/api/exceptions/<int:exception_id>', methods=['DELETE'])
def delete_exception(exception_id):
    """Deletes an exception string by ID."""
    try:
        exception = ExceptionString.query.get_or_404(exception_id)
        db.session.delete(exception)
        db.session.commit()
        return jsonify({'message': 'Exception deleted successfully'})
    except Exception as e:
        db.session.rollback()
        return jsonify({'error': str(e)}), 500
    
# --- Teams Webhook Routes ---

@app.route('/set_webhook', methods=['POST'])
def set_webhook():
    """Saves or updates a Teams webhook URL."""
    data = request.get_json()
    webhook_name = data.get('webhook_name')
    webhook_url = data.get('webhook_url')

    if not webhook_name or not webhook_url:
        return jsonify({'error': 'Webhook name and URL are required'}), 400

    try:
        webhook = TeamsWebhook.query.filter_by(webhook_name=webhook_name).first()
        if webhook:
            webhook.webhook_url = webhook_url
        else:
            webhook = TeamsWebhook(webhook_name=webhook_name, webhook_url=webhook_url)
            db.session.add(webhook)
        db.session.commit()
        return jsonify({'message': 'Webhook saved successfully'})
    except Exception as e:
        db.session.rollback()
        return jsonify({'error': str(e)}), 500

@app.route('/get_webhook/<string:webhook_name>', methods=['GET'])
def get_webhook(webhook_name):
    """Retrieves a Teams webhook URL by name."""
    try:
        webhook = TeamsWebhook.query.filter_by(webhook_name=webhook_name).first()
        if webhook:
            return jsonify({'url': webhook.webhook_url})
        return jsonify({'error': 'Webhook not found'}), 404
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/send_teams_notification', methods=['POST'])
def send_teams_notification():
    """Sends a notification to a Teams channel."""
    data = request.get_json()
    webhook_name = data.get('webhook_name')
    message = data.get('message')

    if not webhook_name or not message:
        return jsonify({'error': 'Webhook name and message are required'}), 400

    webhook = TeamsWebhook.query.filter_by(webhook_name=webhook_name).first()
    if not webhook:
        return jsonify({'error': f'Webhook with name "{webhook_name}" not found'}), 404

    try:
        headers = {'Content-Type': 'application/json'}
        payload = {'text': message}
        response = requests.post(webhook.webhook_url, json=payload, headers=headers)
        response.raise_for_status()
        return jsonify({'message': 'Notification sent successfully'})
    except requests.exceptions.RequestException as e:
        return jsonify({'error': f'Failed to send notification: {e}'}), 500
        
@app.route('/log-rework', methods=['POST'])
def log_rework_to_teams():
    """Endpoint to receive rework data and send a Teams notification."""
    data = request.get_json()
    if not data:
        return jsonify({"error": "No data provided"}), 400

    disposition = data.get('disposition', 'rework')
    webhook = TeamsWebhook.query.filter_by(webhook_name=disposition).first()
    if not webhook:
        return jsonify({'error': f'Webhook for disposition "{disposition}" not found'}), 404

    message = (
        "New Rework Log Entry\n\n"
        f"Job Number: {data.get('jobNo', 'N/A')}\n"
        f"Disposition: {data.get('disposition', 'N/A')}\n"
        f"Inspection Code: {data.get('inspCode', 'N/A')}\n"
        f"Inspection Description: {data.get('inspDescription', 'N/A')}\n"
        f"Clock Number: {data.get('clockNum', 'N/A')}\n"
        f"Quantity: {data.get('quantity', 'N/A')}\n"
        f"Comment: {data.get('Comment', '')}\n"
        f"Timestamp: {data.get('timestamp', 'N/A')}"
    )

    try:
        headers = {'Content-Type': 'application/json'}
        payload = {'text': message}
        response = requests.post(webhook.webhook_url, json=payload, headers=headers)
        response.raise_for_status()
        return jsonify({"message": "Rework logged and notification sent"}), 200
    except requests.exceptions.RequestException as e:
        logging.error(f"Error sending Teams notification: {e}")
        return jsonify({"error": f"Failed to send Teams notification: {e}"}), 500

@app.route('/api/job_orders', methods=['GET'])
def list_job_orders():
    """Returns a list of all job orders."""
    try:
        jobs = JobOrder.query.all()
        return jsonify([job.as_dict() for job in jobs])
    except Exception as e:
        return jsonify({'error': str(e)}), 500
        
@app.route('/update_hardware_location', methods=['POST'])
def update_hardware_location():
    """Updates the location of a hardware item."""
    data = request.get_json()
    job_number = data.get('job_number')
    location_name = data.get('location_name')

    if not job_number or not location_name:
        return jsonify({'error': 'Job number and location name are required'}), 400

    try:
        # Check if the location is predefined
        predefined_location = PredefinedLocation.query.filter_by(name=location_name).first()
        if not predefined_location:
            return jsonify({'error': 'Invalid location name'}), 400
        
        hardware = HardwareLocation.query.filter_by(job_number=job_number).first()
        if hardware:
            hardware.location = location_name
        else:
            hardware = HardwareLocation(job_number=job_number, location=location_name)
            db.session.add(hardware)
        
        db.session.commit()
        return jsonify({'message': f'Location for job {job_number} updated to {location_name}'})
    except Exception as e:
        db.session.rollback()
        return jsonify({'error': str(e)}), 500

@app.route('/get_hardware_location/<string:job_number>', methods=['GET'])
def get_hardware_location(job_number):
    """Retrieves the location of a hardware item."""
    try:
        hardware = HardwareLocation.query.filter_by(job_number=job_number).first()
        if hardware:
            return jsonify({'job_number': hardware.job_number, 'location': hardware.location, 'timestamp': hardware.timestamp.isoformat()})
        return jsonify({'error': 'Hardware not found'}), 404
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# --- Webhook API Endpoints (must be above main block) ---
@app.route('/api/disposition_webhooks', methods=['POST'])
def save_disposition_webhooks():
    import traceback
    try:
        data = request.get_json()
        if not isinstance(data, dict):
            return jsonify({'error': 'Invalid data format. Expected a JSON object.'}), 400
        for disposition, webhook_url in data.items():
            if not webhook_url:
                continue  # skip empty fields
            # Update if exists, else create
            existing = TeamsWebhook.query.filter_by(webhook_name=disposition).first()
            if existing:
                existing.webhook_url = webhook_url
            else:
                db.session.add(TeamsWebhook(webhook_name=disposition, webhook_url=webhook_url))
        db.session.commit()
        return jsonify({'message': 'Webhooks saved successfully.'})
    except Exception as e:
        db.session.rollback()
        print('Error saving webhooks:', e)
        print(traceback.format_exc())
        return jsonify({'error': f'Failed to save webhooks: {str(e)}'}), 500

@app.route('/api/disposition_webhooks', methods=['GET'])
def get_disposition_webhooks():
    webhooks = TeamsWebhook.query.all()
    return jsonify([
        {'disposition': w.webhook_name, 'webhook_url': w.webhook_url}
        for w in webhooks
    ])

# --- Debug Endpoint: List All Label Templates ---
@app.route('/api/debug/label_templates', methods=['GET'])
def list_label_templates():
    templates = LabelTemplate.query.all()
    return jsonify([
        {'id': t.id, 'name': t.name, 'xml_length': len(t.xml) if t.xml else 0}
        for t in templates
    ])

# --- Main Application Execution ---
if __name__ == '__main__':
    import os
    import sys
    print("\n========== FLASK ROUTES ==========")
    for rule in app.url_map.iter_rules():
        print(f"{rule.endpoint}: {rule}")
    print("========== END FLASK ROUTES =========\n")

    with app.app_context():
        db.create_all()
        # Only reset JobOrder table automatically. All other tables (webhooks, bin locations, exceptions, etc.) are only updated via intentional UI/API actions.
        update_job_orders_from_excel()  # Only updates JobOrder table

    def scheduled_job_wrapper():
        with app.app_context():
            update_job_orders_from_excel()

    scheduler.add_job(id='update_db_job', func=scheduled_job_wrapper, trigger='interval', seconds=10)
    scheduler.start()
    app.run(host='0.0.0.0', port=5000, debug=False)




# --- File Upload Endpoints ---
from flask import current_app

# --- Inspection Codes API Endpoint ---
@app.route('/api/inspection_codes', methods=['GET'])
def api_inspection_codes():
    try:
        codes = InspectionCode.query.all()
        return jsonify([
            {'code': c.code, 'description': c.description}
            for c in codes
        ])
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/upload_inspcode', methods=['POST'])
def upload_inspcode():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part in request'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    if not allowed_file(file.filename, {'xlsx'}):
        return jsonify({'error': 'Invalid file type. Only .xlsx allowed.'}), 400
    try:
        save_path = os.path.join(current_app.root_path, 'inspectioncode.xlsx')
        file.save(save_path)
        return jsonify({'message': 'inspectioncode.xlsx uploaded and replaced successfully.'})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/upload_reworklabel', methods=['POST'])
def upload_reworklabel():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part in request'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    if not allowed_file(file.filename, {'dymo'}):
        return jsonify({'error': 'Invalid file type. Only .dymo allowed.'}), 400
    try:
        save_path = os.path.join(current_app.root_path, 'ReworkLabel.dymo')
        file.save(save_path)
        return jsonify({'message': 'ReworkLabel.dymo uploaded and replaced successfully.'})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/disposition_webhooks', methods=['POST'])
def save_disposition_webhooks():
    import traceback
    try:
        data = request.get_json()
        if not isinstance(data, dict):
            return jsonify({'error': 'Invalid data format. Expected a JSON object.'}), 400
        for disposition, webhook_url in data.items():
            if not webhook_url:
                continue  # skip empty fields
            # Update if exists, else create
            existing = TeamsWebhook.query.filter_by(webhook_name=disposition).first()
            if existing:
                existing.webhook_url = webhook_url
            else:
                db.session.add(TeamsWebhook(webhook_name=disposition, webhook_url=webhook_url))
        db.session.commit()
        return jsonify({'message': 'Webhooks saved successfully.'})
    except Exception as e:
        db.session.rollback()
        print('Error saving webhooks:', e)
        print(traceback.format_exc())
        return jsonify({'error': f'Failed to save webhooks: {str(e)}'}), 500

@app.route('/api/disposition_webhooks', methods=['GET'])
def get_disposition_webhooks():
    webhooks = TeamsWebhook.query.all()
    return jsonify([
        {'disposition': w.webhook_name, 'webhook_url': w.webhook_url}
        for w in webhooks
    ])
