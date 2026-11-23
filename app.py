import os
import subprocess
import shutil
import json
import time
import hashlib
import secrets
from functools import wraps
from flask import Flask, request, jsonify, render_template, session, redirect, url_for
from werkzeug.utils import secure_filename
from pdf2image import convert_from_path
from pymongo import MongoClient
from pymongo.errors import DuplicateKeyError
from bson import ObjectId

app = Flask(__name__)
app.secret_key = secrets.token_hex(32)  # Secure random secret key for sessions

app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['SLIDES_FOLDER'] = 'static/slides'
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

ALLOWED_EXTENSIONS = {'ppt', 'pptx'}

# Windows LibreOffice full path
SOFFICE = r'C:\Program Files\LibreOffice\program\soffice.exe'

# ─── MongoDB Setup ────────────────────────────────────────────────────────────

client = MongoClient("mongodb://localhost:27017/")
db = client["gestureslides"]
users_collection = db["users"]

# Create unique index on username (replaces SQLite UNIQUE constraint)
users_collection.create_index("username", unique=True)

# ─── Password Helpers ─────────────────────────────────────────────────────────

def hash_password(password):
    """Hash password with SHA-256 + salt."""
    salt = secrets.token_hex(16)
    hashed = hashlib.sha256((salt + password).encode()).hexdigest()
    return f"{salt}:{hashed}"

def verify_password(password, stored_hash):
    """Verify a password against its stored hash."""
    try:
        salt, hashed = stored_hash.split(':')
        return hashlib.sha256((salt + password).encode()).hexdigest() == hashed
    except Exception:
        return False

# ─── Auth Helpers ─────────────────────────────────────────────────────────────

def login_required(f):
    """Decorator to protect routes — redirects to /login if not logged in."""
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            # For API routes, return JSON error
            if request.path.startswith('/api/'):
                return jsonify({'error': 'Not logged in'}), 401
            return redirect(url_for('login_page'))
        return f(*args, **kwargs)
    return decorated

def current_user_id():
    return session.get('user_id')

def current_username():
    return session.get('username')

# ─── Middleware ───────────────────────────────────────────────────────────────

@app.after_request
def after_request(response):
    response.headers['Access-Control-Allow-Origin'] = '*'
    response.headers['Access-Control-Allow-Methods'] = 'GET, POST, DELETE, OPTIONS'
    response.headers['Access-Control-Allow-Headers'] = 'Content-Type'
    return response

@app.route('/api/<path:path>', methods=['OPTIONS'])
def options_handler(path):
    return jsonify({})

# ─── Auth Routes ──────────────────────────────────────────────────────────────

@app.route('/login')
def login_page():
    if 'user_id' in session:
        return redirect(url_for('index'))
    return render_template('login.html')

@app.route('/api/signup', methods=['POST'])
def signup():
    data = request.get_json()
    username = (data.get('username') or '').strip()
    password = data.get('password') or ''

    if not username or not password:
        return jsonify({'error': 'Username and password are required.'}), 400
    if len(username) < 3:
        return jsonify({'error': 'Username must be at least 3 characters.'}), 400
    if len(username) > 30:
        return jsonify({'error': 'Username too long (max 30 chars).'}), 400
    if not username.replace('_', '').isalnum():
        return jsonify({'error': 'Username can only contain letters, numbers, underscores.'}), 400
    if len(password) < 4:
        return jsonify({'error': 'Password must be at least 4 characters.'}), 400

    try:
        user_doc = {
            "username": username,
            "password_hash": hash_password(password),
            "created_at": int(time.time())
        }
        result = users_collection.insert_one(user_doc)
        session['user_id'] = str(result.inserted_id)
        session['username'] = username
        return jsonify({'success': True, 'username': username})
    except DuplicateKeyError:
        return jsonify({'error': 'Username already taken.'}), 409
    except Exception as e:
        return jsonify({'error': 'Signup failed. Please try again.'}), 500

@app.route('/api/login', methods=['POST'])
def login():
    data = request.get_json()
    username = (data.get('username') or '').strip()
    password = data.get('password') or ''

    if not username or not password:
        return jsonify({'error': 'Username and password are required.'}), 400

    user = users_collection.find_one({"username": username})

    if not user or not verify_password(password, user['password_hash']):
        return jsonify({'error': 'Invalid username or password.'}), 401

    session['user_id'] = str(user['_id'])
    session['username'] = user['username']
    return jsonify({'success': True, 'username': user['username']})

@app.route('/api/logout', methods=['POST'])
def logout():
    session.clear()
    return jsonify({'success': True})

@app.route('/api/me')
def me():
    if 'user_id' in session:
        return jsonify({'logged_in': True, 'username': current_username()})
    return jsonify({'logged_in': False})

# ─── File Helpers ─────────────────────────────────────────────────────────────

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def user_slides_dir(user_id, presentation_id=None):
    """Get the slides directory for a specific user (and optionally a presentation)."""
    base = os.path.join(app.config['SLIDES_FOLDER'], f'user_{user_id}')
    if presentation_id:
        return os.path.join(base, presentation_id)
    return base

def convert_pptx_to_images(pptx_path, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    pptx_path = os.path.abspath(pptx_path)
    pdf_dir = os.path.dirname(pptx_path)

    print(f"Converting: {pptx_path}")

    if not os.path.exists(SOFFICE):
        raise Exception(
            f"LibreOffice not found at {SOFFICE}. "
            "Please install from https://www.libreoffice.org/download/libreoffice/"
        )

    result = subprocess.run(
        [SOFFICE, '--headless', '--convert-to', 'pdf', '--outdir', pdf_dir, pptx_path],
        capture_output=True, text=True, timeout=120
    )

    print(f"LibreOffice output: {result.stdout} {result.stderr}")

    base_name = os.path.splitext(os.path.basename(pptx_path))[0]
    pdf_path = os.path.join(pdf_dir, base_name + '.pdf')

    if not os.path.exists(pdf_path):
        raise Exception(f"PDF not generated. LibreOffice said: {result.stderr or result.stdout}")

    try:
        images = convert_from_path(pdf_path, dpi=150)
    except Exception as e:
        raise Exception(
            f"Poppler error: {str(e)}. "
            "Download poppler from https://github.com/oschwartz10612/poppler-windows/releases "
            "and add C:\\poppler\\Library\\bin to PATH"
        )

    image_paths = []
    for i, img in enumerate(images):
        img_filename = f'slide_{i+1:03d}.jpg'
        img_path = os.path.join(output_dir, img_filename)
        img.save(img_path, 'JPEG', quality=85)
        image_paths.append(img_filename)
        print(f"Saved slide {i+1}")

    try:
        os.remove(pdf_path)
    except:
        pass

    return image_paths

# ─── Page Routes ──────────────────────────────────────────────────────────────

@app.route('/')
@login_required
def index():
    return render_template('index.html', username=current_username())

@app.route('/present/<presentation_id>')
@login_required
def present(presentation_id):
    # Verify this presentation belongs to the current user
    slides_dir = user_slides_dir(current_user_id(), presentation_id)
    meta_path = os.path.join(slides_dir, 'meta.json')
    if not os.path.exists(meta_path):
        return redirect(url_for('index'))
    return render_template('present.html', presentation_id=presentation_id)

# ─── API Routes ───────────────────────────────────────────────────────────────

@app.route('/api/upload', methods=['POST'])
@login_required
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    if not allowed_file(file.filename):
        return jsonify({'error': 'Only .ppt and .pptx files are allowed'}), 400

    filename = secure_filename(file.filename)
    presentation_id = str(int(time.time() * 1000))
    user_id = current_user_id()

    # Store uploads under user-specific folder
    user_upload_dir = os.path.join(app.config['UPLOAD_FOLDER'], f'user_{user_id}')
    os.makedirs(user_upload_dir, exist_ok=True)
    upload_path = os.path.join(user_upload_dir, presentation_id + '_' + filename)
    file.save(upload_path)
    print(f"File saved: {upload_path}")

    slides_dir = user_slides_dir(user_id, presentation_id)

    try:
        image_files = convert_pptx_to_images(upload_path, slides_dir)
        meta = {
            'id': presentation_id,
            'user_id': user_id,
            'filename': filename,
            'slide_count': len(image_files),
            'slides': image_files,
            'uploaded_at': int(time.time())
        }
        with open(os.path.join(slides_dir, 'meta.json'), 'w') as f:
            json.dump(meta, f)

        return jsonify({
            'success': True,
            'presentation_id': presentation_id,
            'filename': filename,
            'slide_count': len(image_files)
        })

    except Exception as e:
        print(f"ERROR: {e}")
        if os.path.exists(upload_path):
            os.remove(upload_path)
        if os.path.exists(slides_dir):
            shutil.rmtree(slides_dir)
        return jsonify({'error': str(e)}), 500

@app.route('/api/presentations', methods=['GET'])
@login_required
def list_presentations():
    presentations = []
    slides_dir = user_slides_dir(current_user_id())
    if os.path.exists(slides_dir):
        for folder in os.listdir(slides_dir):
            meta_path = os.path.join(slides_dir, folder, 'meta.json')
            if os.path.exists(meta_path):
                with open(meta_path) as f:
                    meta = json.load(f)
                presentations.append(meta)
    presentations.sort(key=lambda x: x['id'], reverse=True)
    return jsonify(presentations)

@app.route('/api/presentation/<presentation_id>', methods=['GET'])
@login_required
def get_presentation(presentation_id):
    meta_path = os.path.join(user_slides_dir(current_user_id(), presentation_id), 'meta.json')
    if not os.path.exists(meta_path):
        return jsonify({'error': 'Presentation not found'}), 404

    with open(meta_path) as f:
        meta = json.load(f)

    # Security: make sure this presentation belongs to logged-in user
    if meta.get('user_id') != current_user_id():
        return jsonify({'error': 'Access denied'}), 403

    return jsonify(meta)

@app.route('/api/presentation/<presentation_id>', methods=['DELETE'])
@login_required
def delete_presentation(presentation_id):
    user_id = current_user_id()
    slides_dir = user_slides_dir(user_id, presentation_id)

    # Verify ownership before deleting
    meta_path = os.path.join(slides_dir, 'meta.json')
    if os.path.exists(meta_path):
        with open(meta_path) as f:
            meta = json.load(f)
        if meta.get('user_id') != user_id:
            return jsonify({'error': 'Access denied'}), 403
        shutil.rmtree(slides_dir)

    # Clean up upload file
    user_upload_dir = os.path.join(app.config['UPLOAD_FOLDER'], f'user_{user_id}')
    if os.path.exists(user_upload_dir):
        for fname in os.listdir(user_upload_dir):
            if fname.startswith(presentation_id + '_'):
                os.remove(os.path.join(user_upload_dir, fname))

    return jsonify({'success': True})

# ─── Static slide serving (user-scoped path) ──────────────────────────────────
# Slides are at: /static/slides/user_<id>/<presentation_id>/slide_001.jpg
# This is handled automatically by Flask's static file serving.

# ─── Run ──────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    os.makedirs('uploads', exist_ok=True)
    os.makedirs('static/slides', exist_ok=True)
    print("GestureSlides starting...")
    print("Open http://localhost:5000")
    app.run(debug=True, port=5000, host='0.0.0.0')
