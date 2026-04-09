from datetime import datetime, UTC
from pathlib import Path
import json
import uuid

from flask import Flask, abort, flash, jsonify, redirect, render_template, request, send_file, url_for
from werkzeug.utils import secure_filename

APP_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = APP_DIR / 'uploads'
JOBS_DIR = APP_DIR / 'jobs'
UPLOAD_DIR.mkdir(exist_ok=True)
JOBS_DIR.mkdir(exist_ok=True)
ALLOWED_EXTENSIONS = {'.xlsx'}
MAX_UPLOAD_MB = 10
STALE_FILE_HOURS = 24
PREVIEW_LIMIT = 20

app = Flask(__name__)
app.secret_key = 'onpe-web-secret'
app.config['MAX_CONTENT_LENGTH'] = MAX_UPLOAD_MB * 1024 * 1024


def utc_now_iso():
    return datetime.now(UTC).isoformat(timespec='seconds').replace('+00:00', 'Z')


def allowed_file(filename):
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS


def cleanup_old_artifacts():
    cutoff = datetime.now(UTC).timestamp() - (STALE_FILE_HOURS * 3600)
    for folder in (UPLOAD_DIR, JOBS_DIR):
        for file_path in folder.iterdir():
            try:
                if file_path.is_file() and file_path.stat().st_mtime < cutoff:
                    file_path.unlink()
            except OSError:
                continue


def job_path(job_id):
    return JOBS_DIR / f'{job_id}.json'


def read_job(job_id):
    path = job_path(job_id)
    if not path.exists():
        return None
    return json.loads(path.read_text(encoding='utf-8'))


def write_job(job):
    job['updated_at'] = utc_now_iso()
    job_path(job['id']).write_text(json.dumps(job, ensure_ascii=False, indent=2), encoding='utf-8')


def create_job(uploaded):
    safe_name = secure_filename(uploaded.filename)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    job_id = uuid.uuid4().hex[:12]
    stored_name = f'{timestamp}_{job_id}_{safe_name}'
    uploaded_path = UPLOAD_DIR / stored_name
    uploaded.save(uploaded_path)

    job = {
        'id': job_id,
        'original_name': safe_name,
        'uploaded_name': stored_name,
        'result_name': '',
        'status': 'pending',
        'message': 'Esperando al worker local para consultar ONPE.',
        'created_at': utc_now_iso(),
        'updated_at': utc_now_iso(),
        'total': 0,
        'errors': 0,
        'elapsed': '',
        'records': [],
    }
    write_job(job)
    return job


def list_jobs():
    jobs = []
    for path in sorted(JOBS_DIR.glob('*.json'), reverse=True):
        try:
            jobs.append(json.loads(path.read_text(encoding='utf-8')))
        except Exception:
            continue
    return jobs[:10]


def next_pending_job():
    for path in sorted(JOBS_DIR.glob('*.json')):
        try:
            job = json.loads(path.read_text(encoding='utf-8'))
        except Exception:
            continue
        if job.get('status') == 'pending':
            job['status'] = 'processing'
            job['message'] = 'Worker local procesando archivo.'
            write_job(job)
            return job
    return None


@app.route('/', methods=['GET'])
def index():
    cleanup_old_artifacts()
    job_id = request.args.get('job_id', '').strip()
    selected_job = read_job(job_id) if job_id else None
    return render_template(
        'index.html',
        max_upload_mb=MAX_UPLOAD_MB,
        job=selected_job,
        jobs=list_jobs(),
        preview_limit=PREVIEW_LIMIT,
    )


@app.route('/procesar', methods=['POST'])
def procesar():
    cleanup_old_artifacts()
    uploaded = request.files.get('excel_file')
    if not uploaded or not uploaded.filename:
        flash('Selecciona un archivo Excel para continuar.', 'error')
        return redirect(url_for('index'))

    if not allowed_file(uploaded.filename):
        flash('Solo se admiten archivos .xlsx', 'error')
        return redirect(url_for('index'))

    job = create_job(uploaded)
    flash('Archivo recibido. Ahora el worker local debe procesarlo.', 'success')
    return redirect(url_for('index', job_id=job['id']))


@app.route('/descargar/<job_id>', methods=['GET'])
def descargar(job_id):
    job = read_job(job_id)
    if not job:
        abort(404)

    if job.get('status') != 'completed' or not job.get('result_name'):
        flash('Ese archivo todavia no esta listo para descargar.', 'error')
        return redirect(url_for('index', job_id=job_id))

    file_path = UPLOAD_DIR / job['result_name']
    if not file_path.exists():
        flash('El archivo procesado ya no esta disponible.', 'error')
        return redirect(url_for('index', job_id=job_id))

    return send_file(file_path, as_attachment=True, download_name=job['original_name'])


@app.route('/api/jobs/next', methods=['POST'])
def api_next_job():
    cleanup_old_artifacts()
    job = next_pending_job()
    if not job:
        return jsonify({'job': None})

    return jsonify(
        {
            'job': {
                'id': job['id'],
                'original_name': job['original_name'],
                'download_url': url_for('api_download_job_file', job_id=job['id'], _external=True),
            }
        }
    )


@app.route('/api/jobs/<job_id>/file', methods=['GET'])
def api_download_job_file(job_id):
    job = read_job(job_id)
    if not job:
        abort(404)

    file_path = UPLOAD_DIR / job['uploaded_name']
    if not file_path.exists():
        abort(404)

    return send_file(file_path, as_attachment=True, download_name=job['original_name'])


@app.route('/api/jobs/<job_id>/complete', methods=['POST'])
def api_complete_job(job_id):
    job = read_job(job_id)
    if not job:
        abort(404)

    result_file = request.files.get('result_file')
    if not result_file or not result_file.filename:
        return jsonify({'error': 'Falta result_file'}), 400

    result_name = f"result_{job['uploaded_name']}"
    result_path = UPLOAD_DIR / result_name
    result_file.save(result_path)

    records_raw = request.form.get('records_json', '[]')
    try:
        records = json.loads(records_raw)
    except json.JSONDecodeError:
        records = []

    job['status'] = 'completed'
    job['message'] = request.form.get('message', 'Proceso completado correctamente.')
    job['result_name'] = result_name
    job['total'] = int(request.form.get('total', '0') or 0)
    job['errors'] = int(request.form.get('errors', '0') or 0)
    job['elapsed'] = request.form.get('elapsed', '')
    job['records'] = records[:PREVIEW_LIMIT]
    write_job(job)

    return jsonify({'ok': True})


@app.route('/api/jobs/<job_id>/error', methods=['POST'])
def api_error_job(job_id):
    job = read_job(job_id)
    if not job:
        abort(404)

    payload = request.get_json(silent=True) or {}
    job['status'] = 'error'
    job['message'] = payload.get('message', 'El worker local reporto un error.')
    write_job(job)
    return jsonify({'ok': True})


@app.route('/api/jobs/<job_id>/status', methods=['GET'])
def api_job_status(job_id):
    job = read_job(job_id)
    if not job:
        return jsonify({'error': 'Job no encontrado'}), 404
    return jsonify(job)


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8000, debug=False)
