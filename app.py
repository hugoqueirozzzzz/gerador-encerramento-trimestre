import os
import shutil
import tempfile
import uuid

from flask import Flask, render_template, request, send_file, flash, redirect, url_for

from build_deck import build_encerramento
from lesson_detect import LessonError

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'dev-secret-key-troque-em-producao')

MAX_CONTENT_LENGTH = 400 * 1024 * 1024  # 400 MB (13 pptx grandes)
app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH

BASE_TMP = tempfile.gettempdir()


@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')


@app.route('/gerar', methods=['POST'])
def gerar():
    turma = request.form.get('turma')
    if turma not in ('adultos', 'jovens'):
        flash('Selecione a turma (adultos ou jovens).')
        return redirect(url_for('index'))

    files = [f for f in request.files.getlist('licoes') if f and f.filename]
    if not files:
        flash('Envie pelo menos um arquivo de lição (.pptx).')
        return redirect(url_for('index'))

    job_dir = os.path.join(BASE_TMP, f'encerramento_{uuid.uuid4().hex}')
    os.makedirs(job_dir, exist_ok=True)
    uploads_dir = os.path.join(job_dir, 'uploads')
    os.makedirs(uploads_dir, exist_ok=True)

    lesson_paths = []
    try:
        for f in files:
            if not f.filename.lower().endswith('.pptx'):
                raise LessonError(f'"{f.filename}" não é um arquivo .pptx')
            dest = os.path.join(uploads_dir, f.filename)
            f.save(dest)
            lesson_paths.append(dest)

        out_path = os.path.join(job_dir, f'SLIDE_DE_ENCERRAMENTO_{turma.upper()}.pptx')
        result_path, warnings = build_encerramento(turma, lesson_paths, out_path, job_dir)

        for w in warnings:
            flash(w)

        return send_file(
            result_path,
            as_attachment=True,
            download_name=os.path.basename(result_path),
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
        )
    except LessonError as e:
        flash(str(e))
        return redirect(url_for('index'))
    except Exception as e:
        flash(f'Erro inesperado: {e}')
        return redirect(url_for('index'))
    finally:
        # cleanup happens on next request pass or via OS tmp cleanup;
        # keep it simple and don't block the response cleaning up now.
        pass


@app.route('/healthz')
def healthz():
    return {'status': 'ok'}


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
