import os
import re
import tempfile
import zipfile

from pptx_merge import Target, make_blank_base
from lesson_detect import analyze_lesson, LessonError
from verb_fix import fix_future_to_past

PACKAGES_DIR = os.path.join(os.path.dirname(__file__), 'static_packages')

TURMA_FILES = {
    'adultos': {
        'head': os.path.join(PACKAGES_DIR, 'pkg_head_adultos.pptx'),
        'tail': os.path.join(PACKAGES_DIR, 'pkg_tail_adultos.pptx'),
    },
    'jovens': {
        'head': os.path.join(PACKAGES_DIR, 'pkg_head_jovens.pptx'),
        'tail': os.path.join(PACKAGES_DIR, 'pkg_tail_jovens.pptx'),
    },
}


def _slide_count(unpacked_dir):
    slides_dir = os.path.join(unpacked_dir, 'ppt/slides')
    return len([f for f in os.listdir(slides_dir) if re.match(r'^slide\d+\.xml$', f)])


def build_encerramento(turma, lesson_paths, out_path, work_root):
    """turma: 'adultos' or 'jovens'. lesson_paths: list of uploaded .pptx paths
    (any order — lesson number is auto-detected from each file's title slide).
    Returns (out_path, warnings)."""
    if turma not in TURMA_FILES:
        raise LessonError(f'Turma desconhecida: {turma}')

    warnings = []
    lessons = []
    for path in lesson_paths:
        try:
            lessons.append(analyze_lesson(path))
        except (LessonError, ValueError) as e:
            raise LessonError(f'Erro no arquivo "{os.path.basename(path)}": {e}')

    numbers = [l['number'] for l in lessons]
    if len(set(numbers)) != len(numbers):
        dupes = sorted({n for n in numbers if numbers.count(n) > 1})
        raise LessonError(f'Lições duplicadas (mesmo número): {dupes}')
    lessons.sort(key=lambda l: l['number'])

    missing = sorted(set(range(1, 14)) - set(numbers))
    if missing:
        warnings.append(f'Faltando lições: {missing} (gerando mesmo assim com {len(lessons)} lição(ões))')
    extra = sorted(n for n in numbers if n > 13 or n < 1)
    if extra:
        warnings.append(f'Números de lição fora do intervalo 1-13: {extra}')

    blank_dir = os.path.join(work_root, 'blank_work')
    blank_path = os.path.join(work_root, 'blank_base.pptx')
    make_blank_base(TURMA_FILES[turma]['head'], blank_path, blank_dir)

    t = Target(blank_path, os.path.join(work_root, 'target'))

    head_unpacked = os.path.join(work_root, 'unpack_head')
    with zipfile.ZipFile(TURMA_FILES[turma]['head']) as z:
        z.extractall(head_unpacked)
    head_slide_count = _slide_count(head_unpacked)
    for n in range(1, head_slide_count + 1):
        t.copy_slide(head_unpacked, f'ppt/slides/slide{n}.xml', 'HEAD')

    for lesson in lessons:
        unpacked = os.path.join(work_root, f"unpack_L{lesson['number']:02d}")
        with zipfile.ZipFile(lesson['path']) as z:
            z.extractall(unpacked)
        tag = f"L{lesson['number']:02d}"
        for kind, idx in lesson['order']:
            slide_relpath = f'ppt/slides/slide{idx + 1}.xml'
            dst = t.copy_slide(unpacked, slide_relpath, tag)
            if kind == 'title':
                t.strip_shape_by_text(dst, 'Nome do Professor')
            elif kind == 'intro':
                t.apply_text_transform(dst, fix_future_to_past)

    tail_unpacked = os.path.join(work_root, 'unpack_tail')
    with zipfile.ZipFile(TURMA_FILES[turma]['tail']) as z:
        z.extractall(tail_unpacked)
    tail_slide_count = _slide_count(tail_unpacked)
    for n in range(1, tail_slide_count + 1):
        t.copy_slide(tail_unpacked, f'ppt/slides/slide{n}.xml', 'TAIL')

    t.save(out_path)
    return out_path, warnings
