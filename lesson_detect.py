import re
from extract_logic import extract_indices, slide_texts
from pptx import Presentation


class LessonError(Exception):
    pass


def detect_lesson_number(pptx_path, title_idx):
    p = Presentation(pptx_path)
    slide = list(p.slides)[title_idx]
    texts = slide_texts(slide)
    if not texts:
        raise LessonError(f'Não encontrei texto no slide de título de {pptx_path}')
    m = re.search(r'LIÇÃO\s*0*(\d+)', texts[0], re.I)
    if not m:
        raise LessonError(f'Não encontrei "LIÇÃO NN" no slide de título de {pptx_path}')
    return int(m.group(1))


def analyze_lesson(pptx_path):
    """Returns dict: {number, order (slide indices to copy, in order,
    tagged as ('title'|'intro'|'bumper', idx)), path}."""
    r = extract_indices(pptx_path)
    number = detect_lesson_number(pptx_path, r['title'])
    order = [('title', r['title'])]
    order += [('intro', i) for i in [r['intro_bumper']] + r['intro_content']]
    order += [('bumper', i) for i in r['main_bumpers']]
    return {'number': number, 'order': order, 'path': pptx_path}
