import re
import io
from difflib import SequenceMatcher
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_COLOR_INDEX
from rapidfuzz.distance import Levenshtein


PAGE_MARKER_RE = re.compile(r'^\s*[Сс]\.\s*\d+[\d\s\-–—]*\s*$')
DASH_MARKER_RE = re.compile(r'^--$')


def is_page_marker(text: str) -> bool:
    t = text.strip()
    return bool(PAGE_MARKER_RE.match(t) or DASH_MARKER_RE.match(t))


def extract_left_paragraphs(cell) -> tuple:
    """
    Параграфи між PAGE MARKER і першим порожнім параграфом.
    Порожний параграф = початок коментарів складача (ігноруємо).
    """
    paragraphs = cell.paragraphs
    marker_idx = None
    for i, para in enumerate(paragraphs):
        if is_page_marker(para.text):
            marker_idx = i
            break

    if marker_idx is None:
        return [p for p in paragraphs if p.text.strip()], "WARNING: page marker not found"

    result = []
    for para in paragraphs[marker_idx + 1:]:
        if para.text.strip() == '':
            break
        result.append(para)

    return result, None


COMMENT_RE = re.compile(r'^(\d+\.|\u041fосилання на:|\u0422ам само$)')


def extract_right_paragraphs(cell) -> tuple:
    """
    Бібліографія → PAGE MARKER → текст оригіналу → (коментарі).
    Беремо ОСТАННІЙ PAGE MARKER (бо бібліографія може містити 'С.').
    """
    paragraphs = cell.paragraphs
    marker_idx = None
    for i, para in enumerate(paragraphs):
        if is_page_marker(para.text):
            marker_idx = i  # не break — шукаємо останній

    if marker_idx is None:
        return [p for p in paragraphs if p.text.strip()], "WARNING: page marker not found"

    result = []
    for para in paragraphs[marker_idx + 1:]:
        txt = para.text.strip()
        if not txt:
            continue
        if COMMENT_RE.match(txt):
            break
        result.append(para)

    return result, None


def tokenize_paragraphs(paragraphs: list) -> list:
    """
    Повертає [(token, para_idx), ...].
    '\n' між параграфами = маркер нового абзацу (не підсвічується).
    """
    result = []
    for idx, para in enumerate(paragraphs):
        tokens = re.split(r'(\s+)', para.text)
        for t in tokens:
            if t:
                result.append((t, idx))
        if idx < len(paragraphs) - 1:
            result.append(('\n', idx))
    return result


def is_word(token: str) -> bool:
    return bool(re.search(r'\w', token))


def fuzzy_match(a: str, b: str, threshold: float) -> bool:
    if not a or not b:
        return False
    min_len = min(len(a), len(b))
    effective_threshold = 70.0 if min_len <= 4 else threshold
    ratio = Levenshtein.normalized_similarity(a, b) * 100
    return ratio >= effective_threshold


def align_tokens(
    left_pairs: list,
    right_pairs: list,
    threshold: float = 75.0
) -> tuple:
    """
    Повертає [(token, para_idx, status), ...]
    status: 'match' | 'diff' | None (пробіл/пунктуація/перенос)
    """
    left_words  = [(t, i) for t, i in left_pairs  if is_word(t)]
    right_words = [(t, i) for t, i in right_pairs if is_word(t)]

    left_keys  = [re.sub(r'[^\w]', '', t.lower()) for t, _ in left_words]
    right_keys = [re.sub(r'[^\w]', '', t.lower()) for t, _ in right_words]

    matcher = SequenceMatcher(None, left_keys, right_keys, autojunk=False)

    left_status  = {i: 'diff' for i in range(len(left_words))}
    right_status = {i: 'diff' for i in range(len(right_words))}

    for opcode, i1, i2, j1, j2 in matcher.get_opcodes():
        if opcode == 'equal':
            for i in range(i1, i2):
                left_status[i] = 'match'
            for j in range(j1, j2):
                right_status[j] = 'match'
        elif opcode == 'replace':
            for i, j in zip(range(i1, i2), range(j1, j2)):
                if fuzzy_match(left_keys[i], right_keys[j], threshold):
                    left_status[i]  = 'match'
                    right_status[j] = 'match'

    def rebuild(token_pairs, words, statuses):
        word_iter = iter(enumerate(words))
        current = next(word_iter, None)
        result = []
        for token, para_idx in token_pairs:
            if is_word(token):
                w_idx = current[0] if current is not None else None
                status = statuses.get(w_idx, 'diff') if w_idx is not None else 'diff'
                result.append((token, para_idx, status))
                current = next(word_iter, None)
            else:
                result.append((token, para_idx, None))
        return result

    return (
        rebuild(left_pairs,  left_words,  left_status),
        rebuild(right_pairs, right_words, right_status)
    )


HIGHLIGHT = {
    'match': WD_COLOR_INDEX.YELLOW,
    'diff':  WD_COLOR_INDEX.TURQUOISE,
}


def rewrite_cell(cell, token_result: list, font_name: str, font_size: int):
    """
    Очищає ячейку і записує токени з підсвічуванням.
    token_result: [(token, para_idx, status), ...]
    status: 'match' | 'diff' | None
    Токен '\n' → створює новий параграф у ячейці.
    """
    tc = cell._tc
    ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    for p in tc.findall(f'{{{ns}}}p')[1:]:
        tc.remove(p)

    first_para = cell.paragraphs[0]
    first_para.clear()
    current_para = first_para

    for token, para_idx, status in token_result:
        if token == '\n':
            current_para = cell.add_paragraph()
            continue

        run = current_para.add_run(token)
        run.font.name = font_name
        run.font.size = Pt(font_size)
        run.bold   = False
        run.italic = False
        run.font.highlight_color = HIGHLIGHT.get(status)


def process_document(
    docx_bytes: bytes,
    threshold: float,
    font_name: str,
    font_size: int,
    progress_callback=None
) -> tuple:
    """
    Повертає: (result_bytes, warnings, stats)
    stats = {'processed': int, 'warnings': int, 'skipped': int}
    """
    warnings = []
    stats = {'processed': 0, 'warnings': 0, 'skipped': 0}

    doc = Document(io.BytesIO(docx_bytes))
    table = doc.tables[0]
    total_rows = len(table.rows)

    for row_idx, row in enumerate(table.rows):
        if progress_callback:
            progress_callback(row_idx, total_rows, row_idx + 1)

        if len(row.cells) < 2:
            stats['skipped'] += 1
            continue

        left_cell, right_cell = row.cells[0], row.cells[1]

        if not left_cell.text.strip() and not right_cell.text.strip():
            stats['skipped'] += 1
            continue

        left_paras,  warn_l = extract_left_paragraphs(left_cell)
        right_paras, warn_r = extract_right_paragraphs(right_cell)

        for warn, side in [(warn_l, 'ліва'), (warn_r, 'права')]:
            if warn:
                warnings.append(f"Рядок {row_idx + 1} ({side} колонка): {warn}")
                stats['warnings'] += 1

        if not left_paras or not right_paras:
            stats['skipped'] += 1
            continue

        left_pairs  = tokenize_paragraphs(left_paras)
        right_pairs = tokenize_paragraphs(right_paras)

        left_result, right_result = align_tokens(left_pairs, right_pairs, threshold)

        rewrite_cell(left_cell,  left_result,  font_name, font_size)
        rewrite_cell(right_cell, right_result, font_name, font_size)

        stats['processed'] += 1

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output.read(), warnings, stats
