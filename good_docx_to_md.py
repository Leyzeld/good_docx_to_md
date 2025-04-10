import os
import uuid
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.document import Document as _Document
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph
#from split_into_runs import split_words_into_runs

def iter_block_items(parent):
    if isinstance(parent, _Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError('Unsupported parent type')
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

def extract_images_from_run(run, doc, image_dir):
    image_paths = []
    for drawing in run._element.xpath('.//w:drawing'):
        embed = drawing.xpath('.//a:blip/@r:embed')
        if embed:
            rel_id = embed[0]
            image_part = doc.part.related_parts.get(rel_id)
            if image_part:
                ext = image_part.content_type.split('/')[-1]
                image_name = f'{uuid.uuid4()}.{ext}'
                image_path = os.path.join(image_dir, image_name)
                if not os.path.exists(image_dir):
                    os.makedirs(image_dir)
                with open(image_path, 'wb') as f:
                    f.write(image_part.blob)
                image_paths.append(image_path)
    return image_paths

def is_bullet(paragraph): # сделать разделение на точки и цифры
    return paragraph._p.pPr is not None and paragraph._p.pPr.numPr is not None
def get_hyperlinks(paragraph):
    links = []
    hyperlinks = paragraph.hyperlinks
    if hyperlinks:
        for hl in hyperlinks:
            links.append((hl.text, hl.address))
    return links


def merge_runs(runs):
    merged = []
    current = {'text': '', 'bold': None, 'italic': None}
    for run in runs:
        bold, italic = run.bold, run.italic
        if current['bold'] == bold and current['italic'] == italic:
            current['text'] += run.text
        else:
            if current['text']:
                merged.append(current.copy())
            current = {'text': run.text, 'bold': bold, 'italic': italic}
    if current['text']:
        merged.append(current)
    return merged


def format_text_block(text, bold, italic): #добавить больше опций
    if not text.strip():
        return ''
    text = text.strip()
    if bold:
        text = f'<strong>{text}</strong>'
    if italic:
        text = f'<em>{text}</em>'
    text = f' {text} '
    return text


def parse_paragraph(paragraph, doc, image_dir):
    lines = []
    result = []

    hyperlinks = get_hyperlinks(paragraph)

    for run in paragraph.iter_inner_content():
        if '<w:hyperlink' in run._element.xml:
            result.append(f'[{run.text}]({run.address}) ')
        else:
            result.append(format_text_block(run.text, run.bold, run.italic))

    for run in paragraph.runs:
        images = extract_images_from_run(run, doc, image_dir)
        for img_path in images:
            result.append(f'![Image]({img_path})')

    text_line = ''.join(result).strip()
    if not text_line:
        return []

    style = paragraph.style.name if paragraph.style else ''
    if style.startswith('Heading'):
        level = ''.join(filter(str.isdigit, style)) or '1'
        return [f'{"#" * int(level)} {text_line}']

    if is_bullet(paragraph): # сделать разделение на точки и цифры. Добавить подпункты
        return [f'+ {text_line}']

    return [text_line]


def parse_table(table):
    rows = table.rows
    if not rows:
        return ''
    md = []
    header = '| ' + ' | '.join(cell.text.strip().replace('\n', ' ') for cell in rows[0].cells) + ' |'
    separator = '| ' + ' | '.join(['---'] * len(rows[0].cells)) + ' |'
    md.append(header)
    md.append(separator)

    for row in rows[1:]:
        row_md = '| ' + ' | '.join(cell.text.strip().replace('\n', ' ') for cell in row.cells) + ' |'
        md.append(row_md)

    return '\n'.join(md)


def docx_to_md(docx_path, output_md_path, image_dir='images'):
    doc = Document(docx_path)
    md_lines = []

    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            lines = parse_paragraph(block, doc, image_dir)
            md_lines.extend(lines)
        elif isinstance(block, Table):
            md_lines.append(parse_table(block))

    with open(output_md_path, 'w', encoding='utf-8') as f:
        f.write('\n\n'.join(md_lines))

#split_words_into_runs('test.docx', 'out.docx')
docx_to_md('out.docx', 'output.md')
