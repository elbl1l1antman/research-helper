"""Check the saved structure of the experimental HWPX chart containers."""

from base64 import b64encode
from hashlib import md5
import re
from pathlib import Path
from tempfile import NamedTemporaryFile
from xml.etree import ElementTree as ET
from zipfile import ZipFile


def verify_chart_document(path: Path, *, prefix: str, expected: tuple[bool, ...], titles=None) -> list[dict]:
    return _verify_document(path, prefix, expected, require_emf=True, titles=titles)


def _verify_document(path, prefix, expected, *, require_emf, titles=None):
    with ZipFile(path) as archive:
        manifest = ET.fromstring(archive.read('Contents/content.hpf'))
        media = {item.get('id'): item.get('href') for item in manifest.findall('./{*}manifest/{*}item')}
        sections = sorted((n for n in archive.namelist() if re.fullmatch(r'Contents/section\d+\.xml', n)),
                          key=lambda n: int(re.search(r'(\d+)\.xml$', n)[1]))
        tables = [table for section in sections for table in ET.fromstring(archive.read(section)).findall('.//{*}tbl')
                  if table.get('numberingType') == 'PICTURE']
        if titles is not None:
            tables = [table for table in tables if _caption_text(table.find('./{*}caption')).startswith(f'<그림 {prefix}-> ')]
        if len(tables) != len(expected):
            raise ValueError('Unexpected figure container count')
        figures = [_verify_container(archive, media, table, prefix, index, has_chart, require_emf)
                   for index, (table, has_chart) in enumerate(zip(tables, expected), 1)]
        if titles is not None and tuple(item['title'] for item in figures) != titles:
            raise ValueError('Unexpected figure caption titles')
        return figures


def _caption_text(caption):
    return ''.join(node.text or '' for node in caption.findall('.//{*}t')) if caption is not None else ''


def _is_emf(header):
    return len(header) >= 88 and header[:4] == b'\x01\0\0\0' and header[40:44] == b' EMF'


def _verify_container(archive, media, table, prefix, index, has_chart, require_emf):
    if (table.get('rowCnt'), table.get('colCnt')) != ('1', '1') or table.find('./{*}pos').get('treatAsChar') != '1':
        raise ValueError(f'Figure {index}: expected inline 1x1 table')
    caption = table.find('./{*}caption')
    if caption is None or caption.get('side') != 'TOP':
        raise ValueError(f'Figure {index}: expected top caption')
    fields = caption.findall('.//{*}autoNum')
    if [(node.get('numType'), node.get('num')) for node in fields] != [('PICTURE', str(index))]:
        raise ValueError(f'Figure {index}: invalid picture number sequence')
    text = _caption_text(caption)
    leader = f'<그림 {prefix}-> '
    if not text.startswith(leader) or not text[len(leader):].strip():
        raise ValueError(f'Figure {index}: invalid caption prefix or empty title')
    cell = table.find('./{*}tr/{*}tc')
    content = cell.find('./{*}subList')
    pictures = content.findall('.//{*}pic')
    if len(pictures) != int(has_chart) or any(node.tag.rsplit('}', 1)[-1] in {'ole', 'tbl', 'container'} for node in content.iter()):
        raise ValueError(f'Figure {index}: unexpected cell contents')
    if any((node.text or '').strip() for node in content.findall('.//{*}t')):
        raise ValueError(f'Figure {index}: expected no text inside the chart cell')
    target = None
    if has_chart:
        picture, = pictures
        if picture.get('numberingType') != 'NONE':
            raise ValueError(f'Figure {index}: inner picture must be NONE')
        img = picture.find('.//{*}img')
        target = media.get(img.get('binaryItemIDRef')) if img is not None else None
        if not target or not target.startswith('BinData/') or target not in archive.namelist():
            raise ValueError(f'Figure {index}: no embedded EMF link')
        with archive.open(target) as stream:
            header = stream.read(88)
        if not _is_emf(header) and (require_emf or not header.startswith((b'\x01\0\x09\0\0\x03', b'\x02\0\x09\0\0\x03', b'\xd7\xcd\xc6\x9a'))):
            raise ValueError(f'Figure {index}: embedded media is not EMF')
        margins = cell.find('./{*}cellMargin')
        available = int(cell.find('./{*}cellSz').get('width')) - sum(int(margins.get(side, '0')) for side in ('left', 'right'))
        if not 0 < int(picture.find('./{*}sz').get('width')) <= available:
            raise ValueError(f'Figure {index}: picture exceeds cell width')
    return {'number': index, 'title': text[len(leader):], 'emf': target}


def preserve_chart_emfs(source: Path, target: Path, *, prefix: str, originals: tuple[bytes | None, ...]) -> list[dict]:
    """Replace only verified chart media in a new probe file, not a user document."""
    if target.exists() or source.resolve() == target.resolve():
        raise FileExistsError(target)
    expected = tuple(data is not None for data in originals)
    figures = _verify_document(source, prefix, expected, require_emf=False)
    sources = [figure['emf'] for figure in figures if figure['emf']]
    if len(sources) != len(set(sources)):
        raise ValueError('Expected independent media for each chart')
    if any(not _is_emf(data) for data in originals if data is not None):
        raise ValueError('PowerPoint source is not EMF')

    replacements = {figure['emf']: (str(Path(figure['emf']).with_suffix('.emf')).replace('\\', '/'), data)
                    for figure, data in zip(figures, originals) if data is not None}
    if len({name for name, _ in replacements.values()}) != len(replacements):
        raise ValueError('Conflicting chart EMF paths')
    manifest = 'Contents/content.hpf'
    with NamedTemporaryFile(dir=target.parent, suffix='.tmp', delete=False) as stream:
        temporary = Path(stream.name)
    try:
        with ZipFile(source) as archive, ZipFile(temporary, 'w') as output:
            xml = ET.fromstring(archive.read(manifest))
            matched = set()
            for item in xml.findall('./{*}manifest/{*}item'):
                href = item.get('href')
                if href in replacements:
                    name, data = replacements[href]
                    if href in matched or (name != href and name in archive.namelist()):
                        raise ValueError('Chart media path is not unique')
                    matched.add(href)
                    item.set('href', name)
                    item.set('media-type', 'image/emf')
                    if 'hashkey' in item.attrib:
                        item.set('hashkey', b64encode(md5(data, usedforsecurity=False).digest()).decode('ascii'))
            if matched != set(replacements):
                raise ValueError('Missing chart manifest items')
            for entry in archive.infolist():
                if entry.filename in replacements:
                    name, data = replacements[entry.filename]
                    output.writestr(name, data, compress_type=entry.compress_type)
                elif entry.filename == manifest:
                    output.writestr(entry, ET.tostring(xml, encoding='utf-8', xml_declaration=True))
                else:
                    output.writestr(entry, archive.read(entry))
        figures = verify_chart_document(temporary, prefix=prefix, expected=expected)
        temporary.rename(target)
        return figures
    finally:
        temporary.unlink(missing_ok=True)
