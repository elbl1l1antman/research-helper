"""Check the saved structure of the experimental HWPX chart containers."""

import re
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile


def verify_chart_document(path: Path, *, prefix: str, expected: tuple[bool, ...]) -> list[dict]:
    with ZipFile(path) as archive:
        manifest = ET.fromstring(archive.read('Contents/content.hpf'))
        media = {item.get('id'): item.get('href') for item in manifest.findall('./{*}manifest/{*}item')}
        sections = sorted((n for n in archive.namelist() if re.fullmatch(r'Contents/section\d+\.xml', n)),
                          key=lambda n: int(re.search(r'(\d+)\.xml$', n)[1]))
        tables = [table for section in sections for table in ET.fromstring(archive.read(section)).findall('.//{*}tbl')
                  if table.get('numberingType') == 'PICTURE']
        if len(tables) != len(expected):
            raise ValueError('Unexpected figure container count')
        return [_verify_container(archive, media, table, prefix, index, has_chart)
                for index, (table, has_chart) in enumerate(zip(tables, expected), 1)]


def _verify_container(archive, media, table, prefix, index, has_chart):
    if (table.get('rowCnt'), table.get('colCnt')) != ('1', '1') or table.find('./{*}pos').get('treatAsChar') != '1':
        raise ValueError(f'Figure {index}: expected inline 1x1 table')
    caption = table.find('./{*}caption')
    if caption is None or caption.get('side') != 'TOP':
        raise ValueError(f'Figure {index}: expected top caption')
    fields = caption.findall('.//{*}autoNum')
    if [(node.get('numType'), node.get('num')) for node in fields] != [('PICTURE', str(index))]:
        raise ValueError(f'Figure {index}: invalid picture number sequence')
    text = ''.join(node.text or '' for node in caption.findall('.//{*}t'))
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
        if len(header) < 88 or header[:4] != b'\x01\0\0\0' or header[40:44] != b' EMF':
            raise ValueError(f'Figure {index}: embedded media is not EMF')
        margins = cell.find('./{*}cellMargin')
        available = int(cell.find('./{*}cellSz').get('width')) - sum(int(margins.get(side, '0')) for side in ('left', 'right'))
        if not 0 < int(picture.find('./{*}sz').get('width')) <= available:
            raise ValueError(f'Figure {index}: picture exceeds cell width')
    return {'number': index, 'title': text[len(leader):], 'emf': target}
