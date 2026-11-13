from struct import pack_into
from zipfile import ZipFile

import pytest

from report_automation_engine.hwp_chart_probe import preserve_chart_emfs, verify_chart_document


@pytest.mark.parametrize('change,match', [('', ''), ('raster', 'EMF'), ('wmf', 'EMF'), ('double', 'NONE'),
    ('overflow', 'width'), ('wrong-number', 'sequence'), ('wrong-prefix', 'caption'), ('floating', 'inline')])
def test_saved_chart_container(tmp_path, change, match):
    path = tmp_path / 'probe.hwpx'
    emf = bytearray(108)
    pack_into('<II', emf, 0, 1, 88)
    emf[40:44] = b' EMF'
    number = 3 if change == 'wrong-number' else 2
    prefix = '2' if change == 'wrong-prefix' else '3'
    inside = 'PICTURE' if change == 'double' else 'NONE'
    width = 10001 if change == 'overflow' else 9000
    inline = '0' if change == 'floating' else '1'
    with ZipFile(path, 'w') as archive:
        suffix = 'wmf' if change == 'wmf' else 'emf'
        archive.writestr('Contents/content.hpf', f'<package><manifest><item id="image1" href="BinData/image1.{suffix}" media-type="image/{suffix}"/></manifest></package>')
        media = {'raster': b'PNG', 'wmf': b'\x01\0\x09\0\0\x03' + bytes(102)}.get(change, emf)
        archive.writestr(f'BinData/image1.{suffix}', media)
        archive.writestr('Contents/section0.xml', f'''<sec><p><run>
          <tbl rowCnt="1" colCnt="1" numberingType="PICTURE"><pos treatAsChar="{inline}"/>
            <caption side="TOP"><p><t>&lt;그림 {prefix}-</t><autoNum num="1" numType="PICTURE"/><t>&gt; 첫 그림</t></p></caption>
            <tr><tc><cellSz width="10200"/><cellMargin left="100" right="100"/><subList>
              <p><run><pic numberingType="{inside}"><sz width="{width}"/><img binaryItemIDRef="image1"/></pic></run></p>
            </subList></tc></tr>
          </tbl></run></p><p><run>
          <tbl rowCnt="1" colCnt="1" numberingType="PICTURE"><pos treatAsChar="1"/>
            <caption side="TOP"><p><t>&lt;그림 3-</t><autoNum num="{number}" numType="PICTURE"/><t>&gt; 보류 그림</t></p></caption>
            <tr><tc><subList><p><run><t/></run></p></subList></tc></tr>
          </tbl></run></p></sec>''')
    if match:
        with pytest.raises(ValueError, match=match):
            verify_chart_document(path, prefix='3', expected=(True, False))
        if change == 'wmf':
            target = tmp_path / 'preserved.hwpx'
            original = path.read_bytes()
            with pytest.raises(ValueError, match='EMF'):
                preserve_chart_emfs(path, target, prefix='3', originals=(b'PNG', None))
            assert not target.exists()
            preserve_chart_emfs(path, target, prefix='3', originals=(bytes(emf), None))
            assert path.read_bytes() == original
            assert verify_chart_document(target, prefix='3', expected=(True, False))[0]['emf'] == 'BinData/image1.emf'
            with ZipFile(target) as archive:
                assert archive.read('BinData/image1.emf') == emf
                assert 'BinData/image1.wmf' not in archive.namelist()
    else:
        assert verify_chart_document(path, prefix='3', expected=(True, False)) == [
            {'number': 1, 'title': '첫 그림', 'emf': 'BinData/image1.emf'},
            {'number': 2, 'title': '보류 그림', 'emf': None},
        ]
