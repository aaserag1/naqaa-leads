"""Verify the repository's lossless, network-free font transport."""
import base64
import hashlib
import io
from pathlib import Path

from pypdf import PdfReader
import pytest

import app

ROOT = Path(__file__).resolve().parents[1]
EXPECTED = '667c987182391c91f4e57a2f455b1794fb5e3ee6ca4ef3383e86bb690fa9c964'


def encoded_font():
    return (ROOT / 'assets/Cairo.ttf.b64').read_bytes()


def isolate_font(monkeypatch, tmp_path, data):
    (tmp_path / 'assets').mkdir()
    (tmp_path / 'assets/Cairo.ttf.b64').write_bytes(data)
    monkeypatch.setattr(app, '__file__', str(tmp_path / 'app.py'))


def test_base64_is_exact_original_font():
    decoded = base64.b64decode(b''.join(encoded_font().split()), validate=True)
    assert hashlib.sha256(decoded).hexdigest() == EXPECTED
    assert app._export_font_bytes() == decoded


def test_loads_without_binary_file_or_disk_materialization(monkeypatch, tmp_path):
    isolate_font(monkeypatch, tmp_path, encoded_font())
    assert hashlib.sha256(app._export_font_bytes()).hexdigest() == EXPECTED
    assert not (tmp_path / 'assets/Cairo.ttf').exists()


@pytest.mark.parametrize('bad', [b'@@ invalid @@', base64.b64encode(b'\x00\x01\x00\x00' + b'x' * 200), b''])
def test_modified_or_invalid_font_is_rejected(monkeypatch, tmp_path, bad):
    isolate_font(monkeypatch, tmp_path, bad)
    with pytest.raises(app.UserDataError):
        app._export_font_bytes()


def test_pdf_worker_in_text_only_repository(monkeypatch, tmp_path):
    (tmp_path / 'assets').mkdir()
    (tmp_path / 'assets/Cairo.ttf.b64').write_bytes(encoded_font())
    (tmp_path / 'app.py').write_bytes((ROOT / 'app.py').read_bytes())
    assert not (tmp_path / 'assets/Cairo.ttf').exists()
    monkeypatch.setattr(app, '__file__', str(tmp_path / 'app.py'))
    source = app.read_source(app.demo_csv(), 'synthetic.csv')
    frame = app.table_from_source(source, source.header_row)
    frame = app.clean_leads(frame, app.guess_mapping(frame), 'synthetic.csv').iloc[:1]
    pdf = PdfReader(io.BytesIO(app.build_pdf(frame, 'FONTTEST')))
    assert pdf.metadata.title.startswith('FONTTEST')
    assert not (tmp_path / 'assets/Cairo.ttf').exists()
