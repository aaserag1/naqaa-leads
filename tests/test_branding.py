"""Regressions for the approved platform name and independent report branding."""
import base64
import io
from pathlib import Path

import openpyxl
from PIL import Image
from pypdf import PdfReader
import pytest
from streamlit.testing.v1 import AppTest

import app

ROOT = Path(__file__).resolve().parents[1]


def fixture_frame():
    source = app.read_source(app.demo_csv(), 'synthetic.csv')
    frame = app.table_from_source(source, source.header_row)
    return app.clean_leads(frame, app.guess_mapping(frame), 'synthetic.csv').iloc[:1].copy()


def logo_bytes(color):
    output = io.BytesIO()
    Image.new('RGB', (90, 36), color).save(output, format='PNG')
    return output.getvalue()


def test_canonical_platform_identity():
    assert app.PRODUCT_NAME == 'Naqaa | نقاء'
    assert app.PRODUCT_SLUG == 'naqaa'
    assert app.APP_TITLE == app.PRODUCT_NAME


def test_company_preview_does_not_rename_platform(monkeypatch):
    monkeypatch.setenv('AUTH_MODE', 'open')
    at = AppTest.from_file(str(ROOT / 'app.py'), default_timeout=30).run()
    at.text_input(key='lc_brand').set_value('ATLASAGENCY').run()
    assert not at.exception
    header = next(item.value for item in at.markdown if 'class="lc-top"' in item.value)
    assert '<bdi class="lc-product" dir="ltr">Naqaa | نقاء</bdi>' in header
    assert '<div class="lc-brand"><div>ATLASAGENCY</div></div>' in header


@pytest.mark.parametrize('company,color', [('ATLASAGENCY','#0077cc'), ('وكالة ثانية للاختبار','#b13c70')])
def test_excel_keeps_custom_company_name_and_logo(company, color):
    data = app.build_excel(fixture_frame(), company, logo_bytes(color))
    book = openpyxl.load_workbook(io.BytesIO(data))
    assert book.properties.creator == app.PRODUCT_NAME
    for sheet in book:
        assert sheet['A1'].value == company
    images = book['الملخص']._images
    assert len(images) == 1
    with Image.open(io.BytesIO(images[0]._data())) as image:
        assert image.convert('RGB').getpixel((0,0)) == Image.new('RGB',(1,1),color).getpixel((0,0))
    book.close()


def test_pdf_keeps_company_title_and_custom_logo():
    company = 'ATLASAGENCY'
    raw_logo = logo_bytes('#0077cc')
    html = app.build_report_html(fixture_frame(), company, raw_logo)
    assert '<h1>ATLASAGENCY</h1>' in html
    assert '<h1>Naqaa' not in html
    assert 'data:image/png;base64,' + base64.b64encode(app.sanitize_logo(raw_logo)).decode() in html
    pdf = PdfReader(io.BytesIO(app.build_pdf(fixture_frame(), company, raw_logo)))
    assert pdf.metadata.title.startswith(company)
    assert 'ATLASAGENCY' in ''.join(page.extract_text() for page in pdf.pages)


def test_blank_report_brand_uses_neutral_title_not_product_name():
    report = app.build_report_html(fixture_frame(), '')
    assert '<h1>تقرير العملاء</h1>' in report
    assert '<h1>Naqaa' not in report
    book = openpyxl.load_workbook(io.BytesIO(app.build_excel(fixture_frame(), '')))
    assert book['الملخص']['A1'].value == 'تقرير العملاء'
    book.close()


def test_project_identity_is_consistent():
    for filename in ('README.md', 'START_HERE.md', 'DEPLOYMENT.md', 'BRANDING.md'):
        assert app.PRODUCT_NAME in (ROOT / filename).read_text(encoding='utf-8')
    assert 'name: naqaa' in (ROOT / 'render.yaml').read_text()
    assert 'name: Naqaa CI' in (ROOT / '.github/workflows/ci.yml').read_text()
    assert 'org.opencontainers.image.title="Naqaa | نقاء"' in (ROOT / 'Dockerfile').read_text()
    for filename in ('README.md', 'DEPLOYMENT.md', 'render.yaml', '.github/workflows/ci.yml'):
        content = (ROOT / filename).read_text()
        assert 'arabic-leads-cleaner' not in content
        assert 'leads-cleaner:local' not in content
        assert 'leads-cleaner:ci' not in content
