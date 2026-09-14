"""Application-level regressions using only synthetic data."""
import io
from pathlib import Path

import pandas as pd
import pytest
from streamlit.testing.v1 import AppTest

import app

ROOT = Path(__file__).resolve().parents[1]


def demo_frame():
    source = app.read_source(app.demo_csv(), 'synthetic-leads.csv')
    table = app.table_from_source(source, source.header_row)
    return app.mark_duplicates(app.clean_leads(table, app.guess_mapping(table), 'synthetic-leads.csv', vip_threshold=1_000_000))


def assert_no_exceptions(at):
    assert len(at.exception) == 0, [item.message for item in at.exception]


def test_demo_ingestion_and_filters():
    frame = demo_frame()
    assert len(frame) == 12
    assert frame['phone_valid'].sum() == 10
    assert frame['duplicate'].sum() == 2
    assert len(app.filter_leads(frame, query='القاهرة')) == 3
    assert len(app.filter_leads(frame, flag='duplicate')) == 2
    assert app.filter_leads(frame, contact='unreachable')['name'].tolist() == ['عميل تجريبي ٠٧']
    assert app.filter_leads(frame, query='x.*(missing)').empty
    amounts = app.filter_leads(frame, currency='EGP', minimum_budget=2_000_000)
    assert len(amounts) == 3
    assert amounts['budget_currency'].eq('EGP').all()


def test_excel_summary_cached_values_and_preserved_formulas():
    import openpyxl
    frame = demo_frame()
    data = app.build_excel(frame, 'اختبار الملخص')
    cached = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    formulas = openpyxl.load_workbook(io.BytesIO(data), data_only=False)
    summary = cached['الملخص']
    assert summary['B5'].value == 12
    assert summary['B6'].value == 10
    assert summary['B7'].value == 4
    assert summary['B8'].value == 4
    assert summary['B9'].value == 4
    assert summary['B10'].value == pytest.approx(4 / 12)
    assert summary['B11'].value == 11
    assert summary['B12'].value == 2
    assert summary['B13'].value == 'الهاتف'
    assert formulas['الملخص']['B10'].value == '=IFERROR(B9/B5,0)'
    assert not any(cell.data_type == 'e' for sheet in cached for row in sheet for cell in row)
    cached.close()
    formulas.close()


def test_demo_no_live_contact_actions():
    html = app._lead_table_html(demo_frame(), allow_contacts=False)
    assert 'href="tel:' not in html
    assert 'href="https://wa.me/' not in html
    assert '<details>' in html


def test_web_escape_and_contact_actions():
    frame = demo_frame().iloc[:1].copy()
    frame.loc[frame.index[0], 'name'] = '<img src=x onerror=alert(1)>'
    html = app._lead_table_html(frame)
    assert '<img src=x' not in html
    assert '&lt;img' in html
    assert 'href="tel:+201000000001"' in html


@pytest.mark.parametrize('claims,allow,expected', [
    ({'email': 'allowed@example.test', 'email_verified': True}, ['allowed@example.test'], True),
    ({'email': ' ALLOWED@example.test ', 'email_verified': True}, ['allowed@example.test'], True),
    ({'email': 'allowed@example.test', 'email_verified': 'true'}, ['allowed@example.test'], False),
    ({'email': 'allowed@example.test'}, ['allowed@example.test'], False),
    ({'email': 'outside@example.test', 'email_verified': True}, ['allowed@example.test'], False),
    ({'email': 'allowed@example.test', 'email_verified': True}, [], False),
    ({'email': 'allowed@example.test', 'email_verified': True, 'exp': 9}, ['allowed@example.test'], False),
    ({'email': 'allowed@example.test', 'email_verified': True, 'exp': 'nan'}, ['allowed@example.test'], False),
    ({'email': 'allowed@example.test', 'email_verified': True, 'exp': 100}, ['allowed@example.test'], True),
])
def test_oidc_authorization(claims, allow, expected):
    assert app.auth_access_allowed(claims, allow, now=10) is expected


def test_oidc_missing_configuration_fails_closed(monkeypatch):
    monkeypatch.setenv('AUTH_MODE', 'oidc')
    monkeypatch.delenv('ALLOWED_EMAILS', raising=False)
    at = AppTest.from_file(str(ROOT / 'app.py'), default_timeout=30).run()
    assert_no_exceptions(at)
    assert len(at.error) == 1
    assert len(at.get('file_uploader')) == 0
    assert len(at.radio) == 0


def test_demo_reruns_exports_filters_and_clear(monkeypatch):
    monkeypatch.setenv('AUTH_MODE', 'open')
    at = AppTest.from_file(str(ROOT / 'app.py'), default_timeout=40).run()
    assert_no_exceptions(at)
    at.radio(key='lc_mode').set_value('تجربة جاهزة').run()
    assert_no_exceptions(at)
    assert len(at.error) == 0
    at.button(key='lc_process').click().run()
    assert_no_exceptions(at)
    assert len(at.session_state['lc_result']) == 12
    assert len(at.error) == 0
    # Parse cache must survive class redefinition on a rerun.
    at.selectbox(key='lc_theme').set_value('داكن').run()
    assert_no_exceptions(at)
    assert len(at.error) == 0
    assert len(at.session_state['lc_result']) == 12
    at.text_input(key='lc_search').set_value('لا يوجد بهذا الاسم').run()
    assert_no_exceptions(at)
    at.text_input(key='lc_search').set_value('').run()
    at.button(key='lc_build_excel').click().run()
    assert_no_exceptions(at)
    assert at.session_state['lc_excel_bytes'].startswith(b'PK')
    at.button(key='lc_build_pdf').click().run()
    assert_no_exceptions(at)
    assert len(at.error) == 0, [item.value for item in at.error]
    assert at.session_state['lc_pdf_bytes'].startswith(b'%PDF-')
    at.selectbox(key='lc_region').set_value('SA').run()
    assert_no_exceptions(at)
    assert 'lc_result' not in at.session_state
    clear = [button for button in at.button if button.label == 'مسح بيانات الجلسة'][0]
    clear.click().run()
    assert_no_exceptions(at)
    assert 'lc_result' not in at.session_state
    assert 'lc_parse_cache' not in at.session_state


def test_independent_sessions(monkeypatch):
    monkeypatch.setenv('AUTH_MODE', 'open')
    first = AppTest.from_file(str(ROOT / 'app.py'), default_timeout=30).run()
    first.radio(key='lc_mode').set_value('تجربة جاهزة').run()
    first.button(key='lc_process').click().run()
    second = AppTest.from_file(str(ROOT / 'app.py'), default_timeout=30).run()
    assert 'lc_result' in first.session_state
    assert 'lc_result' not in second.session_state
    assert 'lc_parse_cache' not in second.session_state
