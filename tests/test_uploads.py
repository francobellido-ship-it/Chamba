"""Temporary inspection reuse: isolation, bounded disk, expiry and publication."""

import os
import stat

from fastapi.testclient import TestClient
import pymupdf

from app.main import create_app
from app.uploads import InspectionCache
from app.documents import prepare_document
from app.pdf import CertificateError
import pytest
from test_excel import workbook_bytes


def test_private_copies_have_distinct_tokens_and_are_deleted():
    cache = InspectionCache()
    first = cache.put('one.xlsx', b'first workbook')
    second = cache.put('two.xlsx', b'second workbook')
    assert first != second
    assert cache.get(first).data == b'first workbook'
    assert cache.get(second).filename == 'two.xlsx'
    temporary = cache.entries[first].file
    assert stat.S_IMODE(os.fstat(temporary.fileno()).st_mode) == 0o600
    cache.discard(first)
    assert temporary.closed and cache.get(first) is None
    assert cache.get('../' + second) is None
    cache.close()
    assert not cache.entries


def test_capacity_and_expiry_release_disk_and_reject_old_tokens(monkeypatch):
    clock = [100.0]
    monkeypatch.setattr('app.uploads.time.monotonic', lambda: clock[0])
    cache = InspectionCache(max_bytes=10, max_files=2, ttl=60)
    first = cache.put('first.xlsx', b'123456')
    temporary = cache.entries[first].file
    second = cache.put('second.xlsx', b'12345')
    assert cache.get(first) is None and temporary.closed
    assert cache.get(second).data == b'12345'
    assert cache.put('too-big.xlsx', b'0' * 11) is None
    clock[0] += 61
    cache.expire()
    assert cache.get(second) is None and not cache.entries
    cache.close()


def test_unavailable_temporary_disk_falls_back_to_normal_upload(monkeypatch):
    def full_disk(**kwargs):
        raise OSError('synthetic disk full')
    monkeypatch.setattr('app.uploads.tempfile.TemporaryFile', full_disk)
    cache = InspectionCache()
    assert cache.put('example.xlsx', b'workbook') is None
    cache.close()


def test_cached_workbook_uses_the_original_name_and_normal_conversion(tmp_path):
    with TestClient(create_app(tmp_path, 'http://127.0.0.1:8000')) as client:
        inspection = client.post('/api/excel/inspect', files={
            'excel': ('QCP-2026-0007.xlsx', workbook_bytes())})
        assert inspection.status_code == 200
        token = inspection.json()['uploadToken']
        assert inspection.headers['cache-control'] == 'no-store'
        fields = {'excelToken': token, 'tipoCertificado': 'no_acreditado'}
        invalid = client.post('/api/certificates', data={**fields, 'codigoCorrecto': 'invalid'})
        assert invalid.status_code == 422
        # A validation error leaves the copy available for a corrected retry.
        response = client.post('/api/certificates', data=fields)
        assert response.status_code == 201, response.text
        info = response.json()
        assert info['filename'] == 'QCP-2026-0007.pdf'
        assert info['finalCode'] == 'QCP-2026-0001'
        assert any('nombre del archivo' in warning for warning in info['warnings'])
        assert info['signatureIncluded'] and info['backgroundVerified']
        with pymupdf.open(stream=client.get(info['downloadUrl']).content, filetype='pdf') as doc:
            assert len(doc) == 2
            assert '17,5' in doc[0].get_text()
            assert info['verificationUrl'] in [link.get('uri') for link in doc[0].get_links()]
        assert client.post('/api/certificates', data=fields).status_code == 410


def test_unknown_or_conflicting_tokens_never_publish(tmp_path):
    with TestClient(create_app(tmp_path)) as client:
        assert client.post('/api/certificates', data={'excelToken': 'x' * 43}).status_code == 410
        assert client.post('/api/certificates', data={'excelToken': '../../etc/passwd'}).status_code == 410
        response = client.post('/api/certificates', data={'excelToken': 'x' * 43},
                               files={'excel': ('test.xlsx', b'not an Excel')})
        assert response.status_code == 422
        assert not list(tmp_path.iterdir())


def test_invalid_code_is_rejected_before_office_conversion(monkeypatch):
    def unexpected_conversion(*args, **kwargs):
        pytest.fail('An invalid code must not start LibreOffice')
    monkeypatch.setattr('app.documents.convert_excel', unexpected_conversion)
    with pytest.raises(CertificateError, match='formato QCP'):
        prepare_document(workbook_bytes(), 'xlsx', 'invalid', 'https://example.com')
