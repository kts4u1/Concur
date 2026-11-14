# -*- coding: utf-8 -*-
"""모든 테스트 공통
- 자동 백업이 실제 '문서' 폴더가 아니라 테스트용 임시 폴더에 생기게 합니다.
- 테스트가 끝나면 닫은 창의 남은 조각을 바로 치웁니다(다음 테스트의 작업 스레드가 치우다 멈추지 않게)."""

import gc

import pytest


@pytest.fixture(autouse=True)
def backup_folder(tmp_path_factory, monkeypatch):
    folder = tmp_path_factory.mktemp("backups")
    monkeypatch.setenv("CONCUR_BACKUP_DIR", str(folder))
    yield folder
    gc.collect()
