from types import SimpleNamespace

import pytest

from app.api import recipients as recipient_api


def test_text_and_csv_parsers_bound_materialized_recipients():
    text = "\n".join(f"user{index}@example.com,User {index}" for index in range(20))

    assert len(recipient_api.parse_txt(text, max_items=5)) == 5
    assert len(recipient_api.parse_csv_content(text, max_items=7)) == 7


def test_xlsx_decompression_bomb_is_rejected_before_pandas(monkeypatch):
    class FakeArchive:
        def __enter__(self):
            return self

        def __exit__(self, *args):
            return False

        def infolist(self):
            return [SimpleNamespace(file_size=101 * 1024 * 1024)]

    monkeypatch.setattr(recipient_api.zipfile, "ZipFile", lambda *args, **kwargs: FakeArchive())
    monkeypatch.setattr(
        recipient_api.pd,
        "read_excel",
        lambda *args, **kwargs: pytest.fail("pandas must not parse a rejected archive"),
    )

    assert recipient_api.parse_excel_content(b"PK\x03\x04payload") == []
