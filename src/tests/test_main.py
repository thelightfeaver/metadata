from src.metadata import Metadata

import pytest


def test_read_doc():
    URL = "src/data/file.docx"
    STRUCTURES = [
        "author",
        "category",
        "comments",
        "content_status",
        "created",
        "identifier",
        "keywords",
        "language",
        "last_modified_by",
        "last_printed",
        "modified",
        "revision",
        "subject",
        "title",
        "version",
    ]
    mt = Metadata(URL)
    results = mt.read_docx_metadata()
    assert len(results.items()) > 0
    assert len(STRUCTURES) > 0 and len(results.keys()) > 0


def test_write_doc():
    URL = "src/data/file.docx"
    mt = Metadata(URL)

    data = {"title": "algo", "author": "author"}

    mt.write_docx_metadata(data)

    results = mt.read_docx_metadata()

    assert results["title"] == "algo"

    assert results["author"] == "author"

    assert len(results.items()) > 0
