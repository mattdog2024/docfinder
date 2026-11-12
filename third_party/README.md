# Vendored dependencies

## olefile 0.47

- Upstream: https://github.com/decalage2/olefile
- Official release: https://pypi.org/project/olefile/0.47/
- Source archive: `olefile-0.47.zip` (PyPI sdist).
- Archive SHA-256: `599383381a0bf3dfbd932ca0ca6515acd174ed48870cbf7fee123d698c192c1c`.
- Included unmodified: `olefile/__init__.py`, `olefile/olefile.py`, and `olefile/LICENSE.txt`.
- License: upstream BSD-style terms, with the original PIL license retained in full in `olefile/LICENSE.txt`.

This is a dependency-free, pure-Python OLE container reader. It is not itself a Word text parser. `core/extractor.py` statically imports `third_party.olefile`; PyInstaller follows this ordinary import without pip-installing olefile, hidden imports, runtime downloads, or changes to the build workflow.

The full upstream license is additionally embedded as `third_party.THIRD_PARTY_NOTICES`, so the frozen Python module retains the notice even when the existing build does not collect extra text files. Binary redistribution must also reproduce this license in the documentation or other supporting materials shipped with the release; include `olefile/LICENSE.txt` or its complete contents with release documentation. No license terms have been removed or replaced.

## DOC extraction scope

The in-project parser reads the Word 97+ FIB, the selected `0Table`/`1Table` CLX, and each piece's actual WordDocument offset. It returns main-story text (including table-cell text) in character-position order. It does not reconstruct table layout or formatting and does not claim complete support for every legacy Office format. Encrypted, pre-Word-97, malformed, and non-OLE `.doc` inputs fail closed to empty text; filename indexing remains the caller's responsibility.

Format reference: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/01d5d8c4-cf9c-4ef9-80fd-439e763cfe01
