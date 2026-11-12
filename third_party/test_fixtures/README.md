# Legacy DOC regression fixture

`mixed_table.doc` is a real OLE/CFB Word 97 DOC generated locally by LibreOffice 24.2.7.2 from the accompanying UTF-8 `mixed_table.html`. The source was authored specifically for this test: Chinese/English paragraphs, repeated paragraphs, a two-column three-row table with Chinese/English cells and numeric values, and accented/punctuation text. It contains no user document or private data. The fixture HTML and DOC are dedicated to the public domain under CC0 1.0: https://creativecommons.org/publicdomain/zero/1.0/ .

Reproduce (with LibreOffice installed):

```sh
libreoffice -env:UserInstallation=file:///tmp/docfinder-legacy-fixture-profile \
  --headless --convert-to 'doc:MS Word 97' --outdir /tmp \
  third_party/test_fixtures/mixed_table.html
```

The checked-in DOC SHA-256 is `a95ce63febb17493653af824c02de72e14c0e2bb9431f9876986b747ffe61860`. Regeneration can differ in metadata and hence binary hash. Tests use the checked-in file and do not require LibreOffice at runtime. This one fixture does not establish compatibility with all Word files or table layouts.

All `synthetic_*` cases in `tests/test_legacy_extract.py` construct FIB/CLX byte streams in memory; those cases are explicitly synthetic and are not represented as real Word-produced DOC samples. Mutated copies of the real DOC exercise fail-closed paths separately.
