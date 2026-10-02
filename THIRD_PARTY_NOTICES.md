# Third-Party Notices

ReportAutomation's custom license applies to this project's own code, not to
the independently licensed components bundled with the launcher. Preserve the
following notices when redistributing the EXE or its extracted contents.

## Bundled Components

The build records exact installed versions in `bundle_manifest.json`.

- CPython and its standard library: original `runtime/LICENSE.txt`, including
  Microsoft Distributable Code notices, and
  `licenses/cpython-3.12.14-incorporated-license.rst` from the
  [CPython release documentation](https://github.com/python/cpython/blob/v3.12.14/Doc/license.rst).
  The runtime is
  repackaged without modification to interpreter source; unused development,
  testing, pip and Tcl/Tk components are omitted.
- openpyxl 3.1.5: MIT, `licenses/openpyxl-3.1.5-LICENCE.rst`.
- et_xmlfile 2.0.0: MIT and incorporated Python code notices,
  `licenses/et_xmlfile-2.0.0-LICENCE.rst` and
  `licenses/et_xmlfile-2.0.0-LICENCE.python`.
- python-pptx, Pillow, pywin32 and their required dependencies: original
  distribution metadata and license files under `runtime/Lib/site-packages/`.
  Pillow and lxml notices also describe their bundled native components.

Additional runtime-native notices are preserved in `licenses/`: OpenSSL 3.5.8
(Apache-2.0), Expat 2.8.3 (MIT), zlib 1.3.2, libffi (MIT), bzip2 1.0.8,
SQLite (public domain), and XZ/liblzma (the upstream COPYING notices for both
the public-domain and newer 0BSD licensing). XZ tools are not shipped.
The sources of these notices are the corresponding upstream release repositories:
[OpenSSL](https://github.com/openssl/openssl/blob/openssl-3.5.8/LICENSE.txt),
[Expat](https://github.com/libexpat/libexpat/blob/R_2_8_3/expat/COPYING),
[zlib](https://github.com/madler/zlib/blob/v1.3.2/LICENSE),
[libffi](https://github.com/libffi/libffi/blob/v3.4.4/LICENSE),
[bzip2](https://github.com/python/cpython-source-deps/blob/bzip2-1.0.8/LICENSE),
[SQLite](https://github.com/sqlite/sqlite/blob/version-3.53.1/src/main.c), and
[XZ](https://github.com/tukaani-project/xz/blob/v5.8.1/COPYING).

The openpyxl and et_xmlfile notices were copied from the official, versioned
PyPI source archives because their installed wheel metadata lacks the full
license text. Archive SHA256 values were verified before extracting notices:

| Component | Source | SHA256 |
| --- | --- | --- |
| openpyxl 3.1.5 | [PyPI](https://pypi.org/project/openpyxl/3.1.5/) | `cf0e3cf56142039133628b5acffe8ef0c12bc902d2aadd3e0fe5878dc08d1050` |
| et_xmlfile 2.0.0 | [PyPI](https://pypi.org/project/et-xmlfile/2.0.0/) | `dab3f4764309081ce75662649be815c4c9081e88f0837825f90fd28317d4da54` |

If dependency versions change, refresh and verify the corresponding license
notices before distribution. Do not strip upstream copyright or license files.
Microsoft Distributable Code remains subject to the restrictions recorded in
the runtime's original license notice; those restrictions are not superseded
by this project's custom license.

Excel, Hancom HWP, PowerPoint, Windows and system fonts are not redistributed.
They remain separately installed software subject to their respective terms.
No rhwp code or dependency is bundled by the standalone launcher build.
