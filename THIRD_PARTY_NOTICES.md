# Third-party notices

## Optional workbook and evaluation dependencies

The `workbooks` extra installs XlsxWriter separately for worksheet authoring.
The `dax` extra installs the MIT-licensed pbix-mcp 0.9.153 distribution separately;
the adapter uses its Python DAX evaluator through a bounded, independently tested
interface. No upstream evaluator source is copied into this package. See
`docs/phase6-evaluation.md` for the audited wheel identity and supported scope.
The `external-query` extra installs pythonnet separately. Microsoft ADOMD.NET
and Analysis Services are not bundled or installed automatically by this library.
Each optional dependency and its transitive dependencies retain their own licenses.

Projects: https://github.com/jmcnamara/XlsxWriter,
https://github.com/d0nk3yhm/pbix-mcp, https://github.com/pythonnet/pythonnet.

## PBIXRay

The package depends on PBIXRay 0.15.5 by Igor Cotruta. Its reader and format
research underpin this project. PBIXRay is installed separately as a dependency.
Its license is reproduced below for attribution; it does not select the license
for this project's original code or establish rights to redistribute fixtures.

MIT License

Copyright (c) 2024 Igor Cotruta

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.


## PivotWright format research

The native source encryption and source-binding format work was cross-checked
against PivotWright by Matthew Moore, in addition to Microsoft MS-XLDM
specifications and native Excel files. PivotWright is not a runtime dependency.
Repository: https://github.com/MatthewMooreZA/PivotWright

MIT License

Copyright (c) 2026 Matthew Moore

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

## Optional desktop adapter

The `desktop` extra installs xlwings 0.33.x separately (open-source core, BSD
3-Clause license). The adapter uses the ordinary desktop automation API, not
xlwings PRO features. Microsoft Excel is separately installed and licensed;
neither Excel nor xlwings binaries are bundled in this distribution.
Project: https://github.com/xlwings/xlwings
License: https://github.com/xlwings/xlwings/blob/main/LICENSE.txt
