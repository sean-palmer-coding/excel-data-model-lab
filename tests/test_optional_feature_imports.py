"""The headless core must not initialize optional workbook/query engines."""
import subprocess
import sys


def test_core_import_does_not_initialize_optional_engines():
    subprocess.run([sys.executable, '-c', '''
import sys
import excel_model_lab as lab
assert lab.compose_workbook and lab.DaxEvaluator and lab.AdomdQueryClient
assert lab.DesktopSession and lab.AgentWorkflow
for prefix in ('xlsxwriter', 'pbix_mcp', 'clr', 'pythonnet', 'xlwings'):
    assert not any(name == prefix or name.startswith(prefix + '.') for name in sys.modules), prefix
'''], check=True)
