"""Create a synthetic workbook and edit it through the Windows desktop bridge.

Requires Windows Excel and the workbooks,desktop extras. Refuses to overwrite
an existing destination. Run: python examples/desktop_agent_workflow.py demo.xlsx
"""
from pathlib import Path
import json
import sys
import time
import pandas as pd
from excel_model_lab import (compose_workbook, WorkbookModel, DesktopSession,
                             AgentWorkflow, AgentOperation)


def run(destination):
    destination = Path(destination).resolve()
    if destination.exists():
        raise FileExistsError(destination)

    def sheets(writer, cubes):
        sheet = writer.add_worksheet('Summary')
        sheet.write('A1', 'Total Amount')
        cubes.write(sheet, 'B1', cubes.value(cubes.measure('Total Amount')))

    compose_workbook(destination, {'Sales': pd.DataFrame({'Amount': [10, 20, 30]})},
                     sheets, configure_model=lambda m: m.add_measure(
                         'Sales', 'Total Amount', 'SUM(Sales[Amount])'))
    session = DesktopSession.open(destination, visible=True)
    workflow = AgentWorkflow(session)
    # Planning binds the exact saved file and process. Do not edit Excel
    # interactively while an operation batch or file handoff is running.
    session.save()  # Persist initial Excel calculation before creating a plan.
    plan = workflow.plan([
        AgentOperation.create('write_range', sheet='Summary', address='A3',
                              values=[['Updated through the Python API']]),
        AgentOperation.create('save'),
    ], description='Add a worksheet note')
    workflow.apply(plan)

    def edit(source, output):
        model = WorkbookModel.open(source, read_only=False)
        model.replace_column('Sales', 'Amount', [20, 30, 40])
        model.save(output)

    handoff = workflow.plan_edit(description='Replace the three model amounts')
    result = workflow.apply_edit(handoff, edit)
    session.calculate()
    deadline = time.monotonic() + 30
    while session.read_range("Summary", "B1") != [[90.0]]:
        if time.monotonic() >= deadline:
            raise RuntimeError("Excel did not produce the expected CUBE total within 30 seconds")
        time.sleep(0.2)
    session.save()
    print(json.dumps({'workbook': str(destination),
                      'checkpoint': str(result.checkpoint.workbook),
                      'expected_total_after_calculation': 90}, indent=2))
    session.close()
    session.quit()
    # Exceptions deliberately leave the workbook/recovery files available;
    # no finally block discards an unknown unsaved state.


if __name__ == '__main__':
    run(sys.argv[1] if len(sys.argv) > 1 else 'desktop-workflow.xlsx')
