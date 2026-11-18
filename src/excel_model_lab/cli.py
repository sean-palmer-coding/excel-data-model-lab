from __future__ import annotations

import argparse
import json
from pathlib import Path

from .package import inspect_workbook, repack_workbook, semantic_snapshot
from .abf import AbfImage, diff_inventories, replace_model_part
from .vertipaq import patch_single_segment_bitpacked_id
from .mutate import set_integer_value
from .model import ExcelModelTemplate
from .spec import build_from_spec, load_spec


def main() -> None:
    parser = argparse.ArgumentParser(prog="xmodel-lab")
    commands = parser.add_subparsers(dest="command", required=True)

    inspect_cmd = commands.add_parser("inspect")
    inspect_cmd.add_argument("workbook", type=Path)
    inspect_cmd.add_argument("--semantic", action="store_true")

    repack_cmd = commands.add_parser("repack")
    repack_cmd.add_argument("source", type=Path)
    repack_cmd.add_argument("destination", type=Path)

    abf_inspect = commands.add_parser("abf-inventory")
    abf_inspect.add_argument("workbook", type=Path)

    abf_extract = commands.add_parser("abf-extract")
    abf_extract.add_argument("workbook", type=Path)
    abf_extract.add_argument("destination", type=Path)

    abf_clone = commands.add_parser("abf-clone")
    abf_clone.add_argument("source", type=Path)
    abf_clone.add_argument("destination", type=Path)

    abf_diff = commands.add_parser("abf-diff")
    abf_diff.add_argument("left", type=Path)
    abf_diff.add_argument("right", type=Path)

    idf_patch = commands.add_parser("idf-patch")
    idf_patch.add_argument("source", type=Path)
    idf_patch.add_argument("destination", type=Path)
    idf_patch.add_argument("storage_path")
    idf_patch.add_argument("--position", type=int, required=True)
    idf_patch.add_argument("--bit-width", type=int, required=True)
    idf_patch.add_argument("--min-data-id", type=int, required=True)
    idf_patch.add_argument("--new-data-id", type=int, required=True)

    set_int = commands.add_parser("set-int")
    set_int.add_argument("source", type=Path)
    set_int.add_argument("destination", type=Path)
    set_int.add_argument("table")
    set_int.add_argument("column")
    set_int.add_argument("position", type=int)
    set_int.add_argument("value", type=int)

    schema = commands.add_parser("schema")
    schema.add_argument("workbook", type=Path)
    dependencies = commands.add_parser('dependencies', help='Inspect lexical/model dependencies as JSON; not a DAX engine')
    dependencies.add_argument('workbook', type=Path)
    validate = commands.add_parser('validate', help='Run static writer preflight; does not certify Excel loading')
    validate.add_argument('workbook', type=Path)

    replace_csv = commands.add_parser("replace-csv")
    replace_csv.add_argument("source", type=Path)
    replace_csv.add_argument("destination", type=Path)
    replace_csv.add_argument("table")
    replace_csv.add_argument("csv", type=Path)
    replace_csv.add_argument("--columns", nargs="+")

    replace_table = commands.add_parser("replace-table")
    replace_table.add_argument("source", type=Path)
    replace_table.add_argument("destination", type=Path)
    replace_table.add_argument("table")
    replace_table.add_argument("csv", type=Path)

    add_measure = commands.add_parser("add-measure")
    add_measure.add_argument("source", type=Path)
    add_measure.add_argument("destination", type=Path)
    add_measure.add_argument("table")
    add_measure.add_argument("name")
    add_measure.add_argument("expression")
    add_measure.add_argument("--format-string", default="")
    add_measure.add_argument("--display-folder", default="")
    add_measure.add_argument("--description", default="")

    add_relationship = commands.add_parser("add-relationship")
    add_relationship.add_argument("source", type=Path)
    add_relationship.add_argument("destination", type=Path)
    add_relationship.add_argument("from_table")
    add_relationship.add_argument("from_column")
    add_relationship.add_argument("to_table")
    add_relationship.add_argument("to_column")
    add_relationship.add_argument("--inactive", action="store_true")
    add_relationship.add_argument('--cardinality', default='many-to-one', choices=['many-to-one', 'one-to-many', 'one-to-one'])

    build = commands.add_parser("build")
    build.add_argument("template", type=Path)
    build.add_argument("destination", type=Path)
    build.add_argument("spec", type=Path)

    create = commands.add_parser("create", help="Create a Data Model from CSV/Parquet sources")
    create.add_argument("destination", type=Path)
    create.add_argument("spec", type=Path)
    create.add_argument("--workbook", type=Path, help="Insert into an ordinary existing workbook")

    args = parser.parse_args()
    if args.command == "inspect":
        result = inspect_workbook(args.workbook)
        if args.semantic:
            result["semantic"] = semantic_snapshot(args.workbook)
        print(json.dumps(result, indent=2, default=str))
    elif args.command == "repack":
        repack_workbook(args.source, args.destination)
    elif args.command == "abf-inventory":
        print(json.dumps(AbfImage.from_workbook(args.workbook).inventory(), indent=2))
    elif args.command == "abf-extract":
        AbfImage.from_workbook(args.workbook).extract(args.destination)
    elif args.command == "abf-clone":
        image = AbfImage.from_workbook(args.source)
        replace_model_part(args.source, args.destination, image.clone())
    elif args.command == "abf-diff":
        left = AbfImage.from_workbook(args.left)
        right = AbfImage.from_workbook(args.right)
        print(json.dumps(diff_inventories(left, right), indent=2))
    elif args.command == "idf-patch":
        image = AbfImage.from_workbook(args.source)
        logical = image.read_logical_file(args.storage_path)
        patched = patch_single_segment_bitpacked_id(
            logical,
            position=args.position,
            bit_width=args.bit_width,
            min_data_id=args.min_data_id,
            new_data_id=args.new_data_id,
        )
        replace_model_part(
            args.source,
            args.destination,
            image.assemble({args.storage_path: patched}),
        )
    elif args.command == "set-int":
        set_integer_value(
            args.source,
            args.destination,
            table=args.table,
            column=args.column,
            position=args.position,
            value=args.value,
        )
    elif args.command == "schema":
        template = ExcelModelTemplate(args.workbook)
        print(json.dumps([item.__dict__ for item in template.columns], indent=2))
    elif args.command == 'dependencies':
        from dataclasses import asdict
        print(json.dumps(asdict(ExcelModelTemplate(args.workbook).dependency_graph()), indent=2))
    elif args.command == 'validate':
        from dataclasses import asdict
        report = ExcelModelTemplate(args.workbook).validate()
        print(json.dumps({'valid': report.valid, **asdict(report)}, indent=2))
        if not report.valid:
            raise SystemExit(1)
    elif args.command == "replace-csv":
        import pandas as pd

        frame = pd.read_csv(args.csv)
        (
            ExcelModelTemplate(args.source)
            .replace_dataframe(args.table, frame, columns=args.columns)
            .save(args.destination)
        )
    elif args.command == "replace-table":
        import pandas as pd

        (
            ExcelModelTemplate(args.source)
            .replace_table(args.table, pd.read_csv(args.csv))
            .save(args.destination)
        )
    elif args.command == "add-measure":
        (
            ExcelModelTemplate(args.source)
            .add_measure(
                args.table,
                args.name,
                args.expression,
                format_string=args.format_string,
                display_folder=args.display_folder,
                description=args.description,
            )
            .save(args.destination)
        )
    elif args.command == "add-relationship":
        (
            ExcelModelTemplate(args.source)
            .add_relationship(
                args.from_table,
                args.from_column,
                args.to_table,
                args.to_column,
                active=not args.inactive,
                cardinality=args.cardinality,
            )
            .save(args.destination)
        )
    elif args.command == "create":
        build_from_spec(None, args.destination, load_spec(args.spec),
                        base_directory=args.spec.parent, workbook=args.workbook)
    elif args.command == "build":
        build_from_spec(
            args.template,
            args.destination,
            load_spec(args.spec),
            base_directory=args.spec.parent,
        )


if __name__ == "__main__":
    main()
