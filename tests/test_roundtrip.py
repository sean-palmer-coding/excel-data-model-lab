from pathlib import Path
from zipfile import ZipFile

from excel_model_lab.package import MODEL_PART, inspect_workbook, repack_workbook, semantic_snapshot
from excel_model_lab.abf import AbfImage, diff_inventories, replace_model_part
from excel_model_lab.vertipaq import patch_single_segment_bitpacked_id
from excel_model_lab.vertipaq import (
    encode_idf_preserving_layout,
    encode_single_segment_bitpacked_ids,
)
from excel_model_lab.mutate import set_integer_value, replace_integer_column
from excel_model_lab.model import ExcelModelTemplate, WorkbookModel, replace_dataframe
from excel_model_lab.spec import build_from_spec, load_spec
from excel_model_lab.abf import UnsupportedModelError
from excel_model_lab.hash_index import encode_numeric_hash_index

FIXTURE = Path(__file__).parent / "data" / "null_data_id.xlsx"
SUPPLIER_FIXTURE = Path(__file__).parent / "data" / "supplier-quality.xlsx"


def test_fixture_contains_model():
    report = inspect_workbook(FIXTURE)
    assert report["has_data_model"]
    assert report["model"]["size"] > 0


def test_convenience_namespace_exports_public_api():
    from excelmodel import WorkbookModel as PublicWorkbookModel

    assert PublicWorkbookModel is WorkbookModel


def test_byte_preserving_repack_is_semantically_equivalent(tmp_path):
    output = repack_workbook(FIXTURE, tmp_path / "roundtrip.xlsx")
    with ZipFile(FIXTURE) as before, ZipFile(output) as after:
        assert before.namelist() == after.namelist()
        assert before.read(MODEL_PART) == after.read(MODEL_PART)
    assert semantic_snapshot(FIXTURE) == semantic_snapshot(output)


def test_model_member_can_be_replaced_without_touching_other_parts(tmp_path):
    with ZipFile(FIXTURE) as package:
        original_model = package.read(MODEL_PART)
        original_parts = {
            item.filename: package.read(item.filename)
            for item in package.infolist()
            if item.filename != MODEL_PART
        }
    output = repack_workbook(FIXTURE, tmp_path / "replaced.xlsx", replacement_model=original_model)
    with ZipFile(output) as package:
        assert all(package.read(name) == value for name, value in original_parts.items())


def test_abf_inventory_is_complete():
    image = AbfImage.from_workbook(FIXTURE)
    report = image.inventory()
    assert report["file_count"] == len(report["files"]) == 36
    assert report["data_offset"] == 4096
    assert report["virtual_directory_offset"] == 102400
    assert {row["kind"] for row in report["files"]} >= {
        "backup_log", "column_segment", "dictionary", "table_metadata"
    }


def test_abf_clone_is_reassembled_and_byte_identical(tmp_path):
    image = AbfImage.from_workbook(FIXTURE)
    cloned_model = image.clone()
    assert cloned_model == image.data
    output = replace_model_part(FIXTURE, tmp_path / "abf-clone.xlsx", cloned_model)
    assert semantic_snapshot(output) == semantic_snapshot(FIXTURE)
    assert diff_inventories(image, AbfImage.from_workbook(output))["summary"] == {
        "unchanged": 36, "changed": 0, "added": 0, "removed": 0
    }


def test_abf_extract_writes_every_virtual_file(tmp_path):
    image = AbfImage.from_workbook(FIXTURE)
    output = image.extract(tmp_path / "virtual-files")
    assert (output / "inventory.json").is_file()
    assert len([p for p in output.iterdir() if p.name != "inventory.json"]) == 36


def test_abf_assemble_regenerates_metadata_and_decodes(tmp_path):
    image = AbfImage.from_workbook(FIXTURE)
    target = next(f for f in image.files if f.logical_path and f.logical_path.endswith(".A.0.idf"))
    logical = image.read_logical_file(target.storage_path)
    rebuilt = image.assemble({target.storage_path: logical})
    assert rebuilt != image.data
    output = replace_model_part(FIXTURE, tmp_path / "assembled.xlsx", rebuilt)
    assert semantic_snapshot(output) == semantic_snapshot(FIXTURE)
    new_image = AbfImage.from_workbook(output)
    assert new_image.read_logical_file(target.storage_path) == logical


def test_first_integer_value_can_be_mutated_headlessly(tmp_path):
    image = AbfImage.from_workbook(FIXTURE)
    target = next(f for f in image.files if f.logical_path and f.logical_path.endswith(".K.0.idf"))
    patched = patch_single_segment_bitpacked_id(
        image.read_logical_file(target.storage_path),
        position=0,
        bit_width=10,
        min_data_id=3,
        new_data_id=12,
    )
    output = replace_model_part(
        FIXTURE,
        tmp_path / "mutated-k.xlsx",
        image.assemble({target.storage_path: patched}),
    )
    from pbixray import PBIXRay

    with PBIXRay(FIXTURE) as before, PBIXRay(output) as after:
        left = before.get_table("TheTable")
        right = after.get_table("TheTable")
        differences = left.compare(right)
        assert len(differences) == 1
        assert differences.loc[0, ("K", "self")] == 10
        assert differences.loc[0, ("K", "other")] == 11
        assert before.dax_measures.equals(after.dax_measures)


def test_high_level_integer_mutation_derives_storage_metadata(tmp_path):
    output = set_integer_value(
        FIXTURE,
        tmp_path / "high-level-mutation.xlsx",
        table="TheTable",
        column="K",
        position=0,
        value=11,
    )
    from pbixray import PBIXRay

    with PBIXRay(output) as model:
        table = model.get_table("TheTable")
        assert table.loc[0, "K"] == 11
        assert len(table) == 500


def test_complete_integer_idf_can_be_regenerated(tmp_path):
    from pbixray import PBIXRay

    with PBIXRay(FIXTURE) as model:
        original = model.get_table("TheTable")
    values = original["K"].astype(int).tolist()
    values[0] = 11
    output = replace_integer_column(
        FIXTURE,
        tmp_path / "regenerated-column.xlsx",
        table="TheTable",
        column="K",
        values=values,
    )
    with PBIXRay(output) as model:
        result = model.get_table("TheTable")
        assert result["K"].astype(int).tolist() == values
        assert result.drop(columns=["K"]).equals(original.drop(columns=["K"]))


def test_template_reports_supported_and_unsupported_columns():
    template = ExcelModelTemplate(FIXTURE)
    columns = {(item.table, item.name): item for item in template.columns}
    assert columns[("TheTable", "A")].supported
    assert columns[("TheTable", "K")].supported
    assert columns[("TheTable", "N")].supported
    assert columns[("TheTable", "N")].nullable
    assert columns[("TheTable", "S")].supported


def test_multiple_integer_columns_are_replaced_atomically(tmp_path):
    from pbixray import PBIXRay

    with PBIXRay(FIXTURE) as model:
        original = model.get_table("TheTable")
    expected = original.copy()
    expected.loc[0, "A"] = 6
    expected.loc[1, "K"] = 21

    output = (
        ExcelModelTemplate(FIXTURE)
        .replace_columns(
            "TheTable",
            {
                "A": expected["A"].tolist(),
                "K": expected["K"].tolist(),
            },
        )
        .save(tmp_path / "two-columns.xlsx")
    )
    with PBIXRay(output) as model:
        result = model.get_table("TheTable")
        assert result.equals(expected)
        assert len(model.dax_measures) == 1


def test_dataframe_wrapper_replaces_selected_columns(tmp_path):
    from pbixray import PBIXRay

    with PBIXRay(FIXTURE) as model:
        frame = model.get_table("TheTable")
    frame.loc[0, "A"] = 6
    frame.loc[0, "K"] = 11
    output = replace_dataframe(
        FIXTURE,
        tmp_path / "dataframe.xlsx",
        table="TheTable",
        dataframe=frame,
        columns=["A", "K"],
    )
    with PBIXRay(output) as model:
        result = model.get_table("TheTable")
        assert result.equals(frame)


def test_failed_batch_does_not_leave_partial_pending_mutations(tmp_path):
    from pbixray import PBIXRay

    with PBIXRay(FIXTURE) as model:
        original = model.get_table("TheTable")
    template = ExcelModelTemplate(FIXTURE)
    try:
            template.replace_columns(
                "TheTable",
                {
                    "A": original["A"].tolist(),
                    "__XL_RowNumber": original["__XL_RowNumber"].tolist(),
                },
            )
    except UnsupportedModelError:
        pass
    else:
        raise AssertionError("Constant column replacement should have been rejected")
    output = template.save(tmp_path / "unchanged-after-failure.xlsx")
    with PBIXRay(output) as model:
        assert model.get_table("TheTable").equals(original)


def test_nullable_integer_column_preserves_and_moves_nulls(tmp_path):
    import pandas as pd
    from pbixray import PBIXRay

    with PBIXRay(FIXTURE) as model:
        original = model.get_table("TheTable")
    values = original["N"].tolist()
    values[0] = None
    values[6] = 105
    output = (
        ExcelModelTemplate(FIXTURE)
        .replace_column("TheTable", "N", values)
        .save(tmp_path / "nullable-integer.xlsx")
    )
    with PBIXRay(output) as model:
        result = model.get_table("TheTable")
        assert pd.isna(result.loc[0, "N"])
        assert result.loc[6, "N"] == 105
        assert result.drop(columns=["N"]).equals(original.drop(columns=["N"]))


def test_save_validation_is_enabled_by_default(tmp_path, monkeypatch):
    from pbixray import PBIXRay

    with PBIXRay(FIXTURE) as model:
        values = model.get_table("TheTable")["K"].tolist()
    template = ExcelModelTemplate(FIXTURE).replace_column("TheTable", "K", values)
    calls = []
    original = template._validate_output

    def observed(path):
        assert path.exists()
        assert not (tmp_path / "validated.xlsx").exists()
        calls.append(path)
        return original(path)

    monkeypatch.setattr(template, "_validate_output", observed)
    output = template.save(tmp_path / "validated.xlsx")
    assert len(calls) == 1
    assert output.exists()


def test_hybrid_rle_currency_column_can_be_replaced(tmp_path):
    from pbixray import PBIXRay

    with PBIXRay(FIXTURE) as model:
        original = model.get_table("TheTable")
    values = original["C"].tolist()
    values[100] = 0.02
    values[101] = 0.01
    output = (
        ExcelModelTemplate(FIXTURE)
        .replace_column("TheTable", "C", values)
        .save(tmp_path / "currency-hybrid.xlsx")
    )
    with PBIXRay(output) as model:
        result = model.get_table("TheTable")
        assert result.loc[100, "C"] == 0.02
        assert result.loc[101, "C"] == 0.01
        assert result.drop(columns=["C"]).equals(original.drop(columns=["C"]))


def test_hybrid_rle_rejects_broken_run_without_writing(tmp_path):
    from pbixray import PBIXRay

    with PBIXRay(FIXTURE) as model:
        values = model.get_table("TheTable")["C"].tolist()
    values[0] = 0.01
    destination = tmp_path / "must-not-exist.xlsx"
    template = ExcelModelTemplate(FIXTURE).replace_column("TheTable", "C", values)
    try:
        template.save(destination)
    except UnsupportedModelError as exc:
        assert "RLE run" in str(exc)
    else:
        raise AssertionError("Breaking an RLE run should be rejected")
    assert not destination.exists()


def test_multiple_idf_segments_are_encoded_in_order():
    first = encode_single_segment_bitpacked_ids(
        b"\x01\x00\x00\x00\x00\x00\x00\x00"
        b"\xff\xff\xff\xff\x03\x00\x00\x00"
        b"\x01\x00\x00\x00\x00\x00\x00\x00"
        b"\x00\x00\x00\x00\x00\x00\x00\x00",
        data_ids=[3, 4, 5], bit_width=3, min_data_id=3,
    )
    second = encode_single_segment_bitpacked_ids(
        b"\x01\x00\x00\x00\x00\x00\x00\x00"
        b"\xff\xff\xff\xff\x02\x00\x00\x00"
        b"\x01\x00\x00\x00\x00\x00\x00\x00"
        b"\x00\x00\x00\x00\x00\x00\x00\x00",
        data_ids=[10, 11], bit_width=4, min_data_id=10,
    )
    rebuilt = encode_idf_preserving_layout(
        first + second,
        data_ids=[5, 4, 3, 11, 10],
        segments=[
            {"records": 3, "bit_width": 3, "min_data_id": 3, "has_nulls": False},
            {"records": 2, "bit_width": 4, "min_data_id": 10, "has_nulls": False},
        ],
    )
    assert len(rebuilt) == len(first + second)
    assert rebuilt != first + second


def test_string_dictionary_and_idf_are_replaced_together(tmp_path):
    import pandas as pd
    from pbixray import PBIXRay

    with PBIXRay(FIXTURE) as model:
        original = model.get_table("TheTable")
    values = original["S"].map(lambda value: None if value is None else "new-" + value).tolist()
    output = (
        ExcelModelTemplate(FIXTURE)
        .replace_column("TheTable", "S", values)
        .save(tmp_path / "strings.xlsx")
    )
    with PBIXRay(output) as model:
        result = model.get_table("TheTable")
        expected = pd.Series(values, dtype="string")
        assert result["S"].reset_index(drop=True).equals(expected)
        assert result.drop(columns=["S"]).equals(original.drop(columns=["S"]))


def test_string_values_shrink_dictionary_to_used_members(tmp_path):
    import pandas as pd
    from pbixray import PBIXRay

    with PBIXRay(FIXTURE) as model:
        original = model.get_table("TheTable")
    values = original["S"].map(
        lambda value: None if pd.isna(value) else ("even" if int(value[1:]) % 2 == 0 else "odd")
    ).tolist()
    output = (
        ExcelModelTemplate(FIXTURE)
        .replace_column("TheTable", "S", values)
        .save(tmp_path / "two-string-values.xlsx")
    )
    with PBIXRay(output) as model:
        result = model.get_table("TheTable")
        expected = pd.Series(values, dtype="string")
        assert result["S"].reset_index(drop=True).equals(expected)
        schema = model._vertipaq_decoder._meta.schema_df
        metadata = schema[(schema["TableName"] == "TheTable") & (schema["ColumnName"] == "S")].iloc[0]
        assert int(metadata["Cardinality"]) == 3  # two used strings plus the blank member


def test_numeric_hash_index_can_be_regenerated_from_ids():
    from io import BytesIO
    from pbixray.column_data.hidx import ColumnDataHidx

    image = AbfImage.from_workbook(FIXTURE)
    item = next(f for f in image.files if f.logical_path and f.logical_path.endswith("$K.hidx"))
    original = image.read_logical_file(item)
    rebuilt = encode_numeric_hash_index(original, list(range(3, 503, 2)))
    parsed = ColumnDataHidx.from_io(BytesIO(rebuilt))
    assert parsed.number_of_records == 250
    entries = [
        entry
        for bucket in parsed.hash_bin_entries
        for entry in bucket.m_rg_local_entries
        if entry.m_hash
    ] + [entry for entry in parsed.overflow_hash_entries if entry.m_hash]
    assert sorted(entry.m_hash for entry in entries) == list(range(3, 503, 2))
    assert sorted(entry.m_key for entry in entries) == list(range(250))


def test_table_row_count_can_shrink(tmp_path):
    from pbixray import PBIXRay

    with PBIXRay(FIXTURE) as model:
        original = model.get_table("TheTable")
    frame = original[["A", "N", "C", "S", "K"]].iloc[100:125].reset_index(drop=True)
    output = (
        ExcelModelTemplate(FIXTURE)
        .replace_table("TheTable", frame)
        .save(tmp_path / "resized-table.xlsx")
    )
    with PBIXRay(output) as model:
        result = model.get_table("TheTable")
        assert len(result) == 25
        assert result[["A", "N", "C", "S", "K"]].reset_index(drop=True).equals(frame)


def test_table_row_count_can_grow(tmp_path):
    import pandas as pd
    from pbixray import PBIXRay

    with PBIXRay(FIXTURE) as model:
        original = model.get_table("TheTable")[["A", "N", "C", "S", "K"]]
    frame = pd.concat([original, original.iloc[:250]], ignore_index=True)
    output = (
        ExcelModelTemplate(FIXTURE)
        .replace_table("TheTable", frame)
        .save(tmp_path / "grown-table.xlsx")
    )
    with PBIXRay(output) as model:
        result = model.get_table("TheTable")[["A", "N", "C", "S", "K"]]
        assert len(result) == 750
        assert result.reset_index(drop=True).equals(frame)


def test_excel_date_column_can_be_replaced(tmp_path):
    import pandas as pd
    from pbixray import PBIXRay

    with PBIXRay(SUPPLIER_FIXTURE) as model:
        original = model.get_table("Metrics")
    values = original["Date"].tolist()
    values[0] = pd.Timestamp("2014-11-30")
    output = (
        ExcelModelTemplate(SUPPLIER_FIXTURE)
        .replace_column("Metrics", "Date", values)
        .save(tmp_path / "dates.xlsx")
    )
    with PBIXRay(output) as model:
        result = model.get_table("Metrics")
        assert result.loc[0, "Date"] == pd.Timestamp("2014-11-30")
        assert result.drop(columns=["Date"]).equals(original.drop(columns=["Date"]))


def test_dax_measure_is_added_to_model_script(tmp_path):
    from excel_model_lab.measures import list_script_measures

    output = (
        ExcelModelTemplate(FIXTURE)
        .add_measure(
            "TheTable",
            "Total K",
            "SUM('TheTable'[K])",
            format_string="#,##0",
            display_folder="Generated",
            description="Generated test measure",
        )
        .save(tmp_path / "measure.xlsx")
    )
    image = AbfImage.from_workbook(output)
    script = next(f for f in image.files if f.logical_path and f.logical_path.endswith(".scr.xml"))
    measures = list_script_measures(image.read_logical_file(script))
    assert [(m.table, m.name, m.expression) for m in measures] == [
        ("TheTable", "Total K", "SUM('TheTable'[K])")
    ]


def test_relationship_is_added_and_decoded(tmp_path):
    from pbixray import PBIXRay

    output = (
        ExcelModelTemplate(SUPPLIER_FIXTURE)
        .add_relationship(
            "Metrics", "Defect Qty", "Defect Type", "Sort", active=False
        )
        .save(tmp_path / "relationship.xlsx")
    )
    with PBIXRay(output) as model:
        relationships = model.relationships
        match = relationships[
            (relationships["FromTableName"] == "Metrics")
            & (relationships["FromColumnName"] == "Defect Qty")
            & (relationships["ToTableName"] == "Defect Type")
            & (relationships["ToColumnName"] == "Sort")
        ]
        assert len(match) == 1
        assert not bool(match.iloc[0]["IsActive"])


def test_numeric_dictionary_column_can_be_replaced(tmp_path):
    from pbixray import PBIXRay

    with PBIXRay(SUPPLIER_FIXTURE) as model:
        original = model.get_table("Plant")
    values = original["Plant ID"].tolist()
    values[0], values[1] = values[1], values[0]
    output = (
        ExcelModelTemplate(SUPPLIER_FIXTURE)
        .replace_column("Plant", "Plant ID", values)
        .save(tmp_path / "numeric-dictionary.xlsx")
    )
    with PBIXRay(output) as model:
        result = model.get_table("Plant")
        assert result["Plant ID"].tolist() == values
        assert result.drop(columns=["Plant ID"]).equals(original.drop(columns=["Plant ID"]))


def test_realistic_dictionary_table_can_be_resized(tmp_path):
    from pbixray import PBIXRay

    with PBIXRay(SUPPLIER_FIXTURE) as model:
        frame = model.get_table("Plant").iloc[:12].reset_index(drop=True)
    output = (
        ExcelModelTemplate(SUPPLIER_FIXTURE)
        .replace_table("Plant", frame)
        .save(tmp_path / "resized-plant.xlsx")
    )
    with PBIXRay(output) as model:
        result = model.get_table("Plant").reset_index(drop=True)
        assert result.equals(frame)


def test_workbook_model_can_serialize_to_bytes():
    from io import BytesIO
    from pbixray import PBIXRay

    with PBIXRay(FIXTURE) as model:
        values = model.get_table("TheTable")["K"].tolist()
    values[0] = 11
    payload = (
        WorkbookModel.from_template(FIXTURE)
        .replace_column("TheTable", "K", values)
        .to_bytes()
    )
    with ZipFile(BytesIO(payload)) as package:
        assert package.read(MODEL_PART)


def test_multiple_complete_tables_are_replaced_atomically(tmp_path):
    from pbixray import PBIXRay

    with PBIXRay(SUPPLIER_FIXTURE) as model:
        plant = model.get_table("Plant").iloc[:10].reset_index(drop=True)
        defect_type = model.get_table("Defect Type").iloc[:2].reset_index(drop=True)
    output = (
        WorkbookModel.from_template(SUPPLIER_FIXTURE)
        .replace_tables({"Plant": plant, "Defect Type": defect_type})
        .save(tmp_path / "two-tables.xlsx")
    )
    with PBIXRay(output) as model:
        assert model.get_table("Plant").reset_index(drop=True).equals(plant)
        assert model.get_table("Defect Type").reset_index(drop=True).equals(defect_type)


def test_declarative_spec_builds_a_valid_workbook(tmp_path):
    from pbixray import PBIXRay

    with PBIXRay(FIXTURE) as model:
        frame = model.get_table("TheTable")[["A", "N", "C", "S", "K"]].iloc[:20]
    frame.to_csv(tmp_path / "table.csv", index=False)
    spec_path = tmp_path / "model.yaml"
    spec_path.write_text(
        """\
tables:
  TheTable:
    source: table.csv
measures:
  - table: TheTable
    name: Generated Total K
    expression: SUM('TheTable'[K])
    format: '#,##0'
""",
        encoding="utf-8",
    )
    output = build_from_spec(
        FIXTURE,
        tmp_path / "from-spec.xlsx",
        load_spec(spec_path),
        base_directory=tmp_path,
    )
    with PBIXRay(output) as model:
        assert model.get_table("TheTable")[["A", "N", "C", "S", "K"]].reset_index(drop=True).equals(
            frame.reset_index(drop=True)
        )
