from __future__ import annotations

from .reader_adapter import storage_schema, segment_metadata, decode_dictionary


from fractions import Fraction
from pathlib import Path
import pandas as pd

from .abf import AbfImage, UnsupportedModelError, replace_model_part
from .model import ExcelModelTemplate
from .vertipaq import patch_single_segment_bitpacked_id, encode_single_segment_bitpacked_ids


def set_integer_value(
    source: str | Path,
    destination: str | Path,
    *,
    table: str,
    column: str,
    position: int,
    value: int,
) -> Path:
    """Set one value in a supported value-encoded integer model column.

    Position is VertiPaq storage order, which may differ from worksheet source
    order. This prototype API derives all physical encoding parameters and is
    intended to disappear behind whole-table replacement.
    """
    from pbixray import PBIXRay

    source = Path(source)
    with PBIXRay(source) as model:
        rows = storage_schema(model)
        match = rows[(rows["TableName"] == table) & (rows["ColumnName"] == column)]
        if len(match) != 1:
            raise KeyError(f"Expected one model column {table}.{column}, found {len(match)}")
        metadata = match.iloc[0]
        if metadata["DataType"] != "int64" or pd.notna(metadata["Dictionary"]):
            raise UnsupportedModelError("Only value-encoded int64 columns are supported")
        segments = segment_metadata(model, metadata)
        if len(segments) != 1:
            raise UnsupportedModelError("Only one-segment columns are supported")
        segment = segments[0]
        if segment.get("has_nulls"):
            raise UnsupportedModelError("Nullable segments are not supported by this mutation API")
        idf_name = metadata["IDF"]
        if isinstance(value, bool) or not isinstance(value, int):
            raise TypeError("value must be an integer")
        base_id = int(metadata["BaseId"])
        # Exact rational arithmetic: float math silently rounds integers above
        # 2**53, and Decimal would depend on the caller's context precision.
        magnitude = Fraction(repr(float(metadata["Magnitude"])))
        exact = value * magnitude - base_id
        if exact.denominator != 1:
            raise ValueError(f"Value {value} is not exactly representable by this column")
        encoded = int(exact)
        # This patch rewrites one data ID only; column statistics, hash
        # indexes and dependent relationship indexes are not rebuilt. Keep the
        # value inside the existing range so those structures stay valid.
        current = model.get_table(table)[column]
        if current.empty or not int(current.min()) <= value <= int(current.max()):
            raise ValueError("set_integer_value only supports values inside the column's existing "
                             "range; use replace_integer_column or ExcelModelTemplate.replace_dataframe")

    image = AbfImage.from_workbook(source)
    matches = [
        item for item in image.files
        if item.logical_path and item.logical_path.replace("/", "\\").endswith("\\" + idf_name)
    ]
    if len(matches) != 1:
        raise KeyError(f"Could not uniquely resolve physical storage for {idf_name}")
    target = matches[0]
    patched = patch_single_segment_bitpacked_id(
        image.read_logical_file(target.storage_path),
        position=position,
        bit_width=int(segment["bit_width"]),
        min_data_id=int(segment["min_data_id"]),
        new_data_id=encoded,
    )
    return replace_model_part(
        source,
        destination,
        image.assemble({target.storage_path: patched}),
    )


def replace_integer_column(
    source: str | Path,
    destination: str | Path,
    *,
    table: str,
    column: str,
    values: list[int],
) -> Path:
    """Replace every value in a supported integer column in storage order."""
    return (
        ExcelModelTemplate(source)
        .replace_column(table, column, values)
        .save(destination)
    )
