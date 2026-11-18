"""Version-specific PBIXRay internals and lossless public read normalization.

Keep private decoder access here. Currency, nullable integers/booleans and native
whole-second dates are normalized consistently for whole and streaming reads.
"""
import pandas as pd


def storage_schema(reader):
    return reader._vertipaq_decoder._meta.schema_df


def segment_metadata(reader, metadata):
    return reader._vertipaq_decoder._meta.get_segment_meta(metadata)


def decode_dictionary(reader, payload, data_type):
    return reader._vertipaq_decoder._read_dictionary(payload, data_type)


def normalize_reader_types(reader):
    # numpy bool turns a decoded NaN into True. Use pandas' nullable bool
    # when asking the independent decoder to materialize boolean columns.
    schema = storage_schema(reader)
    # Native calculated keys retain DataType=Empty and report their effective
    # type in InferredDatatype, which PBIXRay's ASSL schema does not yet apply.
    from .excel_metadata import ExcelMetadata
    metadata = ExcelMetadata.from_files(excel_metadata_files(reader))
    dtypes = {'BigInt': 'Int64', 'Integer': 'Int64', 'Double': 'float64',
              'Currency': 'object', 'Date': 'datetime64[ns]', 'Boolean': 'boolean',
              'WChar': 'string'}
    for column in metadata.calculated_columns.itertuples(index=False):
        if column.processing_state != 'Processed' or column.data_type not in dtypes:
            continue
        mask = (schema['TableName'] == column.table) & (schema['ColumnName'] == column.name)
        schema.loc[mask, 'SemanticType'] = column.data_type
        schema.loc[mask, 'PandasDataType'] = dtypes[column.data_type]
    schema.loc[schema['PandasDataType'] == 'bool', 'PandasDataType'] = 'boolean'
    # Currency is fixed-point int64. Keep Decimal output instead of converting
    # it to binary64, which loses cents for large but valid native values.
    schema.loc[schema['SemanticType'] == 'Currency', 'PandasDataType'] = 'object'
    decoder = reader._vertipaq_decoder
    original_codes_to_series = decoder._codes_to_series
    def exact_integer_dictionary(codes, lookup, as_categorical):
        import numbers
        import numpy as np
        if not as_categorical and len(lookup.categories) and all(
                isinstance(value, numbers.Integral) for value in lookup.categories):
            categories = lookup.categories
            # Explicit object dtype prevents pandas inferring binary64 when a
            # BLANK shares a vector with large signed int64 dictionary members.
            lookup_values = np.append(categories, None)
            return pd.Series(lookup_values[np.where(codes >= 0, codes, len(categories))],
                             dtype=object)
        return original_codes_to_series(codes, lookup, as_categorical)
    decoder._codes_to_series = exact_integer_dictionary
    # Date serials pass through float day counts in PBIXRay. Restore the
    # exact whole-second grid declared by value-encoded storage metadata.
    # Dictionary dates and subsecond encodings must remain untouched.
    date_grids = {}
    for _, row in schema.iterrows():
        if row.get('SemanticType') != 'Date' or pd.notna(row.get('Dictionary')):
            continue
        magnitude = row.get('Magnitude')
        if pd.notna(magnitude) and magnitude > 0 and 86400 % magnitude == 0:
            date_grids[(row['TableName'], row['ColumnName'])] = int(86400 / magnitude)
    if date_grids:
        get_table, iter_table = reader.get_table, reader.iter_table
        def normalize_dates(result, table):
            for (table_name, column), seconds in date_grids.items():
                if table == table_name and column in result:
                    result[column] = result[column].dt.round(f'{seconds}s')
            return result
        def normalized_table(table, *args, **kwargs):
            return normalize_dates(get_table(table, *args, **kwargs), table)
        def normalized_chunks(table, *args, **kwargs):
            for result in iter_table(table, *args, **kwargs):
                yield normalize_dates(result, table)
        reader.get_table = normalized_table
        reader.iter_table = normalized_chunks



def excel_metadata_files(reader):
    """Yield ASSL metadata using the already loaded (possibly mmap) model.

    Only individual metadata streams are materialized; model data is never
    reopened or copied. Private PBIXRay access remains confined to this adapter.
    """
    from pbixray.utils import get_data_slice
    model = reader._data_model
    suffixes = ('.dim.xml', '.cub.xml', '.persp.xml')
    for entry in model.file_log:
        name = entry['FileName']
        if name.endswith(suffixes):
            yield name, get_data_slice(model, name)
