"""Historical local references are intentionally absent from release bundles."""
from pathlib import Path
import pytest

DATA = Path(__file__).parent / 'data'


def pytest_collection_modifyitems(items):
    for item in items:
        module = Path(str(item.path)).name
        name = item.originalname or item.name
        required = []
        if module == 'test_abf_mutation.py':
            required = ['null_data_id.xlsx']
        elif module == 'test_roundtrip.py' and name != 'test_convenience_namespace_exports_public_api':
            required = ['null_data_id.xlsx', 'supplier-quality.xlsx']
        elif module == 'test_local_regressions.py':
            if name in {'test_numeric_dictionary_growth_rebuilds_member_order',
                        'test_supplier_resize_retains_numeric_dictionary_ids'}:
                required = ['supplier-quality.xlsx']
            elif name not in {'test_native_empty_column_without_dictionary_is_inspectable',
                              'test_bit_width_rounds_to_supported_codec'}:
                required = ['null_data_id.xlsx']
        elif module == 'test_native_large_storage.py':
            required = ['native-10000.hidx' if 'hash' in name else 'native-10000.idf']
        elif module == 'test_lookup.py' and name == 'test_small_numeric_lookup_matches_native_and_can_grow_again':
            required = ['native-two-values-pos.idf']
        elif module == 'test_schema.py' and name == 'test_import_table_retains_donor_versions_and_rebinds_database':
            required = [item.callspec.params['filename']]
        missing = [name for name in required if not (DATA / name).exists()]
        if missing:
            item.add_marker(pytest.mark.skip(reason='Local reference excluded from source distribution: ' + ', '.join(missing)))
