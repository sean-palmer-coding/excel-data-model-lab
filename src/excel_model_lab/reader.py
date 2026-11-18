"""Read-only model access, independent of the Excel writer's format support.

Metadata availability is determined by PBIXRay and the source format. An empty
metadata collection does not establish that the writer can author that feature.
"""
from pbixray import PBIXRay


class ModelReader:
    """Read XLSX, PBIX or ABF models without constructing an editable model.

    Use as a context manager, particularly with on_disk=True. This reads saved
    source data only; it does not calculate DAX or reflect another editor's edits.
    Currency remains Decimal and nullable native numbers retain exact values;
    native whole-second date encodings are normalized in full and chunked reads.
    """

    def __init__(self, source, *, on_disk=False, temp_dir=None):
        self._reader = PBIXRay(source, on_disk=on_disk, temp_dir=temp_dir)
        from .reader_adapter import normalize_reader_types
        self._closed = False
        try:
            normalize_reader_types(self._reader)
        except BaseException:
            self.close()
            raise

    def _ensure_open(self):
        if self._closed:
            raise RuntimeError('ModelReader is closed')

    def get_table(self, table, columns=None, strings_as_categorical=False):
        self._ensure_open()
        return self._reader.get_table(table, columns=columns,
            strings_as_categorical=strings_as_categorical)

    def iter_table(self, table, columns=None, chunk_size=None,
                   strings_as_categorical=True):
        self._ensure_open()
        chunks = self._reader.iter_table(table, columns=columns,
            chunk_size=chunk_size, strings_as_categorical=strings_as_categorical)
        while True:
            self._ensure_open()
            try:
                chunk = next(chunks)
            except StopIteration:
                return
            yield chunk

    def close(self):
        if not self._closed:
            self._reader.close()
            self._closed = True

    def __enter__(self):
        self._ensure_open()
        return self

    def __exit__(self, *exc):
        self.close()
        return False

    @property
    def excel_metadata(self):
        """Excel ASSL metadata snapshot, without changing PBIXRay properties.

        Includes date-table keys, calculated-column definitions, hierarchies,
        display properties and perspective membership. No DAX is evaluated.
        Unsupported/non-ASSL models return an unavailable, empty snapshot.
        """
        self._ensure_open()
        from .excel_metadata import ExcelMetadata
        from .reader_adapter import excel_metadata_files
        return ExcelMetadata.from_files(excel_metadata_files(self._reader))

    @property
    def tables(self):
        """PBIXRay tables metadata from the saved source."""
        self._ensure_open()
        return self._reader.tables

    @property
    def statistics(self):
        """PBIXRay statistics metadata from the saved source."""
        self._ensure_open()
        return self._reader.statistics

    @property
    def power_query(self):
        """PBIXRay power_query metadata from the saved source."""
        self._ensure_open()
        return self._reader.power_query

    @property
    def m_parameters(self):
        """PBIXRay m_parameters metadata from the saved source."""
        self._ensure_open()
        return self._reader.m_parameters

    @property
    def dax_tables(self):
        """PBIXRay dax_tables metadata from the saved source."""
        self._ensure_open()
        return self._reader.dax_tables

    @property
    def dax_measures(self):
        """PBIXRay dax_measures metadata from the saved source."""
        self._ensure_open()
        return self._reader.dax_measures

    @property
    def dax_columns(self):
        """PBIXRay dax_columns metadata from the saved source."""
        self._ensure_open()
        return self._reader.dax_columns

    @property
    def aggregations(self):
        """PBIXRay aggregations metadata from the saved source."""
        self._ensure_open()
        return self._reader.aggregations

    @property
    def metadata(self):
        """PBIXRay metadata metadata from the saved source."""
        self._ensure_open()
        return self._reader.metadata

    @property
    def size(self):
        """PBIXRay size metadata from the saved source."""
        self._ensure_open()
        return self._reader.size

    @property
    def schema(self):
        """PBIXRay schema metadata from the saved source."""
        self._ensure_open()
        return self._reader.schema

    @property
    def relationships(self):
        """PBIXRay relationships metadata from the saved source."""
        self._ensure_open()
        return self._reader.relationships

    @property
    def connections(self):
        """PBIXRay connections metadata from the saved source."""
        self._ensure_open()
        return self._reader.connections

    @property
    def data_mashup(self):
        """PBIXRay data_mashup metadata from the saved source."""
        self._ensure_open()
        return self._reader.data_mashup

    @property
    def mashup_queries(self):
        """PBIXRay mashup_queries metadata from the saved source."""
        self._ensure_open()
        return self._reader.mashup_queries

    @property
    def rls(self):
        """PBIXRay rls metadata from the saved source."""
        self._ensure_open()
        return self._reader.rls

    @property
    def ols(self):
        """PBIXRay ols metadata from the saved source."""
        self._ensure_open()
        return self._reader.ols

    @property
    def perspectives(self):
        """PBIXRay perspectives metadata from the saved source."""
        self._ensure_open()
        return self._reader.perspectives

    @property
    def tmschema_model(self):
        """PBIXRay tmschema_model metadata from the saved source."""
        self._ensure_open()
        return self._reader.tmschema_model

    @property
    def tmschema_tables(self):
        """PBIXRay tmschema_tables metadata from the saved source."""
        self._ensure_open()
        return self._reader.tmschema_tables

    @property
    def tmschema_columns(self):
        """PBIXRay tmschema_columns metadata from the saved source."""
        self._ensure_open()
        return self._reader.tmschema_columns

    @property
    def tmschema_partitions(self):
        """PBIXRay tmschema_partitions metadata from the saved source."""
        self._ensure_open()
        return self._reader.tmschema_partitions

    @property
    def tmschema_hierarchies(self):
        """PBIXRay tmschema_hierarchies metadata from the saved source."""
        self._ensure_open()
        return self._reader.tmschema_hierarchies

    @property
    def tmschema_levels(self):
        """PBIXRay tmschema_levels metadata from the saved source."""
        self._ensure_open()
        return self._reader.tmschema_levels

    @property
    def tmschema_datasources(self):
        """PBIXRay tmschema_datasources metadata from the saved source."""
        self._ensure_open()
        return self._reader.tmschema_datasources

    @property
    def tmschema_perspectives(self):
        """PBIXRay tmschema_perspectives metadata from the saved source."""
        self._ensure_open()
        return self._reader.tmschema_perspectives

    @property
    def tmschema_perspective_tables(self):
        """PBIXRay tmschema_perspective_tables metadata from the saved source."""
        self._ensure_open()
        return self._reader.tmschema_perspective_tables

    @property
    def tmschema_perspective_columns(self):
        """PBIXRay tmschema_perspective_columns metadata from the saved source."""
        self._ensure_open()
        return self._reader.tmschema_perspective_columns

    @property
    def tmschema_perspective_hierarchies(self):
        """PBIXRay tmschema_perspective_hierarchies metadata from the saved source."""
        self._ensure_open()
        return self._reader.tmschema_perspective_hierarchies

    @property
    def tmschema_perspective_measures(self):
        """PBIXRay tmschema_perspective_measures metadata from the saved source."""
        self._ensure_open()
        return self._reader.tmschema_perspective_measures

    @property
    def tmschema_kpis(self):
        """PBIXRay tmschema_kpis metadata from the saved source."""
        self._ensure_open()
        return self._reader.tmschema_kpis

    @property
    def tmschema_annotations(self):
        """PBIXRay tmschema_annotations metadata from the saved source."""
        self._ensure_open()
        return self._reader.tmschema_annotations

    @property
    def tmschema_extended_properties(self):
        """PBIXRay tmschema_extended_properties metadata from the saved source."""
        self._ensure_open()
        return self._reader.tmschema_extended_properties

    @property
    def tmschema_cultures(self):
        """PBIXRay tmschema_cultures metadata from the saved source."""
        self._ensure_open()
        return self._reader.tmschema_cultures

    @property
    def tmschema_translations(self):
        """PBIXRay tmschema_translations metadata from the saved source."""
        self._ensure_open()
        return self._reader.tmschema_translations

    @property
    def tmschema_linguistic_metadata(self):
        """PBIXRay tmschema_linguistic_metadata metadata from the saved source."""
        self._ensure_open()
        return self._reader.tmschema_linguistic_metadata

    @property
    def tmschema_query_groups(self):
        """PBIXRay tmschema_query_groups metadata from the saved source."""
        self._ensure_open()
        return self._reader.tmschema_query_groups

    @property
    def tmschema_calculation_groups(self):
        """PBIXRay tmschema_calculation_groups metadata from the saved source."""
        self._ensure_open()
        return self._reader.tmschema_calculation_groups

    @property
    def tmschema_calculation_items(self):
        """PBIXRay tmschema_calculation_items metadata from the saved source."""
        self._ensure_open()
        return self._reader.tmschema_calculation_items

    @property
    def tmschema_calculation_expressions(self):
        """PBIXRay tmschema_calculation_expressions metadata from the saved source."""
        self._ensure_open()
        return self._reader.tmschema_calculation_expressions

    @property
    def tmschema_variations(self):
        """PBIXRay tmschema_variations metadata from the saved source."""
        self._ensure_open()
        return self._reader.tmschema_variations

    @property
    def tmschema_attribute_hierarchies(self):
        """PBIXRay tmschema_attribute_hierarchies metadata from the saved source."""
        self._ensure_open()
        return self._reader.tmschema_attribute_hierarchies

    @property
    def tmschema_sets(self):
        """PBIXRay tmschema_sets metadata from the saved source."""
        self._ensure_open()
        return self._reader.tmschema_sets

    @property
    def tmschema_refresh_policies(self):
        """PBIXRay tmschema_refresh_policies metadata from the saved source."""
        self._ensure_open()
        return self._reader.tmschema_refresh_policies

    @property
    def tmschema_detail_rows_definitions(self):
        """PBIXRay tmschema_detail_rows_definitions metadata from the saved source."""
        self._ensure_open()
        return self._reader.tmschema_detail_rows_definitions

    @property
    def tmschema_format_string_definitions(self):
        """PBIXRay tmschema_format_string_definitions metadata from the saved source."""
        self._ensure_open()
        return self._reader.tmschema_format_string_definitions

    @property
    def tmschema_functions(self):
        """PBIXRay tmschema_functions metadata from the saved source."""
        self._ensure_open()
        return self._reader.tmschema_functions

    @property
    def tmschema_calendars(self):
        """PBIXRay tmschema_calendars metadata from the saved source."""
        self._ensure_open()
        return self._reader.tmschema_calendars

    @property
    def tmschema_calendar_column_groups(self):
        """PBIXRay tmschema_calendar_column_groups metadata from the saved source."""
        self._ensure_open()
        return self._reader.tmschema_calendar_column_groups

    @property
    def tmschema_calendar_column_refs(self):
        """PBIXRay tmschema_calendar_column_refs metadata from the saved source."""
        self._ensure_open()
        return self._reader.tmschema_calendar_column_refs

    @property
    def tmschema_alternate_of(self):
        """PBIXRay tmschema_alternate_of metadata from the saved source."""
        self._ensure_open()
        return self._reader.tmschema_alternate_of

    @property
    def tmschema_related_column_details(self):
        """PBIXRay tmschema_related_column_details metadata from the saved source."""
        self._ensure_open()
        return self._reader.tmschema_related_column_details

    @property
    def tmschema_group_by_columns(self):
        """PBIXRay tmschema_group_by_columns metadata from the saved source."""
        self._ensure_open()
        return self._reader.tmschema_group_by_columns

    @property
    def tmschema_binding_info(self):
        """PBIXRay tmschema_binding_info metadata from the saved source."""
        self._ensure_open()
        return self._reader.tmschema_binding_info

    @property
    def tmschema_analytics_ai_metadata(self):
        """PBIXRay tmschema_analytics_ai_metadata metadata from the saved source."""
        self._ensure_open()
        return self._reader.tmschema_analytics_ai_metadata

    @property
    def tmschema_data_coverage_definitions(self):
        """PBIXRay tmschema_data_coverage_definitions metadata from the saved source."""
        self._ensure_open()
        return self._reader.tmschema_data_coverage_definitions

    @property
    def tmschema_role_memberships(self):
        """PBIXRay tmschema_role_memberships metadata from the saved source."""
        self._ensure_open()
        return self._reader.tmschema_role_memberships

    @property
    def tmschema_column_permissions(self):
        """PBIXRay tmschema_column_permissions metadata from the saved source."""
        self._ensure_open()
        return self._reader.tmschema_column_permissions
