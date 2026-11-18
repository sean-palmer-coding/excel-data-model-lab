"""Experimental headless tooling for embedded Excel Data Models."""

from .package import inspect_workbook, repack_workbook, semantic_snapshot
from .abf import AbfImage, diff_inventories
from .vertipaq import (
    encode_single_segment_bitpacked_ids,
    encode_single_segment_preserving_layout,
    encode_idf_preserving_layout,
    patch_single_segment_bitpacked_id,
)
from .mutate import set_integer_value, replace_integer_column
from .model import (
    ColumnInfo,
    ExcelModelTemplate,
    ModelValidationError,
    WorkbookModel,
    replace_dataframe,
)
from .spec import build_from_spec, load_spec
from .dictionary import encode_numeric_dictionary, encode_uncompressed_string_dictionary
from .hash_index import HASH_MULTIPLIER, encode_numeric_hash_index
from .measures import Measure, add_measures_to_script, list_script_measures
from .relationships import RelationshipValidationError, Relationship, add_relationships_to_dimension
from .validation import ValidationIssue, ValidationReport
from .reader import ModelReader
from .excel_metadata import ExcelMetadata
from .kpis import KpiValidationError
from .calculated_columns import CalculatedColumnValidationError
from .date_tables import DateTableValidationError, DateTableLookupError
from .perspectives import PerspectiveValidationError, PerspectiveLookupError, PerspectiveTypeError
from .dependencies import DependencyError
from .sources import CsvSource, ParquetSource, SqlSource, TableSource
from .native_sources import LinkedTableSource, OleDbSource, NativeSourceBinding
from .reports import WorkbookReport, PivotTableSpec, PivotChartSpec, SlicerSpec, TimelineSpec
from .report_filters import PivotMemberFilter, PivotSort
from .report_editing import ExistingReport, ExistingControl, inspect_reports, inspect_controls, edit_reports
from .composition import compose_workbook, attach_model
from .cube_formulas import CubeFormulas, CubeFormula, CubeCell
from .dax_evaluation import (DaxEvaluator, DaxEvaluation, DaxBackend, PbixMcpBackend,
    DaxEvaluationError, DaxUnsupportedError, DaxBackendUnavailableError)
from .external_connections import (ExternalModelConnection, ExternalConnectionInfo,
    add_external_connection, inspect_external_connections, AdomdQueryClient,
    ExternalQueryResult, ExternalQueryError)

from .desktop_session import DesktopSession, DesktopSnapshot, DesktopSessionError, DesktopConflictError
from .agent_workflow import (AgentWorkflow, AgentOperation, WorkflowPlan, WorkbookCheckpoint,
    WorkflowResult, WorkflowConflictError, WorkflowExecutionError)

__all__ = [
    "DesktopSession", "DesktopSnapshot", "DesktopSessionError", "DesktopConflictError",
    "AgentWorkflow", "AgentOperation", "WorkflowPlan", "WorkbookCheckpoint",
    "WorkflowResult", "WorkflowConflictError", "WorkflowExecutionError",
    "compose_workbook", "attach_model", "CubeFormulas", "CubeFormula", "CubeCell",
    "DaxEvaluator", "DaxEvaluation", "DaxBackend", "PbixMcpBackend",
    "DaxEvaluationError", "DaxUnsupportedError", "DaxBackendUnavailableError",
    "ExternalModelConnection", "ExternalConnectionInfo", "add_external_connection",
    "inspect_external_connections", "AdomdQueryClient", "ExternalQueryResult", "ExternalQueryError",
    "WorkbookReport", "PivotTableSpec", "PivotChartSpec",
    "SlicerSpec", "TimelineSpec", "PivotMemberFilter", "PivotSort",
    "ExistingReport", "inspect_reports", "edit_reports",
    "ExistingControl", "inspect_controls",
    "ModelReader", "ValidationIssue", "ValidationReport",
    "ExcelMetadata", "KpiValidationError", "DependencyError",
    "CalculatedColumnValidationError", "DateTableValidationError", "DateTableLookupError",
    "PerspectiveValidationError", "PerspectiveLookupError", "PerspectiveTypeError",
    "CsvSource", "ParquetSource", "SqlSource", "TableSource",
    "LinkedTableSource", "OleDbSource", "NativeSourceBinding",
    "AbfImage",
    "diff_inventories",
    "inspect_workbook",
    "repack_workbook",
    "semantic_snapshot",
    "patch_single_segment_bitpacked_id",
    "set_integer_value",
    "encode_single_segment_bitpacked_ids",
    "encode_single_segment_preserving_layout",
    "encode_idf_preserving_layout",
    "replace_integer_column",
    "ColumnInfo",
    "ExcelModelTemplate",
    "WorkbookModel",
    "replace_dataframe",
    "ModelValidationError",
    "encode_uncompressed_string_dictionary",
    "encode_numeric_dictionary",
    "HASH_MULTIPLIER",
    "encode_numeric_hash_index",
    "Measure",
    "add_measures_to_script",
    "list_script_measures",
    "Relationship",
    "RelationshipValidationError",
    "add_relationships_to_dimension",
    "build_from_spec",
    "load_spec",
]
