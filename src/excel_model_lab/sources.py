"""Explicit headless refresh sources. Connections and credentials are not saved."""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Mapping, Protocol

import pandas as pd


class TableSource(Protocol):
    def read(self) -> pd.DataFrame: ...


@dataclass(frozen=True)
class CsvSource:
    path: str | Path
    read_options: Mapping[str, Any] = field(default_factory=dict)

    def read(self) -> pd.DataFrame:
        return pd.read_csv(self.path, **dict(self.read_options))


@dataclass(frozen=True)
class ParquetSource:
    path: str | Path
    read_options: Mapping[str, Any] = field(default_factory=dict)

    def read(self) -> pd.DataFrame:
        return pd.read_parquet(self.path, **dict(self.read_options))


@dataclass(frozen=True)
class SqlSource:
    """Read a query using a caller-owned SQLite or SQLAlchemy connection.

    The caller controls connection lifetime and the query's database permissions.
    Parameter values are passed to the driver without string interpolation.
    """
    connection: Any = field(repr=False)
    query: str
    params: Any = field(default=None, repr=False)
    read_options: Mapping[str, Any] = field(default_factory=dict)

    def read(self) -> pd.DataFrame:
        return pd.read_sql_query(self.query, self.connection, params=self.params,
                                 **dict(self.read_options))
