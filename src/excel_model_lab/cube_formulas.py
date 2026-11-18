"""Typed Excel CUBE formula construction; Excel evaluates every result."""
from dataclasses import dataclass
import re

from .errors import ModelValidationError
from .report_filters import member_unique_name, source_member_values, _scalar

MODEL_CONNECTION = 'ThisWorkbookDataModel'


def _text(value, label, *, empty=False):
    if not isinstance(value, str) or (not empty and not value):
        raise ModelValidationError(f'{label} must be text')
    if any(not (ord(c) in (9, 10, 13) or 32 <= ord(c) <= 0xD7FF or
                0xE000 <= ord(c) <= 0xFFFD or 0x10000 <= ord(c) <= 0x10FFFF) for c in value):
        raise ModelValidationError(f'{label} contains invalid XML text')
    return value


def _literal(value):
    _text(value, 'Formula argument', empty=True)
    if len(value.encode('utf-16-le')) // 2 > 255:
        raise ModelValidationError('Inline CUBE text arguments are limited to 255 characters')
    return '"' + value.replace('"', '""') + '"'


def _identifier(value):
    return '[' + value.replace(']', ']]') + ']'


@dataclass(frozen=True)
class CubeCell:
    """Reference to a cell containing a CUBE member or set, not its caption."""
    address: str
    kind: str = 'member'
    sheet: str | None = None

    def __post_init__(self):
        match = re.fullmatch(r'\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})', self.address) if isinstance(self.address, str) else None
        if not match:
            raise ModelValidationError('CubeCell requires a single A1 cell address')
        col = 0
        for c in match[1].upper():
            col = col * 26 + ord(c) - 64
        if col > 16384 or int(match[2]) > 1048576:
            raise ModelValidationError('CubeCell is outside Excel worksheet bounds')
        if self.kind not in ('member', 'set'):
            raise ModelValidationError('CubeCell kind must be member or set')
        if self.sheet is not None:
            from .reports import _sheet
            _sheet(self.sheet)

    def __str__(self):
        return (("'" + self.sheet.replace("'", "''") + "'!") if self.sheet is not None else '') + self.address.upper()


@dataclass(frozen=True)
class CubeFormula:
    """A formula plus its semantic result kind; no calculated value is cached."""
    formula: str
    kind: str
    connection: str
    member_key: str | None = None
    hierarchy: tuple | None = None

    def __str__(self):
        return self.formula


class CubeFormulas:
    """Build bounded, model-validated formulas for the embedded Data Model.

    Model member validation uses current stored values, not DAX evaluation.
    Cell references are syntactically checked; Excel verifies their contents.
    """
    def __init__(self, model):
        self.model = model
        self.connection = MODEL_CONNECTION

    def _formula(self, name, arguments, kind, **metadata):
        expression = '=' + name + '(' + ','.join(arguments) + ')'
        if len(expression.encode('utf-16-le')) // 2 > 8192:
            raise ModelValidationError('Formula exceeds Excel\'s 8192-character limit')
        return CubeFormula(expression, kind, self.connection, **metadata)

    def _argument(self, value, kinds):
        if isinstance(value, CubeCell) and value.kind in kinds:
            return str(value)
        if isinstance(value, CubeFormula) and value.kind in kinds and value.connection == self.connection:
            return value.formula.removeprefix('=')
        raise ModelValidationError('Expected a compatible CUBE ' + '/'.join(kinds) + ' formula or cell reference')

    def member(self, table, column, value, *, caption=None):
        fields = {(c.table, c.name) for c in self.model.columns}
        if (table, column) not in fields:
            raise ModelValidationError(f'Unknown model field: {table}.{column}')
        key = member_unique_name(table, column, value)
        wanted = _scalar(value)
        if not any(wanted == _scalar(v) for v in source_member_values(self.model, table, column)):
            raise ModelValidationError(f'Selected member is absent from {table}.{column}')
        args = [_literal(self.connection), _literal(key)]
        if caption is not None:
            args.append(_literal(caption))
        return self._formula('CUBEMEMBER', args, 'member', member_key=key, hierarchy=(table, column))

    def measure(self, name, *, caption=None):
        from .measures import list_script_measures
        measures = list(self.model._pending_measures)
        for item in self.model._image.files:
            if item.logical_path and item.logical_path.endswith('.scr.xml'):
                measures.extend(list_script_measures(self.model._image.read_logical_file(item)))
        if name not in {m.name for m in measures}:
            raise ModelValidationError(f'Unknown model measure: {name}')
        key = '[Measures].' + _identifier(name)
        args = [_literal(self.connection), _literal(key)]
        if caption is not None:
            args.append(_literal(caption))
        return self._formula('CUBEMEMBER', args, 'member', member_key=key, hierarchy=('Measures',))

    def value(self, *members):
        if not members:
            raise ModelValidationError('CUBEVALUE requires at least one explicit member or set')
        known = [m.hierarchy for m in members if isinstance(m, CubeFormula) and m.hierarchy]
        if len(set(known)) != len(known):
            raise ModelValidationError('CUBEVALUE cannot select the same hierarchy twice')
        return self._formula('CUBEVALUE', [_literal(self.connection)] +
                             [self._argument(v, ('member', 'set')) for v in members], 'value')

    def set(self, members, *, caption=''):
        if isinstance(members, (str, bytes)):
            raise ModelValidationError('A CUBE set requires member formulas, not MDX text')
        try:
            members = tuple(members)
        except TypeError as exc:
            raise ModelValidationError('A CUBE set requires a sequence of member formulas') from exc
        if not members or any(not isinstance(m, CubeFormula) or m.kind != 'member' or
                              not m.member_key or m.connection != self.connection for m in members):
            raise ModelValidationError('A CUBE set requires explicit model member formulas')
        if len({m.hierarchy for m in members}) != 1:
            raise ModelValidationError('A CUBE set must contain members from one hierarchy')
        if len({m.member_key for m in members}) != len(members):
            raise ModelValidationError('Duplicate CUBE set member')
        mdx = '{' + ','.join(m.member_key for m in members) + '}'
        return self._formula('CUBESET', [_literal(self.connection), _literal(mdx), _literal(caption)], 'set',
                             hierarchy=members[0].hierarchy)

    def ranked_member(self, members, rank, *, caption=None):
        if isinstance(rank, bool) or not isinstance(rank, int) or rank < 1:
            raise ModelValidationError('CUBE rank must be a positive integer')
        args = [_literal(self.connection), self._argument(members, ('set',)), str(rank)]
        if caption is not None:
            args.append(_literal(caption))
        return self._formula('CUBERANKEDMEMBER', args, 'member')

    def set_count(self, members):
        return self._formula('CUBESETCOUNT', [self._argument(members, ('set',))], 'value')

    def member_property(self, member, property_name):
        raise ModelValidationError(
            'CUBEMEMBERPROPERTY is not supported for the embedded Excel Data Model; '
            'use CUBEMEMBER captions instead')

    def write(self, worksheet, cell, formula, cell_format=None):
        """Write with an explicit unavailable cache, never a fabricated zero.

        Excel replaces #N/A when it calculates. Non-calculating readers see #N/A.
        """
        CubeCell(cell)
        if not isinstance(formula, CubeFormula) or formula.connection != self.connection:
            raise ModelValidationError('Expected a CUBE formula for this model connection')
        status = worksheet.write_formula(cell, str(formula), cell_format, '#N/A')
        if status != 0:
            raise ModelValidationError(f'Could not write CUBE formula at {cell}: {status}')
        return status
