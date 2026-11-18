"""Locate measure statements without interpreting or evaluating DAX."""
import re
from dataclasses import dataclass
from .abf import UnsupportedModelError

HEADER = re.compile(r"CREATE\s+MEASURE\s+'(?P<table>(?:''|[^'])+)'\[(?P<name>(?:\]\]|[^\]])+)\]\s*=\s*", re.I)


def referenced_identifiers(text):
    """Yield bracket identifiers, excluding quoted text and comments.

    This is conservative dependency detection, not DAX name resolution.
    Escaped closing brackets are decoded for comparison with public names.
    """
    for start, end in protected_spans(text):
        if text[start] == '[':
            yield text[start + 1:end - 1].replace(']]', ']')


@dataclass(frozen=True)
class FormulaReference:
    """A lexical reference and its exact source locations, not an evaluated name."""
    name: str
    table: str | None
    start: int
    end: int
    name_span: tuple[int, int]
    table_span: tuple[int, int] | None = None
    kind: str = 'field'


def formula_references(text, *, table_names=()):
    """Read DAX field/table references without touching strings or comments.

    Bare identifiers are tables only when present in ``table_names``. Names of
    variables shadowing tables are excluded. Field references remain unresolved:
    callers must distinguish measures from local columns using model metadata.
    """
    protected = dict(protected_spans(text))
    tokens = []
    i = 0
    while i < len(text):
        if i in protected:
            end = protected[i]
            if text[i] == '[':
                tokens.append(('field', text[i + 1:end - 1].replace(']]', ']'), i, end))
            elif text[i] == "'":
                tokens.append(('quoted', text[i + 1:end - 1].replace("''", "'"), i, end))
            elif text[i] == '"':
                tokens.append(('literal', '', i, end))
            # Comments are trivia between a qualifier and field.
            i = end
        elif text[i].isalpha() or text[i] == '_':
            start = i
            i += 1
            while i < len(text) and (text[i].isalnum() or text[i] == '_'):
                i += 1
            tokens.append(('word', text[start:i], start, i))
        elif text[i].isspace():
            i += 1
        else:
            tokens.append(('symbol', text[i], i, i + 1))
            i += 1
    tables = {name.casefold() for name in table_names}
    variables = {tokens[j + 1][1].casefold() for j, token in enumerate(tokens[:-1])
                 if token[0] == 'word' and token[1].casefold() == 'var'
                 and tokens[j + 1][0] == 'word'}
    consumed = set()
    result = []
    for j, (kind, name, start, end) in enumerate(tokens):
        if kind != 'field':
            continue
        qualifier = tokens[j - 1] if j else None
        if qualifier and qualifier[0] in ('quoted', 'word') and (
                qualifier[0] == 'quoted' or qualifier[1].casefold() in tables):
            consumed.add(j - 1)
            result.append(FormulaReference(name, qualifier[1], qualifier[2], end,
                                           (start, end), (qualifier[2], qualifier[3])))
        else:
            result.append(FormulaReference(name, None, start, end, (start, end)))
    for j, (kind, name, start, end) in enumerate(tokens):
        if j in consumed or kind not in ('quoted', 'word'):
            continue
        if kind == 'word' and (name.casefold() not in tables or name.casefold() in variables):
            continue
        if j + 1 < len(tokens) and tokens[j + 1][1] == '(':
            continue
        result.append(FormulaReference(name, None, start, end, (start, end), kind='table'))
    return tuple(sorted(result, key=lambda ref: ref.start))


def rewrite_references(text, *, table_names=(), table_renames=None, field_renames=None):
    """Rewrite only explicitly bound lexical names; preserve all other bytes.

    ``field_renames`` maps (qualifier-or-None, field) to a new field name.
    Unqualified fields are rewritten only when explicitly provided by the caller,
    which is responsible for resolving their meaning before using this helper.
    """
    table_map = {key.casefold(): value for key, value in (table_renames or {}).items()}
    field_map = {((table.casefold() if table is not None else None), name.casefold()): value
                 for (table, name), value in (field_renames or {}).items()}
    edits = {}
    for ref in formula_references(text, table_names=table_names):
        if ref.kind == 'table':
            if ref.name.casefold() in table_map:
                edits[ref.name_span] = "'" + table_map[ref.name.casefold()].replace("'", "''") + "'"
        else:
            if ref.table is not None and ref.table.casefold() in table_map:
                edits[ref.table_span] = "'" + table_map[ref.table.casefold()].replace("'", "''") + "'"
            key = (ref.table.casefold() if ref.table else None, ref.name.casefold())
            if key in field_map:
                edits[ref.name_span] = '[' + field_map[key].replace(']', ']]') + ']'
    for (start, end), replacement in sorted(edits.items(), reverse=True):
        text = text[:start] + replacement + text[end:]
    return text


def protected_spans(text):
    spans = []
    i = 0
    while i < len(text):
        start = i
        if text.startswith(('--', '//'), i):
            end = text.find('\n', i)
            i = len(text) if end < 0 else end
        elif text.startswith('/*', i):
            end = text.find('*/', i + 2)
            if end < 0:
                raise UnsupportedModelError('Unterminated formula comment')
            i = end + 2
        elif text[i] in ('"', "'", '['):
            close = ']' if text[i] == '[' else text[i]
            i += 1
            while i < len(text):
                if text[i] == close:
                    if i + 1 < len(text) and text[i + 1] == close:
                        i += 2
                        continue
                    i += 1
                    break
                i += 1
            else:
                raise UnsupportedModelError('Unterminated formula string or identifier')
        else:
            i += 1
            continue
        spans.append((start, i))
    return spans


@dataclass
class Definition:
    text: str
    spans: dict

    def start(self, group=0):
        return self.spans[group][0]

    def end(self, group=0):
        return self.spans[group][1]

    def group(self, group=0):
        return self.text[self.start(group):self.end(group)]


def iter_measure_definitions(text):
    masked = list(text)
    protected = protected_spans(text)
    for start, end in protected:
        masked[start:end] = ' ' * (end - start)
    masked = ''.join(masked)
    cursor = 0
    for header in HEADER.finditer(text):
        if header.start() < cursor or any(a <= header.start() < b for a, b in protected):
            continue
        end = masked.find(';', header.end())
        if end < 0:
            raise UnsupportedModelError('Measure statement has no terminator')
        yield Definition(text, {0:(header.start(), end + 1),
            'table':header.span('table'), 'name':header.span('name'),
            'expression':(header.end(), end)})
        cursor = end + 1


def normalize_expression(expression):
    if not isinstance(expression, str) or not expression.strip():
        raise ValueError('Measure expression must be nonempty')
    text = expression.strip()
    masked = list(text)
    for start, end in protected_spans(text):
        masked[start:end] = ' ' * (end - start)
    separators = [i for i, char in enumerate(masked) if char == ';']
    if separators:
        if separators != [len(text)-1]:
            raise ValueError('Measure expression must contain one expression, not multiple statements')
        text = text[:-1].rstrip()
    if not text:
        raise ValueError('Measure expression must be nonempty')
    # A final line comment must not consume the statement terminator we append.
    return text
