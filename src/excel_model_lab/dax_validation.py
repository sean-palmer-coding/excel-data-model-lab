"""Small schema-bound expression surface certified by the optional adapter.

This is deliberately not a full DAX parser. Unrecognized syntax and functions
are rejected before the third-party evaluator can degrade them to BLANK.
"""
from dataclasses import dataclass
import re


@dataclass(frozen=True)
class Node:
    kind: str
    value: object = None
    args: tuple = ()


def _tokens(text):
    from .dax_evaluation import DaxUnsupportedError
    out = []
    index = 0
    while index < len(text):
        if text[index].isspace():
            index += 1
            continue
        if text.startswith(('--', '//'), index):
            end = text.find('\n', index)
            index = len(text) if end < 0 else end
            continue
        if text.startswith('/*', index):
            end = text.find('*/', index + 2)
            if end < 0:
                raise DaxUnsupportedError('Unterminated DAX comment')
            index = end + 2
            continue
        character = text[index]
        if character in "'\"[":
            closer = ']' if character == '[' else character
            kind = {'[': 'field', "'": 'quoted', '"': 'text'}[character]
            value = ''
            index += 1
            while index < len(text):
                if text[index] == closer:
                    if index + 1 < len(text) and text[index + 1] == closer:
                        value += closer
                        index += 2
                    else:
                        index += 1
                        break
                else:
                    value += text[index]
                    index += 1
            else:
                raise DaxUnsupportedError('Unterminated DAX literal or identifier')
            out.append((kind, value))
            continue
        number = re.match(r'(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?', text[index:])
        if number:
            value = number.group()
            out.append(('number', float(value) if any(c in value for c in '.eE') else int(value)))
            index += len(value)
            continue
        identifier = re.match(r'[^\W\d]\w*', text[index:], re.UNICODE)
        if identifier:
            out.append(('identifier', identifier.group()))
            index += len(identifier.group())
            continue
        operator = next((v for v in ('&&', '||', '<=', '>=', '<>', '==') if text.startswith(v, index)), None)
        if operator:
            if operator == '==':
                raise DaxUnsupportedError('Strict == is outside the verified legacy Excel DAX surface')
            out.append(('symbol', operator))
            index += len(operator)
        elif character in '(),+-*/^&=<>':
            out.append(('symbol', character))
            index += 1
        else:
            raise DaxUnsupportedError(f'Unsupported DAX syntax near {text[index:index + 20]!r}')
    return out + [('end', '')]


class Parser:
    _precedence = {'||': 1, '&&': 2, '=': 3, '<>': 3, '<': 3, '>': 3,
                   '<=': 3, '>=': 3, '&': 4, '+': 5, '-': 5, '*': 6, '/': 6, '^': 7}

    def __init__(self, text):
        self.tokens = _tokens(text)
        self.index = 0

    def take(self):
        token = self.tokens[self.index]
        self.index += 1
        return token

    def parse(self, minimum=0):
        from .dax_evaluation import DaxUnsupportedError
        kind, value = self.take()
        if value in ('+', '-') and kind == 'symbol':
            left = Node('unary', value, (self.parse(8),))
        elif value == '(' and kind == 'symbol':
            left = self.parse()
            if self.take() != ('symbol', ')'):
                raise DaxUnsupportedError('Expected closing DAX parenthesis')
        elif kind in ('number', 'text'):
            left = Node(kind, value)
        elif kind == 'field':
            left = Node('reference', (None, value))
        elif kind in ('identifier', 'quoted'):
            if self.tokens[self.index][0] == 'field':
                left = Node('reference', (value, self.take()[1]))
            elif self.tokens[self.index] == ('symbol', '(') and kind == 'identifier':
                self.take()
                args = []
                if self.tokens[self.index] != ('symbol', ')'):
                    while True:
                        args.append(self.parse())
                        if self.tokens[self.index] != ('symbol', ','):
                            break
                        self.take()
                if self.take() != ('symbol', ')'):
                    raise DaxUnsupportedError('Expected closing DAX function parenthesis')
                left = Node('call', value.upper(), tuple(args))
            elif kind == 'identifier' and value.upper() in ('TRUE', 'FALSE'):
                left = Node('boolean', value.upper() == 'TRUE')
            else:
                left = Node('table', value)
        else:
            raise DaxUnsupportedError('Unsupported or incomplete scalar DAX expression')
        # Only symbol tokens are operators; "-" or [+] are a string and a field.
        while self.tokens[self.index][0] == 'symbol' and self.tokens[self.index][1] in self._precedence:
            operation = self.tokens[self.index][1]
            priority = self._precedence[operation]
            if priority < minimum:
                break
            self.take()
            left = Node('binary', operation, (left, self.parse(priority + 1)))
        return left

    def expression(self):
        from .dax_evaluation import DaxUnsupportedError
        result = self.parse()
        if self.tokens[self.index][0] != 'end':
            raise DaxUnsupportedError('Expression shape is outside the verified scalar DAX subset')
        return result


def validate_measure(snapshot, name):
    """Bind the supported AST against current tables, columns and measures."""
    from .dax_evaluation import DaxEvaluationError, DaxUnsupportedError
    table_names = {n.casefold(): n for n in snapshot.tables}
    measures = {n.casefold(): n for n in snapshot.measures}
    stack = set()
    memo = {}

    def table_name(name):
        try:
            return table_names[name.casefold()]
        except KeyError as exc:
            raise DaxEvaluationError(f'Unknown DAX table: {name}') from exc

    def measure(name):
        name = measures.get(name.casefold(), name)
        if name in stack:
            raise DaxEvaluationError(f'Circular measure reference: {name}')
        if name not in snapshot.measures:
            raise DaxEvaluationError(f'Unknown DAX measure: {name}')
        if name not in memo:
            stack.add(name)
            memo[name] = infer(Parser(snapshot.measures[name]).expression(), snapshot.measure_tables[name], None)
            stack.remove(name)
        return memo[name]

    def column(node, home, row):
        if node.kind != 'reference':
            raise DaxUnsupportedError('This aggregate requires a direct model column reference')
        qualifier, name = node.value
        table = table_name(qualifier) if qualifier is not None else (row or home)
        columns = {c.casefold(): c for c in snapshot.tables[table]['columns']}
        if name.casefold() not in columns:
            raise DaxEvaluationError(f'Unknown DAX column: {table}[{name}]')
        name = columns[name.casefold()]
        raw = snapshot.column_types[(table, name)].lower()
        if raw in ('int64', 'bigint', 'integer', 'float64', 'double', 'decimal'):
            kind = 'number'
        elif raw in ('bool', 'boolean'):
            kind = 'boolean'
        elif raw.startswith(('date', 'time')):
            kind = 'datetime'
        elif raw in ('string', 'text', 'wchar', 'object'):
            kind = 'text'
        else:
            raise DaxUnsupportedError(f'Unverified DAX column type: {raw}')
        return kind, table

    def numeric(kind):
        if kind not in ('number', 'blank'):
            raise DaxUnsupportedError('This numeric expression/type combination is not verified')

    def infer(node, home, row):
        if node.kind in ('number', 'text', 'boolean'):
            return node.kind
        if node.kind == 'table':
            return ('table', table_name(node.value))
        if node.kind == 'reference':
            qualifier, name = node.value
            if name.casefold() in measures:
                owner = snapshot.measure_tables[measures[name.casefold()]]
                if qualifier is None or table_name(qualifier) == owner:
                    return measure(name)
            kind, owner = column(node, home, row)
            if row is None:
                raise DaxUnsupportedError('Bare column scalar references require a verified iterator row context')
            if owner != row:
                raise DaxUnsupportedError('Cross-table row references require RELATED')
            return kind
        if node.kind == 'unary':
            numeric(infer(node.args[0], home, row))
            return 'number'
        if node.kind == 'binary':
            left, right = (infer(arg, home, row) for arg in node.args)
            if node.value in ('+', '-', '*', '/', '^'):
                numeric(left); numeric(right)
                return 'number'
            if node.value in ('&&', '||'):
                if left != 'boolean' or right != 'boolean':
                    raise DaxUnsupportedError('Logical operations require Boolean expressions')
                return 'boolean'
            if node.value == '&':
                raise DaxUnsupportedError('Concatenation is outside the verified scalar DAX subset')
            if left != right and 'blank' not in (left, right):
                raise DaxEvaluationError('Incompatible DAX comparison operand types')
            return 'boolean'
        function, args = node.value, node.args
        counts = {'SUM': (1,), 'COUNT': (1,), 'COUNTROWS': (1,), 'COUNTBLANK': (1,),
                  'DISTINCTCOUNT': (1,), 'SUMX': (2,), 'MAXX': (2,), 'FILTER': (2,),
                  'ALL': (1,), 'VALUES': (1,), 'RELATED': (1,), 'DIVIDE': (2, 3),
                  'BLANK': (0,), 'TRUE': (0,), 'FALSE': (0,), 'IF': (2, 3), 'DATE': (3,)}
        if function != 'CALCULATE' and function not in counts:
            raise DaxUnsupportedError(f'{function} is outside the verified scalar DAX subset')
        if (function == 'CALCULATE' and not args) or (function in counts and len(args) not in counts[function]):
            raise DaxEvaluationError(f'Invalid argument count for {function}')
        if function == 'BLANK': return 'blank'
        if function in ('TRUE', 'FALSE'): return 'boolean'
        if function in ('SUM', 'COUNT', 'COUNTBLANK', 'DISTINCTCOUNT'):
            kind, _ = column(args[0], home, row)
            if function == 'SUM': numeric(kind)
            if function == 'COUNT' and kind == 'boolean':
                raise DaxEvaluationError('COUNT does not support Boolean model columns')
            return 'number'
        if function in ('ALL', 'VALUES'):
            if args[0].kind == 'reference':
                _, owner = column(args[0], home, row)
                return ('table', owner)
            result = infer(args[0], home, row)
            if not isinstance(result, tuple): raise DaxEvaluationError(f'{function} requires a table or column')
            return result
        if function in ('COUNTROWS', 'FILTER', 'SUMX', 'MAXX'):
            source = infer(args[0], home, row)
            if not isinstance(source, tuple): raise DaxEvaluationError(f'{function} requires a table')
            if function == 'COUNTROWS': return 'number'
            kind = infer(args[1], home, source[1])
            if function == 'FILTER':
                if kind != 'boolean': raise DaxUnsupportedError('FILTER requires a Boolean predicate')
                return source
            numeric(kind)
            return 'number'
        if function == 'RELATED':
            kind, owner = column(args[0], home, row)
            if row is None: raise DaxEvaluationError('RELATED requires row context')
            if not any(r['FromTable'] == row and r['ToTable'] == owner and r['IsActive'] for r in snapshot.relationships):
                raise DaxUnsupportedError('RELATED requires a verified direct active relationship')
            return kind
        if function == 'CALCULATE':
            for arg in args[1:]:
                # Boolean filters may reference model columns without row context.
                if arg.kind == 'binary' and arg.value in ('=', '<>', '<', '>', '<=', '>='):
                    refs = [a for a in arg.args if a.kind == 'reference']
                    filter_row = column(refs[0], home, row)[1] if refs else row
                    kind = infer(arg, home, filter_row)
                else:
                    kind = infer(arg, home, row)
                if kind != 'boolean' and not isinstance(kind, tuple):
                    raise DaxUnsupportedError('CALCULATE filter shape is not verified')
            return infer(args[0], home, row)
        if function in ('DIVIDE', 'DATE'):
            for arg in args[:2] if function == 'DIVIDE' else args:
                numeric(infer(arg, home, row))
            if function == 'DIVIDE' and len(args) == 3:
                alternate = args[2]
                if alternate.kind != 'number' and not (alternate.kind == 'unary' and alternate.args[0].kind == 'number'):
                    raise DaxUnsupportedError('DIVIDE alternate result must be a numeric constant')
            return 'datetime' if function == 'DATE' else 'number'
        if function == 'IF':
            if infer(args[0], home, row) != 'boolean': raise DaxUnsupportedError('IF requires a Boolean condition')
            kinds = [infer(arg, home, row) for arg in args[1:]]
            nonblank = {kind for kind in kinds if kind != 'blank'}
            if len(nonblank) > 1: raise DaxUnsupportedError('Mixed IF result types are not verified')
            return next(iter(nonblank), 'blank')
        raise DaxUnsupportedError(f'Unsupported DAX shape: {function}')

    result = measure(name)
    if isinstance(result, tuple):
        raise DaxUnsupportedError('Only scalar measure expressions are supported')
    _integer_bounds(snapshot, name)
    return result


def _integer_bounds(snapshot, measure_name):
    """Conservatively rule out signed-Int64 overflow, including intermediates.

    Bounds span all rows, even when a filter could narrow them. Refusing a wide
    range is safer than the backend's unbounded Python integer arithmetic.
    """
    from .dax_evaluation import DaxUnsupportedError
    tables = {name.casefold(): name for name in snapshot.tables}
    measures = {name.casefold(): name for name in snapshot.measures}
    memo, columns = {}, {}

    def checked(value):
        if value is not None and (value[0] < -(2**63) or value[1] >= 2**63):
            raise DaxUnsupportedError('Possible intermediate integer overflow; verified evaluation requires signed Int64 bounds')
        return value

    def data(node, home, row):
        table, column = node.value
        table = tables[table.casefold()] if table else (row or home)
        column = next(c for c in snapshot.tables[table]['columns'] if c.casefold() == column.casefold())
        key = table, column
        if key not in columns:
            dtype = snapshot.column_types[key].lower()
            if dtype not in ('int64', 'integer', 'bigint'):
                columns[key] = None
            else:
                index = snapshot.tables[table]['columns'].index(column)
                values = [r[index] for r in snapshot.tables[table]['rows'] if r[index] is not None]
                columns[key] = values
        return columns[key]

    def owner(node, home, row):
        if node.kind == 'table': return tables[node.value.casefold()]
        if node.kind == 'reference': return tables[node.value[0].casefold()] if node.value[0] else (row or home)
        return owner(node.args[0], home, row)

    def measure(name):
        name = measures[name.casefold()]
        if name not in memo:
            memo[name] = visit(Parser(snapshot.measures[name]).expression(), snapshot.measure_tables[name], None)
        return memo[name]

    def visit(node, home, row):
        if node.kind == 'number':
            return checked((node.value, node.value)) if isinstance(node.value, int) else None
        if node.kind in ('text', 'boolean', 'table'): return None
        if node.kind == 'reference':
            table, name = node.value
            if name.casefold() in measures and (table is None or tables[table.casefold()] == snapshot.measure_tables[measures[name.casefold()]]):
                return measure(name)
            values = data(node, home, row)
            return checked((min([0, *values]), max([0, *values]))) if values is not None else None
        if node.kind == 'unary':
            value = visit(node.args[0], home, row)
            return checked((-value[1], -value[0]) if node.value == '-' else value) if value else None
        if node.kind == 'binary':
            a, b = (visit(arg, home, row) for arg in node.args)
            if a is None or b is None: return None
            if node.value == '+': return checked((a[0] + b[0], a[1] + b[1]))
            if node.value == '-': return checked((a[0] - b[1], a[1] - b[0]))
            if node.value == '*':
                products = [x * y for x in a for y in b]
                return checked((min(products), max(products)))
            return None
        function, args = node.value, node.args
        if function == 'BLANK': return (0, 0)
        if function == 'SUM':
            values = data(args[0], home, row)
            return checked((sum(v for v in values if v < 0), sum(v for v in values if v > 0))) if values is not None else None
        if function in ('COUNT', 'COUNTBLANK', 'DISTINCTCOUNT', 'COUNTROWS'):
            if function == 'COUNTROWS': visit(args[0], home, row)
            return checked((0, len(snapshot.tables[owner(args[0], home, row)]['rows']) + 1))
        if function in ('SUMX', 'MAXX', 'FILTER'):
            visit(args[0], home, row)
            table = owner(args[0], home, row)
            value = visit(args[1], home, table)
            if function == 'FILTER': return None
            if function == 'MAXX' or value is None: return value
            count = len(snapshot.tables[table]['rows'])
            return checked((min(0, value[0] * count), max(0, value[1] * count)))
        if function == 'RELATED':
            values = data(args[0], home, row)
            return checked((min([0, *values]), max([0, *values]))) if values is not None else None
        if function == 'CALCULATE':
            for arg in args[1:]: visit(arg, home, row)
            return visit(args[0], home, row)
        if function == 'IF':
            visit(args[0], home, row)
            values = [visit(arg, home, row) for arg in args[1:]]
            if any(value is None for value in values): return None
            return checked((min(v[0] for v in values), max(v[1] for v in values)))
        for arg in args: visit(arg, home, row)
        return None

    measure(measure_name)
