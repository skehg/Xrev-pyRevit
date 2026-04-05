# -*- coding: utf-8 -*-
"""
formula_analyzer.py
-------------------
Pure IronPython formula analysis library.  No Revit API or WPF imports.

Public API
----------
analyze_formula(formula_str, param_names=None)
    -> dict: {ok, simplified, suggestions, metrics}
       ok          : bool
       simplified  : str  (may equal original if no simplification found)
       suggestions : list[str]
       metrics     : dict {nodes, depth, unique_params}
       error       : str  (only present when ok=False)
"""

import re
from collections import defaultdict

# ---------------------------------------------------------------------------
# Tokenizer
# ---------------------------------------------------------------------------
# Token types: NUM, NAME, OP, LP, RP, COMMA
_TOK_RE = re.compile(
    r'\s*(?:'
    r'(\d+(?:\.\d+)?(?:[eE][+-]?\d+)?)'   # group 1 – number
    r'|([A-Za-z_]\w*)'                      # group 2 – identifier/keyword
    r'|(==|!=|<=|>=|\*\*|[+\-*/^%<>()=]|,)' # group 3 – operator/paren/comma (= is Revit equality)
    r')'
)


def _tokenize(s):
    """Return list of (type, value) tokens.  Raises ValueError on bad input."""
    pos = 0
    tokens = []
    s_len = len(s)
    while pos < s_len:
        m = _TOK_RE.match(s, pos)
        if not m or m.start() == m.end():
            ch = s[pos]
            if ch.strip() == '':
                pos += 1
                continue
            raise ValueError("Unexpected character at position {}: {!r}".format(pos, ch))
        num, name, op_or_paren = m.group(1), m.group(2), m.group(3)
        if num is not None:
            tokens.append(('NUM', float(num)))
        elif name is not None:
            tokens.append(('NAME', name))
        elif op_or_paren is not None:
            ch = op_or_paren
            if ch == '(':
                tokens.append(('LP', ch))
            elif ch == ')':
                tokens.append(('RP', ch))
            elif ch == ',':
                tokens.append(('COMMA', ch))
            else:
                tokens.append(('OP', ch))
        pos = m.end()
    return tokens


# ---------------------------------------------------------------------------
# Operator precedence
# ---------------------------------------------------------------------------
_PREC = {'^': 4, '**': 4, '*': 3, '/': 3, '%': 3, '+': 2, '-': 2,
         '<': 1, '>': 1, '<=': 1, '>=': 1, '==': 1, '!=': 1, '=': 1}
_RIGHT_ASSOC = set(['^', '**'])


# ---------------------------------------------------------------------------
# AST node constructors
# ---------------------------------------------------------------------------
def _const(v):          return ('const', v)
def _sym(n):            return ('sym', n)
def _op(o, a, b):       return ('op', o, a, b)
def _neg(a):            return ('neg', a)
def _fn(name, args):    return ('fn', name, args)


def _collect_additive_terms(node, sign=1):
    """
    Flatten a +/- chain into (term_key, sign) pairs so that any
    permutation/association of the same terms produces an identical
    sorted list.  Used by _node_key for additive normalisation.
    """
    if node[0] == 'op' and node[1] == '+':
        return (_collect_additive_terms(node[2], sign)
                + _collect_additive_terms(node[3], sign))
    if node[0] == 'op' and node[1] == '-':
        return (_collect_additive_terms(node[2], sign)
                + _collect_additive_terms(node[3], -sign))
    if node[0] == 'neg':
        return _collect_additive_terms(node[1], -sign)
    # Non-additive leaf of the chain
    return [(_node_key(node), sign)]


def _node_key(node):
    """Canonical string key for comparison and deduplication.

    Normalisations applied:
    - +/- chains: flattened to a sorted multiset of signed terms, so
      (a+b-c), (b-c+a), (a-c+b), etc. all share one key.
    - Division by a literal constant: treated as multiplication by its
      reciprocal, so (a/2), (a*0.5), (0.5*a) all share one key.
    """
    t = node[0]
    if t == 'const':
        return 'C{}'.format(node[1])
    if t == 'sym':
        return 'S{}'.format(node[1])
    if t == 'op':
        op = node[1]
        a, b = node[2], node[3]

        # Additive normalisation — flatten entire +/- chain and sort
        if op in ('+', '-'):
            terms = _collect_additive_terms(node)
            terms.sort()
            return 'SUM({})'.format(
                ','.join('{}{}'.format('+' if s > 0 else '-', k)
                         for k, s in terms)
            )

        # Division-by-constant: a/k is the same as a*(1/k)
        if op == '/' and b[0] == 'const' and b[1] != 0:
            recip = 1.0 / b[1]
            ka, kb = _node_key(a), 'C{}'.format(recip)
            if ka > kb:
                ka, kb = kb, ka
            return 'O*({},{})'.format(ka, kb)

        return 'O{}({},{})'.format(op, _node_key(a), _node_key(b))
    if t == 'neg':
        return 'N({})'.format(_node_key(node[1]))
    if t == 'fn':
        return 'F{}({})'.format(node[1], ','.join(_node_key(a) for a in node[2]))
    return '?'


def _is_trivial_key(key):
    """True for bare constant/symbol leaves — not worth reporting as subexpressions.

    Additive compound keys start with 'SUM(' and must NOT be classed as
    trivial even though they start with 'S'.
    """
    if key.startswith('C'):
        return True
    if key.startswith('S') and not key.startswith('SUM('):
        return True
    return False


# ---------------------------------------------------------------------------
# Simplification rules
# ---------------------------------------------------------------------------
def _apply_op_const(op, a, b):
    if op == '+':  return a + b
    if op == '-':  return a - b
    if op == '*':  return a * b
    if op == '/':
        if b == 0:
            raise ZeroDivisionError("Division by zero in constant subexpression")
        return a / b
    if op in ('^', '**'): return a ** b
    if op == '%':
        if b == 0:
            raise ZeroDivisionError("Modulo by zero in constant subexpression")
        return a % b
    return None


def _simplify(op, a, b):
    """Build a simplified AST node for (a op b)."""
    # Canonicalise commutative pairs so canonical form is deterministic
    if op in ('+', '*'):
        if _node_key(a) > _node_key(b):
            a, b = b, a

    # Constant folding
    if a[0] == 'const' and b[0] == 'const':
        try:
            result = _apply_op_const(op, a[1], b[1])
            if result is not None:
                return _const(result)
        except ZeroDivisionError:
            pass  # keep as-is so it surfaces in suggestions

    # Static division-by-zero flag (constant denominator)
    if op == '/' and b[0] == 'const' and b[1] == 0:
        return _op(op, a, b)  # kept; flagged separately

    # Additive identities
    if op == '+':
        if a[0] == 'const' and a[1] == 0: return b
        if b[0] == 'const' and b[1] == 0: return a
    if op == '-':
        if b[0] == 'const' and b[1] == 0: return a
        if _node_key(a) == _node_key(b): return _const(0)

    # Multiplicative identities / annihilators
    if op == '*':
        if a[0] == 'const' and a[1] == 1: return b
        if b[0] == 'const' and b[1] == 1: return a
        if a[0] == 'const' and a[1] == 0: return _const(0)
        if b[0] == 'const' and b[1] == 0: return _const(0)

    # Division identities
    if op == '/':
        if b[0] == 'const' and b[1] == 1: return a
        if _node_key(a) == _node_key(b): return _const(1)

    # Power identities
    if op in ('^', '**'):
        if b[0] == 'const' and b[1] == 0: return _const(1)
        if b[0] == 'const' and b[1] == 1: return a
        if a[0] == 'const' and a[1] == 1: return _const(1)

    # Distributive factoring: (a * x) ± (b * x) → (a ± b) * x
    # Checks all four pairings of left/right operands from each multiplication.
    if op in ('+', '-'):
        if a[0] == 'op' and a[1] == '*' and b[0] == 'op' and b[1] == '*':
            al, ar = a[2], a[3]
            bl, br = b[2], b[3]
            if _node_key(al) == _node_key(bl):   # shared left factor
                return _simplify('*', _simplify(op, ar, br), al)
            if _node_key(ar) == _node_key(br):   # shared right factor
                return _simplify('*', _simplify(op, al, bl), ar)
            if _node_key(al) == _node_key(br):   # left of a == right of b
                return _simplify('*', _simplify(op, ar, bl), al)
            if _node_key(ar) == _node_key(bl):   # right of a == left of b
                return _simplify('*', _simplify(op, al, br), ar)

    return _op(op, a, b)


# ---------------------------------------------------------------------------
# Boolean / function simplification
# ---------------------------------------------------------------------------
def _is_truthy(node):
    """Return True if node is a non-zero constant (boolean true in Revit)."""
    return node[0] == 'const' and node[1] != 0


def _is_falsy(node):
    """Return True if node is a zero constant (boolean false in Revit)."""
    return node[0] == 'const' and node[1] == 0


def _simplify_fn(name, args):
    """Apply simplification rules to a function-call AST node."""
    lname = name.lower()

    # ── not() ────────────────────────────────────────────────────────────────────
    if lname == 'not' and len(args) == 1:
        a = args[0]
        if _is_truthy(a): return _const(0)
        if _is_falsy(a):  return _const(1)
        # not(not(x)) → x
        if a[0] == 'fn' and a[1] == 'not' and len(a[2]) == 1:
            return a[2][0]

    # ── and() ──────────────────────────────────────────────────────────────────
    if lname == 'and' and len(args) == 2:
        a, b = args
        if _is_falsy(a) or _is_falsy(b): return _const(0)
        if _is_truthy(a): return b
        if _is_truthy(b): return a
        if _node_key(a) == _node_key(b): return a

    # ── or() ───────────────────────────────────────────────────────────────────
    if lname == 'or' and len(args) == 2:
        a, b = args
        if _is_truthy(a) or _is_truthy(b): return _const(1)
        if _is_falsy(a): return b
        if _is_falsy(b): return a
        if _node_key(a) == _node_key(b): return a

    # ── if() ───────────────────────────────────────────────────────────────────
    if lname == 'if' and len(args) == 3:
        cond, tv, fv = args
        # Constant condition
        if _is_truthy(cond): return tv
        if _is_falsy(cond):  return fv
        # Both branches identical → drop the condition entirely
        if _node_key(tv) == _node_key(fv): return tv
        # Normalise if(not(c), a, b) → if(c, b, a)
        if cond[0] == 'fn' and cond[1] == 'not' and len(cond[2]) == 1:
            return _simplify_fn('if', [cond[2][0], fv, tv])
        # Redundant nested: if(c, if(c, a, b), d) → if(c, a, d)
        if tv[0] == 'fn' and tv[1] == 'if' and len(tv[2]) == 3:
            ic, ia, _ = tv[2]
            if _node_key(cond) == _node_key(ic):
                return _simplify_fn('if', [cond, ia, fv])

    # ── roundup / rounddown / round of integer constant ──────────────────────────
    if lname in ('roundup', 'rounddown', 'round') and len(args) == 1:
        a = args[0]
        if a[0] == 'const' and float(a[1]) == int(a[1]):
            return a

    # ── sqrt of non-negative constant ───────────────────────────────────────
    if lname == 'sqrt' and len(args) == 1:
        a = args[0]
        if a[0] == 'const' and a[1] >= 0:
            return _const(a[1] ** 0.5)

    return _fn(lname, args)


# ---------------------------------------------------------------------------
# Recursive descent parser
# ---------------------------------------------------------------------------
class _Parser(object):
    def __init__(self, tokens):
        self._tok = tokens
        self._pos = 0

    def _peek(self):
        if self._pos < len(self._tok):
            return self._tok[self._pos]
        return ('EOF', None)

    def _consume(self):
        if self._pos >= len(self._tok):
            raise ValueError(
                "Unexpected end of formula — check for missing closing parenthesis"
            )
        t = self._tok[self._pos]
        self._pos += 1
        return t

    def _expect(self, ttype):
        t = self._consume()
        if t[0] != ttype:
            raise ValueError("Expected {} but got {} ({!r})".format(ttype, t[0], t[1]))
        return t

    def parse(self):
        node = self._expr(0)
        if self._pos < len(self._tok):
            raise ValueError("Unexpected token: {!r}".format(self._peek()[1]))
        return node

    def _expr(self, min_prec):
        left = self._unary()
        while True:
            tok = self._peek()
            if tok[0] != 'OP':
                break
            op = tok[1]
            prec = _PREC.get(op, 0)
            if prec == 0 or prec < min_prec:
                break
            self._consume()
            next_prec = prec if op in _RIGHT_ASSOC else prec + 1
            right = self._expr(next_prec)
            left = _simplify(op, left, right)
        return left

    def _unary(self):
        tok = self._peek()
        if tok[0] == 'OP' and tok[1] == '-':
            self._consume()
            operand = self._unary()
            if operand[0] == 'const':
                return _const(-operand[1])
            return _neg(operand)
        if tok[0] == 'OP' and tok[1] == '+':
            self._consume()
            return self._unary()
        return self._primary()

    def _primary(self):
        tok = self._peek()
        if tok[0] == 'NUM':
            self._consume()
            return _const(tok[1])
        if tok[0] == 'NAME':
            self._consume()
            if self._peek()[0] == 'LP':
                return self._fn_call(tok[1])
            return _sym(tok[1])
        if tok[0] == 'LP':
            self._consume()
            node = self._expr(0)
            self._expect('RP')
            return node
        raise ValueError("Unexpected token: {} {!r}".format(tok[0], tok[1]))

    def _fn_call(self, name):
        self._expect('LP')
        args = []
        if self._peek()[0] != 'RP':
            args.append(self._expr(0))
            while self._peek()[0] == 'COMMA':
                self._consume()
                args.append(self._expr(0))
        self._expect('RP')
        return _simplify_fn(name, args)


# ---------------------------------------------------------------------------
# AST → string (pretty printer)
# ---------------------------------------------------------------------------
# Operator precedence for parenthesisation decisions
_OP_PREC_OUT = {'+': 2, '-': 2, '*': 3, '/': 3, '%': 3, '^': 4, '**': 4,
                '<': 1, '>': 1, '<=': 1, '>=': 1, '==': 1, '!=': 1, '=': 1}
_OP_ASSOC = {'*': 'both', '+': 'both', '-': 'left', '/': 'left',
             '%': 'left', '^': 'right', '**': 'right'}


def _node_str(node, parent_op=None, side='none'):
    t = node[0]
    if t == 'const':
        v = node[1]
        if isinstance(v, float) and v == int(v):
            return str(int(v))
        return repr(v)
    if t == 'sym':
        return node[1]
    if t == 'neg':
        inner = node[1]
        s = _node_str(inner)
        if inner[0] == 'op':
            return '-({})'.format(s)
        return '-{}'.format(s)
    if t == 'op':
        op, a, b = node[1], node[2], node[3]
        a_str = _node_str(a, op, 'left')
        b_str = _node_str(b, op, 'right')
        expr = '{} {} {}'.format(a_str, op, b_str)

        # Add parentheses if needed to preserve semantics
        if parent_op is not None:
            p_cur = _OP_PREC_OUT.get(op, 0)
            p_par = _OP_PREC_OUT.get(parent_op, 0)
            needs_parens = False
            if p_cur < p_par:
                needs_parens = True
            elif p_cur == p_par:
                assoc = _OP_ASSOC.get(parent_op, 'left')
                if side == 'right' and assoc == 'left':
                    needs_parens = True
                if side == 'left' and assoc == 'right':
                    needs_parens = True
            if needs_parens:
                return '({})'.format(expr)
        return expr
    if t == 'fn':
        name, args = node[1], node[2]
        return '{}({})'.format(name, ', '.join(_node_str(a) for a in args))
    return '?'


def node_to_str(node):
    """Convert AST node to a readable formula string."""
    return _node_str(node)


def _replace_subexpr_in_ast(node, rep_key, replacement_sym):
    """Return a new AST with every subtree matching rep_key replaced by replacement_sym."""
    if _node_key(node) == rep_key:
        return replacement_sym
    if node[0] == 'op':
        return _op(node[1],
                   _replace_subexpr_in_ast(node[2], rep_key, replacement_sym),
                   _replace_subexpr_in_ast(node[3], rep_key, replacement_sym))
    if node[0] == 'neg':
        return _neg(_replace_subexpr_in_ast(node[1], rep_key, replacement_sym))
    if node[0] == 'fn':
        return _fn(node[1], [_replace_subexpr_in_ast(a, rep_key, replacement_sym)
                              for a in node[2]])
    return node


def replace_formula_subexpr(formula_str, rep_key, new_param_name, param_names=None):
    """
    Replace every occurrence of the subexpression identified by *rep_key* in
    *formula_str* with *new_param_name*.  Returns the rewritten formula string.
    Raises ValueError on parse failure.
    """
    if param_names is None:
        param_names = set()
    formula_work, name_map = _normalize_param_names(formula_str.strip(), param_names)
    tokens = _tokenize(formula_work)
    ast = _Parser(tokens).parse()
    placeholder = u'__RPARAM_NEW__'
    new_ast = _replace_subexpr_in_ast(ast, rep_key, _sym(placeholder))
    name_map[placeholder] = new_param_name
    return _denormalize(node_to_str(new_ast), name_map)


# ---------------------------------------------------------------------------
# Subexpression analysis
# ---------------------------------------------------------------------------
def _collect_subexprs(node, counts, node_map=None):
    """Recursively count occurrences of each canonical subexpression key."""
    key = _node_key(node)
    counts[key] += 1
    if node_map is not None and key not in node_map:
        node_map[key] = node
    if node[0] == 'op':
        _collect_subexprs(node[2], counts, node_map)
        _collect_subexprs(node[3], counts, node_map)
    elif node[0] == 'neg':
        _collect_subexprs(node[1], counts, node_map)
    elif node[0] == 'fn':
        for a in node[2]:
            _collect_subexprs(a, counts, node_map)


def _find_repeated_subexprs(node):
    """Return list of (key, ast_node, count) for non-trivial subexpressions used more than once."""
    counts = defaultdict(int)
    node_map = {}
    _collect_subexprs(node, counts, node_map)
    repeated = []
    for key, cnt in counts.items():
        if cnt > 1 and not _is_trivial_key(key):
            repeated.append((key, node_map[key], cnt))
    return repeated


def _ast_contains_key(node, key):
    """Return True if any subtree of node shares the given canonical key."""
    if _node_key(node) == key:
        return True
    if node[0] == 'op':
        return _ast_contains_key(node[2], key) or _ast_contains_key(node[3], key)
    if node[0] == 'neg':
        return _ast_contains_key(node[1], key)
    if node[0] == 'fn':
        return any(_ast_contains_key(a, key) for a in node[2])
    return False


# ---------------------------------------------------------------------------
# Division-by-zero detection
# ---------------------------------------------------------------------------
def _find_div_by_zero(node):
    """Return True if any constant /0 or %0 is present in the AST."""
    if node[0] == 'op':
        op = node[1]
        if op in ('/', '%') and node[3][0] == 'const' and node[3][1] == 0:
            return True
        return _find_div_by_zero(node[2]) or _find_div_by_zero(node[3])
    if node[0] == 'neg':
        return _find_div_by_zero(node[1])
    if node[0] == 'fn':
        return any(_find_div_by_zero(a) for a in node[2])
    return False


# ---------------------------------------------------------------------------
# Complexity metrics
# ---------------------------------------------------------------------------
def _complexity_metrics(node):
    """Return dict of node count, max operator depth, and unique param count."""
    stats = {'nodes': 0, 'depth': 0, 'syms': set()}

    def _walk(n, d):
        stats['nodes'] += 1
        if d > stats['depth']:
            stats['depth'] = d
        if n[0] == 'sym':
            stats['syms'].add(n[1])
        elif n[0] == 'op':
            _walk(n[2], d + 1)
            _walk(n[3], d + 1)
        elif n[0] == 'neg':
            _walk(n[1], d + 1)
        elif n[0] == 'fn':
            for a in n[2]:
                _walk(a, d + 1)

    _walk(node, 1)
    return {
        'nodes': stats['nodes'],
        'depth': stats['depth'],
        'unique_params': len(stats['syms']),
    }


# ---------------------------------------------------------------------------
# Token-level diff (LCS)
# ---------------------------------------------------------------------------
def _simple_diff(a, b):
    """
    Return a list of (kind, token) where kind is 'same', 'removed', or 'added'.
    Uses token-level LCS for minimal edit distance.
    """
    toks_a = re.findall(r'\w+|[^\s\w]', a)
    toks_b = re.findall(r'\w+|[^\s\w]', b)
    n, m = len(toks_a), len(toks_b)

    # Build LCS table (suffix form for simpler reconstruction)
    lcs = [[0] * (m + 1) for _ in range(n + 1)]
    for i in range(n - 1, -1, -1):
        for j in range(m - 1, -1, -1):
            if toks_a[i] == toks_b[j]:
                lcs[i][j] = 1 + lcs[i + 1][j + 1]
            else:
                lcs[i][j] = max(lcs[i + 1][j], lcs[i][j + 1])

    out = []
    i = j = 0
    while i < n and j < m:
        if toks_a[i] == toks_b[j]:
            out.append(('same', toks_a[i]))
            i += 1
            j += 1
        elif lcs[i + 1][j] >= lcs[i][j + 1]:
            out.append(('removed', toks_a[i]))
            i += 1
        else:
            out.append(('added', toks_b[j]))
            j += 1
    while i < n:
        out.append(('removed', toks_a[i]))
        i += 1
    while j < m:
        out.append(('added', toks_b[j]))
        j += 1
    return out


# ---------------------------------------------------------------------------
# Multi-word parameter name normalisation
# ---------------------------------------------------------------------------
def _normalize_param_names(formula, param_names):
    """
    Replace multi-word parameter names with safe single-token placeholders so
    the tokeniser doesn't split them into separate NAME tokens.

    Returns (normalised_formula, mapping) where mapping maps placeholder -> original.
    Sorts by length descending so longer names are replaced before shorter ones
    that might be substrings of them.
    """
    if not param_names:
        return formula, {}
    multi_word = sorted(
        [n for n in param_names if ' ' in n],
        key=lambda n: -len(n)
    )
    if not multi_word:
        return formula, {}
    mapping = {}
    result = formula
    for i, name in enumerate(multi_word):
        placeholder = '__RPARAM{}__'.format(i)
        if name in result:
            result = result.replace(name, placeholder)
            mapping[placeholder] = name
    return result, mapping


def _denormalize(text, mapping):
    """Reverse placeholder substitution produced by _normalize_param_names."""
    for placeholder, original in mapping.items():
        text = text.replace(placeholder, original)
    return text


# ---------------------------------------------------------------------------
# Top-level entry point
# ---------------------------------------------------------------------------
def analyze_formula(formula_str, param_names=None, param_formulas=None):
    """
    Analyze and simplify a Revit family parameter formula.

    Parameters
    ----------
    formula_str : str
        The raw formula text.
    param_names : set or None
        Known parameter names from the family.
    param_formulas : dict or None
        Map of {parameter_name: formula_text} for other parameters.
        Used to detect if a repeated subexpression already exists as a parameter.

    Returns
    -------
    dict with keys:
        ok          bool
        simplified  str   (may equal formula_str when no simplification found)
        suggestions list[str]
        metrics     dict {nodes, depth, unique_params}
        diff        list[(kind, token)]
        error       str  (only when ok=False)
    """
    if param_names is None:
        param_names = set()
    if param_formulas is None:
        param_formulas = {}

    formula_str = (formula_str or '').strip()
    if not formula_str:
        return {'ok': False, 'error': 'Formula is empty.'}

    suggestions = []

    # Replace multi-word parameter names with safe single-token placeholders
    formula_work, name_map = _normalize_param_names(formula_str, param_names)

    try:
        tokens = _tokenize(formula_work)
    except ValueError as exc:
        return {'ok': False, 'error': 'Tokenization error: {}'.format(exc)}

    try:
        ast = _Parser(tokens).parse()
    except (ValueError, IndexError) as exc:
        return {'ok': False, 'error': 'Parse error: {}'.format(exc)}
    except Exception as exc:
        return {'ok': False, 'error': 'Unexpected parse error: {}'.format(exc)}

    # Pretty-print the simplified form and reverse placeholder substitution
    try:
        simplified = _denormalize(node_to_str(ast), name_map)
    except Exception as exc:
        return {'ok': False, 'error': 'Stringify error: {}'.format(exc)}

    # --- Suggestions ---

    # Constant result
    if ast[0] == 'const':
        suggestions.append(
            u"Expression always evaluates to the constant {}.".format(simplified)
        )

    # Simplification note (the Suggested panel already shows the full diff)
    if simplified != formula_str:
        suggestions.append(u"Simplification applied \u2014 see the Suggested panel above.")

    # Division by zero
    if _find_div_by_zero(ast):
        suggestions.append(
            u"\u26a0  Division or modulo by zero detected. This will cause a "
            u"Revit formula error at runtime."
        )

    # Repeated subexpressions — build structured suggestion dicts
    reps = _find_repeated_subexprs(ast)
    for rep_key, rep_node, rep_count in reps:
        rep_str = _denormalize(node_to_str(rep_node), name_map)
        whole_matches = {}   # {param_name: count} — entire formula equals subexpr
        partial_matches = {}  # {param_name: count} — formula contains subexpr
        for pname, pformula in param_formulas.items():
            if not pformula:
                continue
            try:
                pwork, _ = _normalize_param_names(pformula.strip(), param_names)
                past = _Parser(_tokenize(pwork)).parse()
                pcounts = defaultdict(int)
                _collect_subexprs(past, pcounts, {})
                occ = pcounts.get(rep_key, 0)
                if occ > 0:
                    if _node_key(past) == rep_key:
                        whole_matches[pname] = 1
                    else:
                        partial_matches[pname] = occ
            except Exception:
                pass
        suggestions.append({
            'type': 'repeated_subexpr',
            'subexpr': rep_str,
            'rep_key': rep_key,
            'count_self': rep_count,
            'whole_matches': whole_matches,
            'partial_matches': partial_matches,
        })

    # Complexity
    metrics = _complexity_metrics(ast)
    if metrics['nodes'] > 15 or metrics['depth'] > 6:
        suggestions.append(
            u"Complex expression (nodes={nodes}, depth={depth}). Consider "
            u"splitting into intermediate parameters.".format(**metrics)
        )

    # No issues found
    if not suggestions:
        suggestions.append(u"No simplifications found. Formula looks good.")

    # Diff between original and simplified
    diff = _simple_diff(formula_str, simplified)

    return {
        'ok': True,
        'simplified': simplified,
        'suggestions': suggestions,
        'metrics': metrics,
        'diff': diff,
    }
