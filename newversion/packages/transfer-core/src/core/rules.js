const SUPPORTED_MATH = new Set(['+', '-', '*', '/']);

function toText(value, trim) {
  const raw = value == null ? '' : String(value);
  return trim ? raw.trim() : raw;
}

function toNumber(value, mode = 'strict') {
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  if (mode === 'loose') {
    const num = Number(value);
    return Number.isFinite(num) ? num : Number.NaN;
  }
  return Number.NaN;
}

export function evaluateRule(rule, resolvedSources) {
  const values = resolvedSources.map((entry) => entry.value);

  if (rule.op === 'direct') {
    return { ok: true, value: values[0], details: { op: 'direct', inputs: values } };
  }

  if (rule.op === 'concat') {
    const separator = rule.params?.separator === '\\n' ? '\n' : (rule.params?.separator ?? '');
    const trim = Boolean(rule.params?.trim);
    const skipEmpty = Boolean(rule.params?.skipEmpty);
    const prepared = values.map((value) => toText(value, trim));
    const filtered = skipEmpty ? prepared.filter((value) => value !== '') : prepared;
    return {
      ok: true,
      value: filtered.join(separator),
      details: { op: 'concat', separator, inputs: values, prepared: filtered }
    };
  }

  if (rule.op === 'math') {
    const mathOp = rule.params?.mathOp ?? '+';
    if (!SUPPORTED_MATH.has(mathOp)) {
      return { ok: false, error: 'unsupported_math_op', value: null, details: { op: 'math', inputs: values } };
    }

    const mode = rule.params?.coerceNumeric ?? 'strict';
    const numbers = values.map((value) => toNumber(value, mode));
    if (numbers.some((value) => Number.isNaN(value))) {
      return { ok: false, error: 'math_non_numeric', value: null, details: { op: 'math', inputs: values, numbers } };
    }

    let result = numbers[0] ?? 0;
    for (let i = 1; i < numbers.length; i += 1) {
      const next = numbers[i];
      if (mathOp === '+') result += next;
      if (mathOp === '-') result -= next;
      if (mathOp === '*') result *= next;
      if (mathOp === '/') result /= next;
    }

    if (typeof rule.params?.precision === 'number') {
      result = Number(result.toFixed(rule.params.precision));
    }

    return { ok: true, value: result, details: { op: 'math', mathOp, inputs: values, numbers } };
  }

  return { ok: false, error: 'unsupported_op', value: null, details: { op: rule.op, inputs: values } };
}
