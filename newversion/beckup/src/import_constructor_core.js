function norm(v) {
  return String(v || '').trim().toLowerCase();
}

function toTargetMeta(targetColumns) {
  return (targetColumns || []).map((c, i) => {
    if (typeof c === 'string') return { key: c, name: c, index: i };
    return {
      key: c?.key ?? c?.name ?? `col_${i + 1}`,
      name: c?.name ?? c?.key ?? `col_${i + 1}`,
      index: i
    };
  });
}

/**
 * Suggest mapping source columns to target columns using header names + aliases.
 * Returns mapping format: [{ sourceCol: 1-based, targetKey }]
 */
export function suggestColumnMapping({ headerRow = [], targetColumns = [], aliases = {} } = {}) {
  const targetMeta = toTargetMeta(targetColumns);
  const byHeader = new Map();

  for (let i = 0; i < headerRow.length; i += 1) {
    const h = norm(headerRow[i]);
    if (!h) continue;
    byHeader.set(h, i + 1);
  }

  const mapping = [];
  for (const t of targetMeta) {
    const candidates = [t.name, t.key, ...(aliases[t.key] || []), ...(aliases[t.name] || [])]
      .map(norm)
      .filter(Boolean);

    let sourceCol = null;
    for (const candidate of candidates) {
      if (byHeader.has(candidate)) {
        sourceCol = byHeader.get(candidate);
        break;
      }
    }

    if (Number.isFinite(sourceCol)) mapping.push({ sourceCol, targetKey: t.key });
  }

  return mapping;
}

/**
 * Validate and normalize mapping to canonical plan entries.
 */
export function buildImportPlan({ mapping = [], targetColumns = [] } = {}) {
  const targetMeta = toTargetMeta(targetColumns);
  const targetKeys = new Set(targetMeta.map((t) => t.key));

  const plan = [];
  const warnings = [];

  for (const m of mapping || []) {
    const sourceCol = Number(m?.sourceCol);
    const targetKey = String(m?.targetKey || '');

    if (!Number.isFinite(sourceCol) || sourceCol < 1) {
      warnings.push('Skipped mapping with invalid sourceCol');
      continue;
    }
    if (!targetKey || !targetKeys.has(targetKey)) {
      warnings.push(`Skipped mapping with unknown targetKey: ${targetKey || '(empty)'}`);
      continue;
    }

    plan.push({ sourceCol, targetKey });
  }

  return { plan, warnings };
}

/**
 * Apply normalized import plan to 2D matrix rows.
 */
export function applyImportPlanToRows({ rows = [], plan = [], dataRowStartIndex = 1 } = {}) {
  const records = [];

  for (let r = dataRowStartIndex; r < rows.length; r += 1) {
    const src = rows[r] || [];
    const cells = {};
    let hasAny = false;

    for (const p of plan) {
      const v = src[p.sourceCol - 1] ?? '';
      const text = String(v ?? '');
      if (text.trim() !== '') hasAny = true;
      cells[p.targetKey] = text;
    }

    if (!hasAny) continue;
    records.push({ id: crypto.randomUUID(), cells, subrows: [] });
  }

  return records;
}
