function isEmptyValue(value) {
  return value == null || value === '';
}

function normalizeValue(input) {
  if (input && typeof input === 'object' && 'v' in input) return input.v;
  return input;
}

function normalizeRules(columnSchema = {}, fmt = {}) {
  const fromSchema = Array.isArray(columnSchema.rules) ? columnSchema.rules : [];
  const fromFormat = Array.isArray(fmt.rules) ? fmt.rules : [];
  return [...fromSchema, ...fromFormat];
}

function applyBasicStyle(fmt = {}, targetStyle = {}) {
  if (fmt.color) targetStyle.color = fmt.color;
  if (fmt.bg) targetStyle.backgroundColor = fmt.bg;
  if (fmt.bold) targetStyle.fontWeight = '700';
  if (fmt.wrap === false) targetStyle.whiteSpace = 'nowrap';
  if (fmt.wrap === true) targetStyle.whiteSpace = 'normal';
  return targetStyle;
}

function applyConditionalRules(value, rules, style = {}) {
  for (const rule of rules) {
    if (!rule || typeof rule !== 'object') continue;
    const when = rule.when ?? {};
    let matched = false;

    if ('equals' in when) matched = value === when.equals;
    if (when.empty === true) matched = isEmptyValue(value);

    if (!matched) continue;
    const nextStyle = rule.style ?? {};
    if (nextStyle.color) style.color = nextStyle.color;
    if (nextStyle.bg) style.backgroundColor = nextStyle.bg;
    if (nextStyle.bold) style.fontWeight = '700';
  }
  return style;
}

function pickAlign(columnSchema = {}, fmt = {}) {
  if (fmt.align) return fmt.align;
  if (columnSchema.type === 'number' || columnSchema.type === 'money') return 'right';
  if (columnSchema.type === 'bool') return 'center';
  return 'left';
}

function formatByType(value, columnSchema = {}, ctx = {}) {
  const locale = ctx.locale ?? 'uk-UA';
  const type = columnSchema.type ?? 'text';

  if (isEmptyValue(value)) return '';

  if (type === 'number' || type === 'money') {
    const num = Number(value);
    if (Number.isNaN(num)) return String(value);
    const fractionDigits = Number.isInteger(num) ? 0 : 2;
    return new Intl.NumberFormat(locale, {
      minimumFractionDigits: 0,
      maximumFractionDigits: fractionDigits
    }).format(num);
  }

  if (type === 'date') {
    const date = value instanceof Date ? value : new Date(value);
    if (Number.isNaN(date.getTime())) return String(value);
    const dateFormat = ctx.dateFormat ?? 'DD.MM.YYYY';
    if (dateFormat === 'DD.MM.YYYY') {
      const dd = String(date.getDate()).padStart(2, '0');
      const mm = String(date.getMonth() + 1).padStart(2, '0');
      const yyyy = String(date.getFullYear());
      return `${dd}.${mm}.${yyyy}`;
    }
    return date.toLocaleDateString(locale);
  }

  if (type === 'bool') {
    return value ? 'Так' : 'Ні';
  }

  if (type === 'enum') {
    const options = Array.isArray(columnSchema.options) ? columnSchema.options : [];
    const option = options.find((x) => x?.value === value);
    return option?.label ?? String(value);
  }

  return String(value);
}

function editorByType(columnSchema = {}) {
  const type = columnSchema.type ?? 'text';
  if (type === 'bool') return { type: 'checkbox', props: {} };
  if (type === 'date') return { type: 'date', props: {} };
  if (type === 'enum') {
    return {
      type: 'select',
      props: {
        options: Array.isArray(columnSchema.options) ? columnSchema.options : []
      }
    };
  }
  if (type === 'number' || type === 'money') return { type: 'number', props: {} };
  return { type: 'text', props: {} };
}

export function formatCell(value, fmt = {}, columnSchema = {}, ctx = {}) {
  const rawValue = normalizeValue(value);
  const text = formatByType(rawValue, columnSchema, ctx);
  const style = applyConditionalRules(rawValue, normalizeRules(columnSchema, fmt), applyBasicStyle(fmt, {}));
  const align = pickAlign(columnSchema, fmt);

  return {
    text,
    style,
    className: fmt.className,
    align,
    editor: editorByType(columnSchema)
  };
}

export function parseInput(inputText, columnSchema = {}) {
  const type = columnSchema.type ?? 'text';

  if (inputText == null) return { t: type, v: null };
  const text = String(inputText).trim();
  if (text === '') return { t: type, v: null };

  if (type === 'number' || type === 'money') {
    const normalized = text.replace(/\s/g, '').replace(',', '.');
    const parsed = Number(normalized);
    return { t: type, v: Number.isNaN(parsed) ? null : parsed };
  }

  if (type === 'bool') {
    const lower = text.toLowerCase();
    const truthy = ['1', 'true', 'так', 'yes', 'y'];
    const falsy = ['0', 'false', 'ні', 'no', 'n'];
    if (truthy.includes(lower)) return { t: type, v: true };
    if (falsy.includes(lower)) return { t: type, v: false };
    return { t: type, v: null };
  }

  if (type === 'date') {
    const isoLike = /^\d{4}-\d{2}-\d{2}$/;
    const uaLike = /^(\d{2})\.(\d{2})\.(\d{4})$/;
    let date = null;
    if (isoLike.test(text)) date = new Date(`${text}T00:00:00.000Z`);
    else {
      const m = text.match(uaLike);
      if (m) date = new Date(`${m[3]}-${m[2]}-${m[1]}T00:00:00.000Z`);
    }
    if (!date || Number.isNaN(date.getTime())) return { t: type, v: null };
    return { t: type, v: date.toISOString().slice(0, 10) };
  }

  if (type === 'enum') {
    const options = Array.isArray(columnSchema.options) ? columnSchema.options : [];
    const found = options.find((x) => String(x?.value) === text || String(x?.label) === text);
    return { t: type, v: found?.value ?? text };
  }

  return { t: type, v: text };
}

export function createTableFormatterModule({ dateFormat = 'DD.MM.YYYY', locale = 'uk-UA' } = {}) {
  const moduleSettings = { dateFormat, locale };

  return {
    id: '@sdo/module-table-formatter',
    version: '1.0.0',
    init(ctx) {
      ctx.registerSchema({
        id: '@sdo/module-table-formatter.schema',
        version: '1.0.0',
        domain: 'table',
        appliesTo: { any: true },
        fields: [
          { key: 'type', label: 'Type', type: 'text' },
          { key: 'format', label: 'Format', type: 'text' },
          { key: 'options', label: 'Options', type: 'json' }
        ]
      });

      ctx.registerSettings({
        id: '@sdo/module-table-formatter.settings',
        tab: { id: 'table-formatter', title: 'Table Formatter', order: 21 },
        fields: [
          {
            key: '@sdo/module-table-formatter:dateFormat',
            label: 'Date format',
            type: 'text',
            default: 'DD.MM.YYYY',
            read: async () => moduleSettings.dateFormat,
            write: async (_runtime, value) => { moduleSettings.dateFormat = value || 'DD.MM.YYYY'; }
          },
          {
            key: '@sdo/module-table-formatter:locale',
            label: 'Locale',
            type: 'text',
            default: 'uk-UA',
            read: async () => moduleSettings.locale,
            write: async (_runtime, value) => { moduleSettings.locale = value || 'uk-UA'; }
          }
        ]
      });

      ctx.registerCommands([
        {
          id: '@sdo/module-table-formatter.preview',
          title: 'Preview formatted cell',
          run: async (_runtime, args = {}) => formatCell(args.value, args.fmt, args.columnSchema, moduleSettings)
        }
      ]);

      ctx.ui.registerButton({
        id: '@sdo/module-table-formatter:preview',
        label: 'Formatter Preview',
        location: 'toolbar',
        order: 41,
        onClick: () => ctx.commands.run('@sdo/module-table-formatter.preview', {
          value: { t: 'text', v: 'demo' },
          fmt: {},
          columnSchema: { type: 'text' }
        })
      });
    }
  };
}
