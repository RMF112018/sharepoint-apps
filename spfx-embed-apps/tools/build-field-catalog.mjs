#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';
import process from 'node:process';
import xlsx from 'xlsx';

function toCamelCase(input) {
  const cleaned = String(input)
    .replace(/\*+/g, ' ')
    .replace(/\([^)]*\)/g, ' ') // remove parentheticals
    .replace(/\[[^\]]*\]/g, ' ') // remove brackets content
    .replace(/default\s*[:=].*$/i, ' ') // strip explicit default tail
    .replace(/[^A-Za-z0-9]+/g, ' ') // non-alnum to space
    .trim()
    .toLowerCase();
  if (!cleaned) return 'field';
  const parts = cleaned.split(/\s+/);
  const camel = parts
    .map((p, i) => (i === 0 ? p : p.charAt(0).toUpperCase() + p.slice(1)))
    .join('');
  return /^[A-Za-z_]/.test(camel) ? camel : `field${camel.charAt(0).toUpperCase()}${camel.slice(1)}`;
}

function inferRequired(labelRaw) {
  const s = String(labelRaw);
  return /(\brequired\b|\*)/i.test(s);
}

function extractDefault(labelRaw) {
  const s = String(labelRaw);
  const m = s.match(/default\s*[:=]\s*([^)\]\n]+)/i);
  return m ? m[1].trim() : '';
}

function extractOptions(labelRaw) {
  const s = String(labelRaw);
  // look for parentheses or brackets with delimited values
  const paren = s.match(/\(([^()]*)\)/);
  const bracket = s.match(/\[([^\[\]]*)\]/);
  const candidate = paren?.[1] || bracket?.[1] || '';
  // ignore segments that contain 'default'
  const cleaned = candidate.replace(/default\s*[:=][^,|/]+/ig, '').trim();
  if (!cleaned) return [];
  const pieces = cleaned.split(/\s*[\|,\/]\s*/).map(v => v.trim()).filter(Boolean);
  // Ensure options look like words (not dates or numbers only); otherwise, return []
  const looksLikeOptions = pieces.length >= 2 && pieces.every(p => /[A-Za-z]/.test(p));
  return looksLikeOptions ? Array.from(new Set(pieces)) : [];
}

function inferType(labelRaw, options) {
  const s = String(labelRaw).toLowerCase();
  if (/\?|^\s*(is|has)\b/.test(s) || /\byes\s*\/\s*no\b/.test(s)) return 'Boolean';
  if (/date|deadline|due|start date|end date|time/.test(s)) return 'DateTime';
  if (/amount|budget|cost|currency|price|total|count|number|qty|quantity|hours|days|size|score|percent/.test(s)) return 'Number';
  if (/description|details|notes|justification|summary|comment|objective|scope/.test(s)) return 'Multiline';
  if (/(responsible|owner|manager|lead|requestor|requester|assigned\s*to|approver|contact|email|sponsor)/.test(s)) return 'Person';
  if (options && options.length > 0) return 'Choice';
  if (/status|priority|category|type|phase|stage/.test(s)) return 'Choice';
  return 'Text';
}

function inferGroup(labelRaw) {
  const s = String(labelRaw).toLowerCase();
  if (/deliverable/.test(s)) {
    if (/non[-\s]?standard|nonstandard/.test(s)) return 'Final Deliverables – Non-standard';
    if (/standard/.test(s)) return 'Final Deliverables – Standard';
    return 'Final Deliverables – Standard';
  }
  if (/date|deadline|due|start|end|milestone/.test(s)) return 'Key Dates';
  if (/(owner|manager|lead|requestor|requester|assigned\s*to|approver|contact|department|division|sponsor|status|priority)/.test(s)) return 'Managing Information';
  return 'Project Info';
}

function cleanLabel(labelRaw) {
  return String(labelRaw)
    .replace(/\s*\*+\s*$/, '')
    .replace(/\s*\(\s*required\s*\)\s*/i, '')
    .replace(/\s*\[\s*required\s*\]\s*/i, '')
    .trim();
}

function buildCatalogFromSheet(sheet) {
  const rows = xlsx.utils.sheet_to_json(sheet, { header: 1, raw: true });
  if (!rows.length) return [];
  const catalog = [];
  const seen = new Set();
  // Prefer vertical layout: use first column as labels, iterate until a blank
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    if (!row || row.length === 0) continue;
    const labelCell = row[0];
    if (labelCell === undefined || labelCell === null) continue;
    const labelStr = String(labelCell).trim();
    if (labelStr === '') continue;
    const label = cleanLabel(labelStr);
    const required = inferRequired(labelStr);
    const choice_options = extractOptions(labelStr);
    const suggested_type = inferType(labelStr, choice_options);
    const def = extractDefault(labelStr);
    const field_key = toCamelCase(label);
    const group = inferGroup(labelStr);
    const key = `${field_key}::${label.toLowerCase()}`;
    if (seen.has(key)) continue;
    seen.add(key);
    catalog.push({
      field_key,
      label,
      group,
      suggested_type,
      required,
      help_text: '',
      choice_options,
      default: def,
      multi_select: false
    });
  }
  // If only 1 row and multiple columns, fallback to header-across interpretation
  if (catalog.length <= 1) {
    const header = rows[0];
    for (let col = 0; col < header.length; col++) {
      const labelCell = header[col];
      if (labelCell === undefined || labelCell === null || String(labelCell).trim() === '') continue;
      const label = cleanLabel(labelCell);
      const required = inferRequired(labelCell);
      const choice_options = extractOptions(labelCell);
      const suggested_type = inferType(labelCell, choice_options);
      const def = extractDefault(labelCell);
      const field_key = toCamelCase(label);
      const group = inferGroup(labelCell);
      const key = `${field_key}::${String(label).toLowerCase()}`;
      if (seen.has(key)) continue;
      seen.add(key);
      catalog.push({
        field_key,
        label,
        group,
        suggested_type,
        required,
        help_text: '',
        choice_options,
        default: def,
        multi_select: false
      });
    }
  }
  return catalog;
}

function postProcessCatalog(catalog) {
  function matches(label, pattern) {
    return pattern.test(label.toLowerCase());
  }
  for (const f of catalog) {
    const label = f.label;
    // Proposal Delivered Via → Choice
    if (matches(label, /proposal\s+delivered\s+via/)) {
      f.suggested_type = 'Choice';
      f.choice_options = ['Hand Delivery', 'Email', 'Portal', 'Other'];
      if (!f.default) f.default = '';
    }
    // Type of Proposal → Choice
    if (matches(label, /type\s+of\s+proposal/)) {
      f.suggested_type = 'Choice';
      f.choice_options = ['Lump Sum', 'GMP', 'Cost Plus', 'Budget', 'Alternate', 'Other'];
      if (!f.default) f.default = '';
    }
    // RFI Format → Choice with default Procore (if typical)
    if (matches(label, /rfi\s*format/)) {
      f.suggested_type = 'Choice';
      f.choice_options = ['Excel', 'Procore'];
      if (!f.default || !f.choice_options.includes(f.default)) {
        f.default = 'Procore';
      }
    }
    // Schedule Type(s) → Choice
    if (matches(label, /schedule\s*type/)) {
      f.suggested_type = 'Choice';
      f.choice_options = ['Milestone', 'Precon', 'CPM'];
      if (!f.default) f.default = '';
    }
    // Proposal Book “Tab Required?” and other Yes/No → Boolean
    if (matches(label, /(tab\s*required|tab\s*req\'?d|yes\/?no)/)) {
      f.suggested_type = 'Boolean';
      f.choice_options = [];
      if (typeof f.default !== 'boolean') f.default = false;
    }
    // Managing-Info Yes/No become single Boolean toggles
    if (f.group === 'Managing Information' && matches(label, /yes\/?no/)) {
      f.suggested_type = 'Boolean';
      f.choice_options = [];
      if (typeof f.default !== 'boolean') f.default = false;
    }
    // Multi-select inference for plural role labels
    if (f.suggested_type === 'Person') {
      if (/reviewers|owners|approvers|team|stakeholders|contacts|leads/i.test(label)) {
        f.multi_select = true;
      }
    }
  }
  return catalog;
}

function writeJson(filePath, data) {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  fs.writeFileSync(filePath, JSON.stringify(data, null, 2), 'utf8');
}

function writeMarkdown(filePath, catalog) {
  const lines = [];
  lines.push('# Field Catalog');
  lines.push('');
  lines.push('| field_key | label | group | suggested_type | required | help_text | choice_options | default |');
  lines.push('|---|---|---|---|---:|---|---|---|');
  for (const f of catalog) {
    const opts = (f.choice_options && f.choice_options.length) ? f.choice_options.join(', ') : '';
    lines.push(`| ${f.field_key} | ${f.label} | ${f.group} | ${f.suggested_type} | ${f.required ? 'true' : 'false'} | ${f.help_text} | ${opts} | ${f.default || ''} |`);
  }
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  fs.writeFileSync(filePath, lines.join('\n'), 'utf8');
}

function main() {
  const repoRoot = path.resolve(path.dirname(new URL(import.meta.url).pathname), '..');
  // Arguments
  const excelArgIdx = process.argv.findIndex(a => a === '--xlsx');
  const appArgIdx = process.argv.findIndex(a => a === '--app');
  const excelPath = excelArgIdx !== -1 ? process.argv[excelArgIdx + 1] : path.join(repoRoot, 'mnt', 'data', 'NewProjectForm.xlsx');
  const appName = appArgIdx !== -1 ? process.argv[appArgIdx + 1] : 'NewProjectIntake';
  const wb = xlsx.readFile(excelPath);
  const sheet = wb.Sheets['Sheet1'] || wb.Sheets[wb.SheetNames[0]];
  if (!sheet) {
    console.error('No Sheet1 found in workbook');
    process.exit(1);
  }
  const catalog = buildCatalogFromSheet(sheet);
  postProcessCatalog(catalog);
  const outJson = path.join(repoRoot, 'apps', appName, 'src', 'config', 'fieldCatalog.json');
  const outMd = path.join(repoRoot, 'apps', appName, 'docs', 'FieldCatalog.md');
  writeJson(outJson, catalog);
  writeMarkdown(outMd, catalog);
  // Print to stdout as well
  console.log(JSON.stringify(catalog, null, 2));
}

main();


