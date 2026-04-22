// js/parser.js — Client-side Excel parser for ALMA extraction files
// Parses "Informe TSCAT v3" sheet and produces per-release, per-category metrics
// Column layout (0-based): B=1 task name, E=4 category, F=5 Estimación, P=15 TOTAL AC

(function (global) {
  'use strict';

  const TASK_COL = 1;   // col B — row identifier / task name
  const CAT_COL  = 4;   // col E — category (non-empty only on breakdown rows)
  const DEFAULT_EST_COL   = 5;   // col F — Estimación
  const DEFAULT_SPENT_COL = 15;  // col P — TOTAL AC (accumulated spent)
  const SHEET_NAME = 'Informe TSCAT v3';

  const CATEGORY_KEYS = ['ANALYSIS', 'DEVELOPMENT', 'TESTING', 'DEPLOY', 'PRODUCTION', 'MANAGEMENT'];

  // ------------------------------------------------------------------
  // Normalise raw category string (col E) → internal key
  // ------------------------------------------------------------------
  function normalizeCategory(raw) {
    if (!raw) return null;
    const c = raw.toString().toLowerCase().trim()
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '');  // strip accents
    if (/^anal/.test(c))                          return 'ANALYSIS';
    if (c === 'dt')                               return 'DEVELOPMENT';
    if (/^const|^construc/.test(c))               return 'DEVELOPMENT';
    if (/^prueba/.test(c) || /^testing/.test(c))  return 'TESTING';
    if (c === 'deploy')                           return 'DEPLOY';
    if (/^puesta|^prod/.test(c))                  return 'PRODUCTION';
    if (/^gesti|^tareas/.test(c))                 return 'MANAGEMENT';
    if (/^planif|^resoluc/.test(c))               return 'MANAGEMENT';  // Planificación / Resolución Incidencias → Management
    return null; // ignore unknown categories
  }

  // ------------------------------------------------------------------
  // Classify the parent US type from the breakdown task name (col B)
  // ------------------------------------------------------------------
  function classifyUS(taskName) {
    const t = taskName.toString();
    if (/deuda/i.test(t))                                          return 'TECH_DEBT';
    if (/uat\s+interna|interna.*uat|validaci[oó]n\s+interna/i.test(t)) return 'INTERNAL_UAT';
    if (/\buat\b/i.test(t))                                        return 'CLIENT_UAT';
    return 'REGULAR';
  }

  // Extract release key e.g. "2026R1" from a task name containing "TSCAT_2026R1..."
  function extractReleaseKey(taskName) {
    const m = taskName.toString().match(/TSCAT_(\d{4}R\d+)/i);
    return m ? m[1] : null;
  }

  // Build a zeroed category bucket
  function zeroBucket() { return { estimated: 0, spent: 0 }; }

  // Create a fresh release object
  function newRelease(key) {
    const rel = {
      id:          'TSCAT_' + key,
      releaseKey:  key,
      label:       key.replace(/^\d{4}(R\d+)$/, '$1'),   // "2026R1" → "R1"
      regular:     {},
      techDebt:    zeroBucket(),
      clientUAT:   zeroBucket(),
      internalUAT: zeroBucket(),
      totals:      null
    };
    CATEGORY_KEYS.forEach(k => { rel.regular[k] = zeroBucket(); });
    return rel;
  }

  // Accumulate est/spent into a bucket
  function acc(bucket, est, spent) {
    bucket.estimated += est;
    bucket.spent     += spent;
  }

  // ------------------------------------------------------------------
  // Auto-detect column indices from header rows
  // ------------------------------------------------------------------
  function detectColumns(data) {
    let estimCol = DEFAULT_EST_COL;
    let spentCol = DEFAULT_SPENT_COL;

    // Row 2 (0-indexed: 1) typically has "Estimación" in col F (idx 5)
    const h1 = data[1] || [];
    for (let c = 0; c < h1.length; c++) {
      const v = (h1[c] || '').toString().toLowerCase()
        .normalize('NFD').replace(/[\u0300-\u036f]/g, '');
      if (v.includes('estimac') || v.includes('estimado') || v.includes('total estimado')) {
        estimCol = c;
        break;
      }
    }

    // Row 3 (0-indexed: 2) typically has "TOTAL AC" or "TOTAL" near the end
    const h2 = data[2] || [];
    for (let c = h2.length - 1; c >= 0; c--) {
      const v = (h2[c] || '').toString().toLowerCase();
      if (v.includes('total')) { spentCol = c; break; }
    }

    return { estimCol, spentCol };
  }

  // ------------------------------------------------------------------
  // Round to 1 decimal place to avoid floating-point noise
  // ------------------------------------------------------------------
  function r1(n) { return Math.round(n * 10) / 10; }

  // ------------------------------------------------------------------
  // Main parse function
  // ------------------------------------------------------------------
  function parseALMAExcel(arrayBuffer) {
    if (!global.XLSX) {
      throw new Error('SheetJS (XLSX) library not loaded. Ensure the CDN script is included before parser.js');
    }

    const wb = XLSX.read(arrayBuffer, { type: 'array', raw: true });

    if (!wb.SheetNames.includes(SHEET_NAME)) {
      throw new Error(
        'Sheet "' + SHEET_NAME + '" not found.\n' +
        'Available sheets: ' + wb.SheetNames.join(', ')
      );
    }

    const ws   = wb.Sheets[SHEET_NAME];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });

    const { estimCol, spentCol } = detectColumns(data);

    const releases    = {};
    const releaseOrder = [];

    // Process from row index 4 (Excel row 5) — past the two header rows + blank rows
    for (let i = 4; i < data.length; i++) {
      const row = data[i];
      if (!row) continue;

      const taskName = (row[TASK_COL] || '').toString().trim();
      if (!taskName) continue;

      // Only breakdown rows have a non-empty category (col E)
      const rawCat = (row[CAT_COL] || '').toString().trim();
      if (!rawCat) continue;

      const releaseKey = extractReleaseKey(taskName);
      if (!releaseKey) continue;

      if (!releases[releaseKey]) {
        releases[releaseKey] = newRelease(releaseKey);
        releaseOrder.push(releaseKey);
      }

      const estimated = Math.max(0, parseFloat(row[estimCol]) || 0);
      const spent     = Math.max(0, parseFloat(row[spentCol])  || 0);
      const usType    = classifyUS(taskName);
      const catKey    = normalizeCategory(rawCat);
      const rel       = releases[releaseKey];

      if (usType === 'TECH_DEBT') {
        acc(rel.techDebt, estimated, spent);
      } else if (usType === 'CLIENT_UAT') {
        acc(rel.clientUAT, estimated, spent);
      } else if (usType === 'INTERNAL_UAT') {
        acc(rel.internalUAT, estimated, spent);
      } else if (catKey) {
        acc(rel.regular[catKey], estimated, spent);
      }
    }

    // Compute totals and round
    releaseOrder.forEach(key => {
      const rel = releases[key];

      CATEGORY_KEYS.forEach(k => {
        rel.regular[k].estimated = r1(rel.regular[k].estimated);
        rel.regular[k].spent     = r1(rel.regular[k].spent);
      });
      ['techDebt','clientUAT','internalUAT'].forEach(b => {
        rel[b].estimated = r1(rel[b].estimated);
        rel[b].spent     = r1(rel[b].spent);
      });

      const regEst   = r1(CATEGORY_KEYS.reduce((s, k) => s + rel.regular[k].estimated, 0));
      const regSpent = r1(CATEGORY_KEYS.reduce((s, k) => s + rel.regular[k].spent,     0));
      const allEst   = r1(regEst   + rel.techDebt.estimated + rel.clientUAT.estimated + rel.internalUAT.estimated);
      const allSpent = r1(regSpent + rel.techDebt.spent     + rel.clientUAT.spent     + rel.internalUAT.spent);

      rel.totals = {
        regular: { estimated: regEst, spent: regSpent },
        all:     { estimated: allEst, spent: allSpent },
        uatSpent: r1(rel.clientUAT.spent + rel.internalUAT.spent)
      };
    });

    return { releases, releaseOrder };
  }

  // ------------------------------------------------------------------
  // Expose
  // ------------------------------------------------------------------
  global.ALMAParser = {
    parseALMAExcel: parseALMAExcel,
    CATEGORY_KEYS:  CATEGORY_KEYS
  };

})(window);
