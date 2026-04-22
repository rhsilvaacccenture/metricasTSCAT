// js/renderer.js — DOM updater for parsed ALMA Excel data
// Updates slides S1 (Summary), S2 (R1), S3 (R2), S4 (R3), S5 (Consolidated), S6 (Velocity)
// Stores window.reportData for PPTX generation

(function (global) {
  'use strict';

  // ------------------------------------------------------------------
  // Number formatting helpers
  // ------------------------------------------------------------------
  function fmt(n, dp) {
    dp = (dp === undefined) ? 1 : dp;
    return n.toLocaleString('en-US', { minimumFractionDigits: dp, maximumFractionDigits: dp });
  }

  function fmtH(n) { return fmt(n) + 'h'; }

  function fmtPct(est, spent) {
    if (!est || est === 0) return { text: '—', positive: null };
    const saving = (est - spent) / est;
    const pct    = Math.round(saving * 100);
    return { text: (pct >= 0 ? '+' : '') + pct + '%', positive: pct >= 0 };
  }

  function devText(est, spent) {
    const d = Math.round((est - spent) * 10) / 10;
    return { text: (d >= 0 ? '+' : '') + fmt(d), positive: d >= 0 };
  }

  // ------------------------------------------------------------------
  // HTML building blocks
  // ------------------------------------------------------------------
  function metricCard(label, value, hi, negative) {
    const cls = hi ? 'metric hi' : 'metric';
    const bg  = (hi && negative) ? ' style="background:#DC2626;"' : '';
    return '<div class="' + cls + '"' + bg + '><div class="mlabel">' +
           label + '</div><div class="mval">' + value + '</div></div>';
  }

  function tableRow(cells) {
    return '<tr>' + cells.map(function(c) {
      const cls = c.cls ? ' class="' + c.cls + '"' : '';
      return '<td' + cls + '>' + (c.val !== undefined ? c.val : c) + '</td>';
    }).join('') + '</tr>';
  }

  // Build a standard category row
  function catRow(area, est, spent) {
    if (est === 0 && spent === 0) {
      return tableRow([area, '0', '0', '0', '—']);
    }
    const d   = devText(est, spent);
    const pct = fmtPct(est, spent);
    const dcls = d.positive ? 'pos' : 'neg';
    const pcls = pct.positive === null ? '' : pct.positive ? 'pos' : 'neg';
    return tableRow([
      area,
      { val: fmt(est), cls: '' },
      { val: fmt(spent), cls: '' },
      { val: d.text, cls: dcls },
      { val: pct.text, cls: pcls }
    ]);
  }

  const TABLE_HEAD_5 =
    '<thead><tr><th>Area</th><th>Estimated (h)</th><th>Spent (h)</th>' +
    '<th>Deviation (h)</th><th>% Saving</th></tr></thead>';

  // ------------------------------------------------------------------
  // Build view model for one release (R1/R2/R3)
  // ------------------------------------------------------------------
  function buildReleaseViewModel(rel) {
    const reg   = rel.regular;
    const tot   = rel.totals;
    const isR1  = rel.label === 'R1';

    // Metric cards
    const metrics = [];

    if (isR1) {
      const spentExclUAT = Math.round((tot.regular.spent) * 10) / 10;
      const savingExcl   = fmtPct(tot.regular.estimated, spentExclUAT);
      metrics.push({ label: 'Total Estimated',      value: fmtH(tot.regular.estimated) });
      metrics.push({ label: 'Spent excl. UAT R1',   value: fmtH(spentExclUAT) });
      metrics.push({ label: 'Saving excl. UAT',     value: savingExcl.text, hi: true, negative: !savingExcl.positive });
      if (tot.uatSpent > 0) {
        metrics.push({ label: 'UAT Technical Debt', value: fmtH(tot.uatSpent) });
      }
    } else {
      const saved   = Math.round((tot.regular.estimated - tot.regular.spent) * 10) / 10;
      const saving  = fmtPct(tot.regular.estimated, tot.regular.spent);
      const dtSaving = fmtPct(reg.DEVELOPMENT.estimated, reg.DEVELOPMENT.spent);
      metrics.push({ label: 'Total Estimated', value: fmtH(tot.regular.estimated) });
      metrics.push({ label: 'Total Spent',     value: fmtH(tot.regular.spent) });
      if (saved !== 0) {
        metrics.push({ label: 'Hours Saved',   value: fmtH(Math.abs(saved)) });
      }
      metrics.push({ label: 'Overall Saving',  value: saving.text,   hi: true, negative: !saving.positive });
      if (reg.DEVELOPMENT.estimated > 0) {
        metrics.push({ label: 'DT/Const Saving', value: dtSaving.text, hi: true, negative: !dtSaving.positive });
      }
    }

    // Category table rows  (regular USs only; no tech debt / UAT)
    const DISPLAY = [
      { key: 'ANALYSIS',    label: 'Analysis &amp; DF' },
      { key: 'DEVELOPMENT', label: 'DT / Const' },
      { key: 'TESTING',     label: 'Testing' },
      { key: 'DEPLOY',      label: 'Deploy' },
      { key: 'PRODUCTION',  label: 'Production' },
      { key: 'MANAGEMENT',  label: 'Management' }
    ];

    // For R1: add UAT back into DEVELOPMENT for the table (historic display)
    const displayReg = {};
    DISPLAY.forEach(function(d) { displayReg[d.key] = Object.assign({}, reg[d.key]); });

    if (isR1 && tot.uatSpent > 0) {
      // R1 UAT is traditionally shown merged into DT row
      displayReg['DEVELOPMENT'].spent =
        Math.round((displayReg['DEVELOPMENT'].spent + tot.uatSpent) * 10) / 10;
    }

    const tableRows = DISPLAY.map(function(d) {
      return { key: d.key, label: d.label,
               est: displayReg[d.key].estimated, spent: displayReg[d.key].spent };
    });

    // Total row — for R1 use all-in spent (incl. UAT), else regular
    const totalEst   = tot.regular.estimated;
    const totalSpent = isR1 ? tot.all.spent : tot.regular.spent;

    return { metrics: metrics, tableRows: tableRows, totalEst: totalEst, totalSpent: totalSpent };
  }

  // ------------------------------------------------------------------
  // Render one release deep-dive slide  (s2=R1, s3=R2, s4=R3)
  // ------------------------------------------------------------------
  function renderReleaseSlide(slideEl, rel) {
    const vm = buildReleaseViewModel(rel);

    // Metrics
    const metricsEl = slideEl.querySelector('.metrics');
    if (metricsEl) {
      metricsEl.innerHTML = vm.metrics.map(function(m) {
        return metricCard(m.label, m.value, m.hi, m.negative);
      }).join('');
    }

    // Table body
    const tbody = slideEl.querySelector('tbody');
    if (tbody) {
      let rows = vm.tableRows.map(function(r) { return catRow(r.label, r.est, r.spent); }).join('');

      // TOTAL row
      const tot   = devText(vm.totalEst, vm.totalSpent);
      const pct   = fmtPct(vm.totalEst, vm.totalSpent);
      let totalPctText = pct.text;
      if (rel.label === 'R1' && rel.totals.uatSpent > 0) {
        const excl = fmtPct(rel.totals.regular.estimated, rel.totals.regular.spent);
        totalPctText = pct.text + ' raw / ' + excl.text + ' excl. UAT';
      }
      rows += tableRow([
        'TOTAL',
        { val: fmt(vm.totalEst),   cls: '' },
        { val: fmt(vm.totalSpent), cls: '' },
        { val: tot.text, cls: tot.positive ? 'pos' : 'neg' },
        { val: totalPctText, cls: pct.positive ? 'pos' : 'neg' }
      ]);

      tbody.innerHTML = rows;
    }
  }

  // ------------------------------------------------------------------
  // Summary slide (s1)
  // ------------------------------------------------------------------
  function renderSummary(releases, releaseOrder) {
    const slideEl = document.getElementById('s1');
    if (!slideEl) return;

    let totalEst = 0, totalSpent = 0;
    const releaseLabels = { R1: 'R1 – Completed (raw)', R2: 'R2 – Completed', R3: 'R3 – Completed' };

    const rows = releaseOrder.map(function(key) {
      const rel   = releases[key];
      const label = (releaseLabels[rel.label] || rel.label);
      const est   = rel.totals.regular.estimated;
      const spent = rel.totals.all.spent;  // all-in for summary
      totalEst   += est;
      totalSpent += spent;

      const d   = devText(est, spent);
      const pct = fmtPct(est, spent);
      return tableRow([
        label,
        { val: fmt(est),   cls: '' },
        { val: fmt(spent), cls: '' },
        { val: d.text, cls: d.positive ? 'pos' : 'neg' },
        { val: pct.text, cls: pct.positive ? 'pos' : 'neg' }
      ]);
    });

    const tbody = slideEl.querySelector('tbody');
    if (tbody) tbody.innerHTML = rows.join('');

    // Summary metrics
    const metricsEl = slideEl.querySelector('.metrics');
    if (metricsEl) {
      const r1Rel   = releases[releaseOrder.find(function(k) { return releases[k].label === 'R1'; })];
      const uatSpent = r1Rel ? r1Rel.totals.uatSpent : 0;
      const dev      = Math.round((totalEst - totalSpent) * 10) / 10;
      const rawPct   = fmtPct(totalEst, totalSpent);
      const exclSpent = Math.round((totalSpent - uatSpent) * 10) / 10;
      const exclPct   = fmtPct(totalEst, exclSpent);

      metricsEl.innerHTML =
        metricCard('Total Estimated', fmtH(Math.round(totalEst * 10) / 10)) +
        metricCard('Total Spent', fmtH(Math.round(totalSpent * 10) / 10)) +
        metricCard('Deviation', (dev >= 0 ? '+' : '') + fmtH(Math.abs(dev))) +
        metricCard('Saving (raw)', rawPct.text, true, !rawPct.positive) +
        metricCard('Saving excl. UAT R1', exclPct.text, true, !exclPct.positive);
    }
  }

  // ------------------------------------------------------------------
  // Consolidated slide (s5)
  // ------------------------------------------------------------------
  function renderConsolidated(releases, releaseOrder) {
    const slideEl = document.getElementById('s5');
    if (!slideEl) return;

    const CATS = ALMAParser.CATEGORY_KEYS;
    const totals = {};
    CATS.forEach(function(k) { totals[k] = { estimated: 0, spent: 0 }; });
    let grandEst = 0, grandSpent = 0, r1UATSpent = 0;

    releaseOrder.forEach(function(key) {
      const rel = releases[key];
      CATS.forEach(function(k) {
        totals[k].estimated += rel.regular[k].estimated;
        totals[k].spent     += rel.regular[k].spent;
      });
      // For R1, merge UAT spent back into DEVELOPMENT (consistent with per-release view)
      if (rel.label === 'R1' && rel.totals.uatSpent > 0) {
        totals['DEVELOPMENT'].spent += rel.totals.uatSpent;
        r1UATSpent = rel.totals.uatSpent;
      }
      grandEst   += rel.totals.regular.estimated;
      grandSpent += rel.totals.all.spent;
    });

    const DISPLAY = [
      { key: 'ANALYSIS',    label: 'Analysis &amp; DF' },
      { key: 'DEVELOPMENT', label: 'DT / Const ⚠️' },
      { key: 'TESTING',     label: 'Testing' },
      { key: 'DEPLOY',      label: 'Deploy' },
      { key: 'PRODUCTION',  label: 'Production' },
      { key: 'MANAGEMENT',  label: 'Management' }
    ];

    const tbody = slideEl.querySelector('tbody');
    if (tbody) {
      let rows = DISPLAY.map(function(d) {
        const est   = Math.round(totals[d.key].estimated * 10) / 10;
        const spent = Math.round(totals[d.key].spent     * 10) / 10;
        if (est === 0 && spent === 0) return tableRow([d.label, '0', '0', '0', '—', '—']);
        const dv  = devText(est, spent);
        const pct = fmtPct(est, spent);
        const cls = pct.positive === null ? '' : pct.positive ? 'pos' : 'neg';

        // excl. UAT column — only for DEVELOPMENT
        let exclCell = '—';
        if (d.key === 'DEVELOPMENT' && r1UATSpent > 0) {
          const spentExcl = Math.round((spent - r1UATSpent) * 10) / 10;
          const exclPct   = fmtPct(est, spentExcl);
          exclCell = '<span class="' + (exclPct.positive ? 'pos' : 'neg') + '">' + exclPct.text + '</span>';
        }
        return '<tr><td>' + d.label + '</td><td>' + fmt(est) + '</td><td>' + fmt(spent) +
               '</td><td class="' + (dv.positive ? 'pos' : 'neg') + '">' + dv.text +
               '</td><td class="' + cls + '">' + pct.text +
               '</td><td>' + exclCell + '</td></tr>';
      }).join('');

      // TOTAL row
      const gd      = devText(Math.round(grandEst*10)/10, Math.round(grandSpent*10)/10);
      const rawPct  = fmtPct(grandEst, grandSpent);
      const exclSpent = Math.round((grandSpent - r1UATSpent) * 10) / 10;
      const exclPct   = fmtPct(grandEst, exclSpent);
      rows += '<tr><td>TOTAL</td><td>' + fmt(Math.round(grandEst*10)/10) +
              '</td><td>' + fmt(Math.round(grandSpent*10)/10) +
              '</td><td class="' + (gd.positive?'pos':'neg') + '">' + gd.text +
              '</td><td>' + rawPct.text + ' raw</td>' +
              '<td class="' + (exclPct.positive?'pos':'neg') + '">' + exclPct.text + ' excl. UAT R1</td></tr>';
      tbody.innerHTML = rows;
    }

    // Metrics
    const metricsEl = slideEl.querySelector('.metrics');
    if (metricsEl) {
      const testPct  = fmtPct(totals.TESTING.estimated, totals.TESTING.spent);
      const prodPct  = fmtPct(totals.PRODUCTION.estimated, totals.PRODUCTION.spent);
      const mgmtPct  = fmtPct(totals.MANAGEMENT.estimated, totals.MANAGEMENT.spent);
      const mgmtOver = !mgmtPct.positive;
      const mgmtLbl  = mgmtOver ? 'Mgmt Over Budget' : 'Mgmt Saving';
      const mgmtVal  = mgmtOver ?
        Math.abs(Math.round((totals.MANAGEMENT.spent / Math.max(1, totals.MANAGEMENT.estimated) - 1) * 100)) + '%' :
        mgmtPct.text;

      metricsEl.innerHTML =
        metricCard('Total Estimated', fmtH(Math.round(grandEst*10)/10)) +
        metricCard('Total Spent',     fmtH(Math.round(grandSpent*10)/10)) +
        metricCard('Testing Saving',  testPct.text, true, !testPct.positive) +
        metricCard('Production Saving', prodPct.text, true, !prodPct.positive) +
        metricCard(mgmtLbl, mgmtVal, true, mgmtOver);
    }
  }

  // ------------------------------------------------------------------
  // Velocity slide (s6)
  // ------------------------------------------------------------------
  function renderVelocity(releases, releaseOrder) {
    const slideEl = document.getElementById('s6');
    if (!slideEl) return;

    const tbody = slideEl.querySelector('tbody');
    if (!tbody) return;

    // Build utilization % per category per release
    function util(rel, catKey) {
      const est   = rel.regular[catKey].estimated;
      const spent = rel.regular[catKey].spent;
      if (!est || est === 0) return { text: 'Unbudgeted', cls: 'neg' };
      const pct = Math.round(spent / est * 100);
      const cls = pct <= 100 ? 'pos' : 'neg';
      return { text: pct + '%', cls: cls };
    }

    const rels = releaseOrder.map(function(k) { return releases[k]; });

    const ROWS = [
      { key: 'DEVELOPMENT', label: 'DT / Const' },
      { key: 'TESTING',     label: 'Testing' },
      { key: 'ANALYSIS',    label: 'Analysis &amp; DF' },
      { key: 'MANAGEMENT',  label: 'Management' }
    ];

    const rows = ROWS.map(function(d) {
      let cells = '<td>' + d.label + '</td>';
      rels.forEach(function(rel) {
        const u = util(rel, d.key);
        cells += '<td class="' + u.cls + '">' + u.text + '</td>';
      });
      // Insight column
      cells += '<td>—</td>';
      return '<tr>' + cells + '</tr>';
    });

    tbody.innerHTML = rows.join('');

    // Also update metric cards with last release values
    const lastRel    = rels[rels.length - 1];
    if (!lastRel) return;
    const metricsEl  = slideEl.querySelector('.metrics');
    if (metricsEl) {
      const testPct    = fmtPct(lastRel.regular.TESTING.estimated,    lastRel.regular.TESTING.spent);
      const prodPct    = fmtPct(lastRel.regular.PRODUCTION.estimated, lastRel.regular.PRODUCTION.spent);
      const firstLabel = rels[0] ? rels[0].label : 'R1';
      const lastLabel  = lastRel.label;
      metricsEl.innerHTML =
        metricCard('DT/Const Improvement', firstLabel + '\u2192' + lastLabel, true, false) +
        metricCard('Testing ' + lastLabel + ' Saving',    testPct.text, true, !testPct.positive) +
        metricCard('Production ' + lastLabel + ' Saving', prodPct.text, true, !prodPct.positive);
    }
  }

  // ------------------------------------------------------------------
  // Main entry point: render all slides from parsed data
  // ------------------------------------------------------------------
  function renderFromData(parsedData) {
    const { releases, releaseOrder } = parsedData;

    // Map release label to slide ID
    const SLIDE_MAP = { R1: 's2', R2: 's3', R3: 's4' };

    releaseOrder.forEach(function(key) {
      const rel = releases[key];
      const slideId = SLIDE_MAP[rel.label];
      if (!slideId) return;
      const el = document.getElementById(slideId);
      if (el) renderReleaseSlide(el, rel);
    });

    renderSummary(releases, releaseOrder);
    renderConsolidated(releases, releaseOrder);
    renderVelocity(releases, releaseOrder);

    // Build and store reportData for PPTX generation
    global.reportData = buildReportData(releases, releaseOrder);
  }

  // ------------------------------------------------------------------
  // Build the reportData object consumed by downloadPptx()
  // ------------------------------------------------------------------
  function buildReportData(releases, releaseOrder) {
    const rd = { releases: {}, releaseOrder: releaseOrder };

    releaseOrder.forEach(function(key) {
      const rel = releases[key];
      const vm  = buildReleaseViewModel(rel);
      const CATS = ALMAParser.CATEGORY_KEYS;

      // Category rows for PPTX table
      const DISPLAY = [
        { key: 'ANALYSIS',    label: 'Analysis & DF' },
        { key: 'DEVELOPMENT', label: 'DT / Const' },
        { key: 'TESTING',     label: 'Testing' },
        { key: 'DEPLOY',      label: 'Deploy' },
        { key: 'PRODUCTION',  label: 'Production' },
        { key: 'MANAGEMENT',  label: 'Management' }
      ];

      const tableRows = vm.tableRows.map(function(r) {
        const d   = Math.round((r.est - r.spent) * 10) / 10;
        const pct = r.est > 0 ? Math.round((r.est - r.spent) / r.est * 100) : 0;
        const pctText = r.est === 0 && r.spent === 0 ? '—' :
                        r.est === 0 ? '—' :
                        (pct >= 0 ? '+' : '') + pct + '%';
        const devText2 = r.est === 0 && r.spent === 0 ? '0' : (d >= 0 ? '+' : '') + fmt(d);
        return [r.label, fmt(r.est), fmt(r.spent), devText2, pctText];
      });

      const totalEst   = vm.totalEst;
      const totalSpent = vm.totalSpent;
      const totalDev   = Math.round((totalEst - totalSpent) * 10) / 10;
      const totalPct   = totalEst > 0 ? Math.round((totalEst - totalSpent) / totalEst * 100) : 0;
      let totalPctStr  = (totalPct >= 0 ? '+' : '') + totalPct + '%';
      if (rel.label === 'R1' && rel.totals.uatSpent > 0) {
        const exclPct = Math.round((totalEst - rel.totals.regular.spent) / totalEst * 100);
        totalPctStr   = totalPct + '% raw / ' + (exclPct >= 0 ? '+' : '') + exclPct + '% excl. UAT';
      }
      tableRows.push(['TOTAL', fmt(totalEst), fmt(totalSpent),
        (totalDev >= 0 ? '+' : '') + fmt(totalDev), totalPctStr]);

      rd.releases[key] = {
        label:     rel.label,
        metrics:   vm.metrics,
        tableRows: tableRows
      };
    });

    // ------ Summary data ------
    {
      let totalEst = 0, totalSpent = 0;
      let r1UATSpent = 0;
      const summaryRows = releaseOrder.map(function(key) {
        const rel   = releases[key];
        const est   = rel.totals.regular.estimated;
        const spent = rel.totals.all.spent;
        totalEst   += est;
        totalSpent += spent;
        if (rel.label === 'R1') r1UATSpent = rel.totals.uatSpent;
        const d   = Math.round((est - spent) * 10) / 10;
        const pct = est > 0 ? Math.round((est - spent) / est * 100) : 0;
        return [
          rel.label + ' – ' + (rel.totals.all.spent > 0 ? 'Completed' : 'In Progress'),
          fmt(est), fmt(spent),
          (d >= 0 ? '+' : '') + fmt(d),
          (pct >= 0 ? '+' : '') + pct + '%'
        ];
      });

      totalEst   = Math.round(totalEst   * 10) / 10;
      totalSpent = Math.round(totalSpent * 10) / 10;
      const dev        = Math.round((totalEst - totalSpent) * 10) / 10;
      const rawPct     = totalEst > 0 ? Math.round((totalEst - totalSpent) / totalEst * 100) : 0;
      const exclSpent  = Math.round((totalSpent - r1UATSpent) * 10) / 10;
      const exclPctVal = totalEst > 0 ? Math.round((totalEst - exclSpent) / totalEst * 100) : 0;
      const rawNeg     = rawPct < 0;
      const exclNeg    = exclPctVal < 0;

      rd.summary = {
        metrics: [
          { label: 'Total Estimated',     value: fmtH(totalEst) },
          { label: 'Total Spent',         value: fmtH(totalSpent) },
          { label: 'Deviation',           value: (dev >= 0 ? '+' : '') + fmtH(Math.abs(dev)) },
          { label: 'Saving (raw)',        value: (rawPct >= 0 ? '+' : '') + rawPct + '%',
            hi: true, negative: rawNeg },
          { label: 'Saving excl. UAT R1', value: (exclPctVal >= 0 ? '+' : '') + exclPctVal + '%',
            hi: true, negative: exclNeg }
        ],
        tableRows: summaryRows
      };
    }

    // ------ Consolidated data ------
    {
      const CATS = ALMAParser.CATEGORY_KEYS;
      const totals = {};
      CATS.forEach(function(k) { totals[k] = { estimated: 0, spent: 0 }; });
      let grandEst = 0, grandSpent = 0, r1UATSpent = 0;

      releaseOrder.forEach(function(key) {
        const rel = releases[key];
        CATS.forEach(function(k) {
          totals[k].estimated += rel.regular[k].estimated;
          totals[k].spent     += rel.regular[k].spent;
        });
        if (rel.label === 'R1' && rel.totals.uatSpent > 0) {
          totals['DEVELOPMENT'].spent += rel.totals.uatSpent;
          r1UATSpent = rel.totals.uatSpent;
        }
        grandEst   += rel.totals.regular.estimated;
        grandSpent += rel.totals.all.spent;
      });

      const DISPLAY = [
        { key: 'ANALYSIS',    label: 'Analysis & DF' },
        { key: 'DEVELOPMENT', label: 'DT / Const' },
        { key: 'TESTING',     label: 'Testing' },
        { key: 'DEPLOY',      label: 'Deploy' },
        { key: 'PRODUCTION',  label: 'Production' },
        { key: 'MANAGEMENT',  label: 'Management' }
      ];

      const consoRows = DISPLAY.map(function(d) {
        const est   = Math.round(totals[d.key].estimated * 10) / 10;
        const spent = Math.round(totals[d.key].spent     * 10) / 10;
        if (est === 0 && spent === 0) return [d.label, '0', '0', '0', '—', '—'];
        const dev2 = Math.round((est - spent) * 10) / 10;
        const pct  = est > 0 ? Math.round((est - spent) / est * 100) : 0;
        let exclStr = '—';
        if (d.key === 'DEVELOPMENT' && r1UATSpent > 0) {
          const spExcl   = Math.round((spent - r1UATSpent) * 10) / 10;
          const exclPct2 = est > 0 ? Math.round((est - spExcl) / est * 100) : 0;
          exclStr = (exclPct2 >= 0 ? '+' : '') + exclPct2 + '%';
        }
        return [
          d.label, fmt(est), fmt(spent),
          (dev2 >= 0 ? '+' : '') + fmt(dev2),
          (pct  >= 0 ? '+' : '') + pct + '%',
          exclStr
        ];
      });

      grandEst   = Math.round(grandEst   * 10) / 10;
      grandSpent = Math.round(grandSpent * 10) / 10;
      const gDev   = Math.round((grandEst - grandSpent) * 10) / 10;
      const gPct   = grandEst > 0 ? Math.round((grandEst - grandSpent) / grandEst * 100) : 0;
      const exclSp = Math.round((grandSpent - r1UATSpent) * 10) / 10;
      const exPct  = grandEst > 0 ? Math.round((grandEst - exclSp) / grandEst * 100) : 0;
      consoRows.push([
        'TOTAL', fmt(grandEst), fmt(grandSpent),
        (gDev  >= 0 ? '+' : '') + fmt(gDev),
        gPct + '% raw',
        (exPct >= 0 ? '+' : '') + exPct + '% excl. UAT R1'
      ]);

      const testPct  = totals.TESTING.estimated > 0 ?
        Math.round((totals.TESTING.estimated - totals.TESTING.spent) / totals.TESTING.estimated * 100) : 0;
      const prodPct  = totals.PRODUCTION.estimated > 0 ?
        Math.round((totals.PRODUCTION.estimated - totals.PRODUCTION.spent) / totals.PRODUCTION.estimated * 100) : 0;
      const mgmtPct  = totals.MANAGEMENT.estimated > 0 ?
        Math.round((totals.MANAGEMENT.estimated - totals.MANAGEMENT.spent) / totals.MANAGEMENT.estimated * 100) : 0;
      const mgmtOver = mgmtPct < 0;
      const mgmtLbl  = mgmtOver ? 'Mgmt Over Budget' : 'Mgmt Saving';
      const mgmtVal  = mgmtOver ?
        Math.abs(Math.round((totals.MANAGEMENT.spent / Math.max(1, totals.MANAGEMENT.estimated) - 1) * 100)) + '%' :
        (mgmtPct >= 0 ? '+' : '') + mgmtPct + '%';

      rd.consolidated = {
        metrics: [
          { label: 'Total Estimated',   value: fmtH(grandEst) },
          { label: 'Total Spent',       value: fmtH(grandSpent) },
          { label: 'Testing Saving',    value: (testPct >= 0 ? '+' : '') + testPct + '%',
            hi: true, negative: testPct < 0 },
          { label: 'Production Saving', value: (prodPct >= 0 ? '+' : '') + prodPct + '%',
            hi: true, negative: prodPct < 0 },
          { label: mgmtLbl, value: mgmtVal, hi: true, negative: mgmtOver }
        ],
        tableRows: consoRows
      };
    }

    // ------ Velocity data ------
    {
      const rels       = releaseOrder.map(function(k) { return releases[k]; });
      const VROWS = [
        { key: 'DEVELOPMENT', label: 'DT / Const' },
        { key: 'TESTING',     label: 'Testing' },
        { key: 'ANALYSIS',    label: 'Analysis & DF' },
        { key: 'MANAGEMENT',  label: 'Management' }
      ];
      const velHeaders = ['Area'].concat(rels.map(function(r) { return r.label + ' Utilization'; })).concat(['Insight']);
      const velRows    = VROWS.map(function(d) {
        const cells = [d.label];
        rels.forEach(function(rel) {
          const est   = rel.regular[d.key].estimated;
          const spent = rel.regular[d.key].spent;
          if (!est || est === 0) { cells.push('Unbudgeted'); }
          else { cells.push(Math.round(spent / est * 100) + '%'); }
        });
        cells.push('—');
        return cells;
      });

      const lastRel    = rels[rels.length - 1] || { label: 'R3' };
      const firstLabel = rels[0] ? rels[0].label : 'R1';
      const lastLabel  = lastRel.label;
      const testPct2   = lastRel.regular ?
        (lastRel.regular.TESTING.estimated > 0 ?
          Math.round((lastRel.regular.TESTING.estimated - lastRel.regular.TESTING.spent) / lastRel.regular.TESTING.estimated * 100) : 0)
        : 0;
      const prodPct2   = lastRel.regular ?
        (lastRel.regular.PRODUCTION.estimated > 0 ?
          Math.round((lastRel.regular.PRODUCTION.estimated - lastRel.regular.PRODUCTION.spent) / lastRel.regular.PRODUCTION.estimated * 100) : 0)
        : 0;

      rd.velocity = {
        headers:  velHeaders,
        tableRows: velRows,
        metrics: [
          { label: 'DT/Const Improvement', value: firstLabel + '\u2192' + lastLabel, hi: true, negative: false },
          { label: 'Testing '    + lastLabel + ' Saving', value: (testPct2 >= 0 ? '+' : '') + testPct2 + '%', hi: true, negative: testPct2 < 0 },
          { label: 'Production ' + lastLabel + ' Saving', value: (prodPct2 >= 0 ? '+' : '') + prodPct2 + '%', hi: true, negative: prodPct2 < 0 }
        ]
      };
    }

    return rd;
  }

  // ------------------------------------------------------------------
  // Expose
  // ------------------------------------------------------------------
  global.ALMARenderer = {
    renderFromData:   renderFromData,
    buildReportData:  buildReportData
  };

})(window);
