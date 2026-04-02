const titles = ["Cover", "Summary", "R1", "R2", "R3", "Consolidated", "Notes", "Gràcies"];
const TOTAL = titles.length;
let cur = 0;

function goTo(i) {
  document.getElementById('s' + cur).classList.remove('active');
  document.querySelectorAll('#controls button:not(.nav)')[cur].classList.remove('active');
  cur = i;
  document.getElementById('s' + cur).classList.add('active');
  document.querySelectorAll('#controls button:not(.nav)')[cur].classList.add('active');
  document.getElementById('snum').textContent = cur + 1;
  document.getElementById('stitle').textContent = titles[cur];
  document.getElementById('prev').disabled = cur === 0;
  document.getElementById('next').disabled = cur === TOTAL - 1;
}

function go(d) {
  goTo(Math.min(TOTAL - 1, Math.max(0, cur + d)));
}

// Keyboard navigation
document.addEventListener('keydown', (e) => {
  if (e.key === 'ArrowRight' || e.key === 'ArrowDown') go(1);
  if (e.key === 'ArrowLeft'  || e.key === 'ArrowUp')   go(-1);
});

// Init
document.getElementById('prev').disabled = true;

// ── Fullscreen ──
function toggleFullscreen() {
  const frame = document.getElementById('frame');
  const btn   = document.getElementById('fsBtn');
  if (!document.fullscreenElement) {
    frame.requestFullscreen().catch(() => {});
  } else {
    document.exitFullscreen();
  }
}
document.addEventListener('fullscreenchange', () => {
  const btn = document.getElementById('fsBtn');
  btn.textContent = document.fullscreenElement ? '✕ Exit Fullscreen' : '⛶ Fullscreen';
});

// ── PPTX Export ──
function downloadPptx() {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_16x9';

  const PURPLE = '7B2D8B', PURPLE_ACCENT = 'C084D8', PURPLE_DARK = '5C1F6B';
  const WHITE = 'FFFFFF', BLACK = '000000', GRAY = '6B7280';
  const GREEN = '16A34A', RED = 'DC2626';
  const LIGHT = 'F3F4F6', MID = 'E5E7EB';

  function addHeader(slide, title, sub) {
    slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0.62, w: '100%', h: 0.05, fill: { color: PURPLE } });
    slide.addText(title, { x: 0.4, y: 0.06, w: 9.2, h: 0.42, fontSize: 20, bold: true, color: BLACK });
    slide.addText(sub,   { x: 0.4, y: 0.46, w: 9.2, h: 0.2,  fontSize: 11, bold: true, color: PURPLE });
  }

  function addFooter(slide, num) {
    slide.addShape(pptx.ShapeType.rect, { x: 0, y: 5.1, w: '100%', h: 0.01, fill: { color: MID } });
    slide.addText('Copyright © 2025 Accenture. All rights reserved.', { x: 0.3, y: 5.15, w: 7, h: 0.25, fontSize: 8, color: GRAY });
    slide.addText(String(num), { x: 9.5, y: 5.15, w: 0.3, h: 0.25, fontSize: 8, color: GRAY, align: 'right' });
  }

  function addMetrics(slide, metrics, y) {
    const w = 9.2 / metrics.length;
    metrics.forEach((m, i) => {
      const x = 0.4 + i * w;
      slide.addShape(pptx.ShapeType.roundRect, { x, y, w: w - 0.12, h: 0.72, fill: { color: m.hi ? PURPLE : LIGHT }, rectRadius: 0.05 });
      slide.addText(m.label, { x, y: y + 0.05, w: w - 0.12, h: 0.2, fontSize: 7, bold: true, color: m.hi ? PURPLE_ACCENT : GRAY, align: 'center' });
      slide.addText(m.value, { x, y: y + 0.27, w: w - 0.12, h: 0.35, fontSize: 15, bold: true, color: m.hi ? WHITE : BLACK, align: 'center' });
    });
  }

  function cellColor(v, isFirst, isLast) {
    if (isFirst || isLast) return BLACK;
    if (typeof v === 'string' && v.startsWith('+')) return GREEN;
    if (typeof v === 'string' && v.startsWith('-') && v !== '—') return RED;
    return BLACK;
  }

  function addTable(slide, headers, rows, y, note) {
    const colW = [2.8, ...Array(headers.length - 1).fill((9.2 - 2.8) / (headers.length - 1))];
    const tableRows = [
      headers.map(h => ({ text: h, options: { bold: true, color: WHITE, fill: { color: PURPLE }, align: h === headers[0] ? 'left' : 'right', fontSize: 9 } })),
      ...rows.map((row, ri) => row.map((cell, ci) => ({
        text: String(cell),
        options: { color: cellColor(cell, ci === 0, ri === rows.length - 1), align: ci === 0 ? 'left' : 'right', fontSize: 9, bold: ci === 0 || ri === rows.length - 1, fill: { color: ri % 2 === 0 ? WHITE : LIGHT } }
      })))
    ];
    slide.addTable(tableRows, { x: 0.4, y, w: 9.2, colW, border: { color: MID } });
    if (note) {
      slide.addShape(pptx.ShapeType.rect, { x: 0.4, y: 4.6, w: 9.2, h: 0.38, fill: { color: 'F5F0F8' }, line: { color: PURPLE, width: 1.5 } });
      slide.addText(note, { x: 0.55, y: 4.62, w: 9.0, h: 0.34, fontSize: 8, color: '333333' });
    }
  }

  // ── Slide 0: Cover ──
  const s0 = pptx.addSlide();
  s0.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: 0.09, fill: { color: PURPLE } });
  s0.addText('Release Performance\nMetrics Report', { x: 0.8, y: 1.1, w: 6, h: 1.6, fontSize: 32, bold: true, color: BLACK, breakLine: true });
  s0.addText('R1 · R2 · R3 — Estimated vs. Actual Effort Analysis', { x: 0.8, y: 2.8, w: 7, h: 0.35, fontSize: 14, bold: true, color: PURPLE });
  s0.addText('Extraction date: January 27, 2025 · Data subject to update', { x: 0.8, y: 3.22, w: 7, h: 0.28, fontSize: 11, color: GRAY });
  s0.addText('Accenture', { x: 0.8, y: 3.52, w: 7, h: 0.28, fontSize: 11, color: GRAY });
  s0.addShape(pptx.ShapeType.rect, { x: 0, y: 5.05, w: '100%', h: 0.01, fill: { color: MID } });
  s0.addText('Generalitat de Catalunya · Centre de Telecomunicacions i Tecnologies de la Informació', { x: 0.3, y: 5.1, w: 7, h: 0.28, fontSize: 8, color: GRAY });
  s0.addText('Copyright © 2025 Accenture. All rights reserved.', { x: 7.2, y: 5.1, w: 2.5, h: 0.28, fontSize: 8, color: GRAY, align: 'right' });

  // ── Slide 1: Summary ──
  const s1 = pptx.addSlide();
  addHeader(s1, 'Executive Summary', 'Consolidated results across R1 + R2 + R3');
  addMetrics(s1, [
    { label: 'Total Estimated', value: '7,337.9h' },
    { label: 'Total Spent', value: '6,965.0h' },
    { label: 'Deviation', value: '+372.9h' },
    { label: 'Saving (raw)', value: '+5%', hi: true },
    { label: 'Saving excl. UAT R1', value: '+34%', hi: true }
  ], 0.82);
  addTable(s1,
    ['Release', 'Estimated (h)', 'Spent (h)', 'Deviation (h)', '% Saving'],
    [
      ['R1 – Completed (raw)',          '2,347.5', '4,183.0', '-1,835.5', '-78%'],
      ['R1 – Completed (excl. UAT R1)', '2,347.5', '2,072.0', '+275.5',   '+12%'],
      ['R2 – Completed',                '2,315.3', '1,611.0', '+704.3',   '+30%'],
      ['R3 – In Progress *',            '2,674.4', '1,171.0', '+1,503.4', '+56%']
    ],
    1.7,
    '* R3 still in progress — 2 USs under technical revision, 5 in functional analysis. Final figures may vary.'
  );
  addFooter(s1, 1);

  // ── Slide 2: R1 ──
  const s2 = pptx.addSlide();
  addHeader(s2, 'R1 — Release 1', 'Completed · UAT R1 included');
  addMetrics(s2, [
    { label: 'Total Estimated',   value: '2,347.5h' },
    { label: 'Total Spent (raw)', value: '4,183.0h' },
    { label: 'Spent excl. UAT R1', value: '2,072.0h' },
    { label: 'Saving excl. UAT', value: '+12%', hi: true },
    { label: 'UAT R1 Total Hours', value: '2,111.0h', hi: true }
  ], 0.82);
  addTable(s2,
    ['Area', 'Estimated (h)', 'Spent (h)', 'Deviation (h)', '% Saving'],
    [
      ['Analysis & DF', '151.5',   '119.5',   '+32.0',    '+21%'],
      ['DT / Const',    '1,515.0', '3,183.0', '-1,668.0', '-110%'],
      ['Testing',       '530.3',   '404.0',   '+126.3',   '+24%'],
      ['Deploy',        '75.75',   '0',       '+75.75',   '+100%'],
      ['Production',    '75.75',   '42.0',    '+33.75',   '+45%'],
      ['Management',    '0',       '434.5',   '-434.5',   '—'],
      ['TOTAL',         '2,347.5', '4,183.0', '-1,835.5', '-78% raw / +12% excl. UAT']
    ],
    1.7,
    'DT hours (3,183h) include 2,111h of UAT R1: technical debt from previous UATs + UAT December + R2 UAT. UAT R1: Est. 1,515h · Spent 1,072h · Saving +443h (+29%).'
  );
  addFooter(s2, 2);

  // ── Slide 3: R2 ──
  const s3 = pptx.addSlide();
  addHeader(s3, 'R2 — Release 2', 'Completed');
  addMetrics(s3, [
    { label: 'Total Estimated', value: '2,315.3h' },
    { label: 'Total Spent',     value: '1,611.0h' },
    { label: 'Deviation',       value: '+704.3h' },
    { label: '% Saving',        value: '+30%', hi: true }
  ], 0.82);
  addTable(s3,
    ['Area', 'Estimated (h)', 'Spent (h)', 'Deviation (h)', '% Saving'],
    [
      ['Analysis & DF', '220.5',   '215.5', '+5.0',    '+2%'],
      ['DT / Const',    '1,102.5', '498.0', '+604.5',  '+55%'],
      ['Testing',       '771.8',   '643.5', '+128.3',  '+17%'],
      ['Deploy',        '110.25',  '0',     '+110.25', '+100%'],
      ['Production',    '110.25',  '88.0',  '+22.25',  '+20%'],
      ['Management',    '0',       '166.0', '-166.0',  '—'],
      ['TOTAL',         '2,315.3', '1,611.0', '+704.3', '+30%']
    ],
    1.7,
    'R2 DT/Const estimation was duplicated. Only DT effort was considered in the estimated figure. Real spent reflects both DT and Const hours.'
  );
  addFooter(s3, 3);

  // ── Slide 4: R3 ──
  const s4 = pptx.addSlide();
  addHeader(s4, 'R3 — Release 3', 'In Progress');
  addMetrics(s4, [
    { label: 'Total Estimated',    value: '2,674.4h' },
    { label: 'Spent So Far',       value: '1,171.0h' },
    { label: 'Deviation (partial)', value: '+1,503.4h' },
    { label: '% Saving (partial)', value: '+56%', hi: true }
  ], 0.82);
  addTable(s4,
    ['Area', 'Estimated (h)', 'Spent (h)', 'Deviation (h)', '% Saving'],
    [
      ['Analysis & DF', '494.8', '373.0', '+121.8',   '+25%'],
      ['DT / Const',    '655.2', '264.5', '+390.7',   '+60%'],
      ['Testing',       '989.6', '221.5', '+768.1',   '+78%'],
      ['Deploy',        '0',     '0',     '0',         '—'],
      ['Production',    '494.8', '23.0',  '+471.8',   '+95%'],
      ['Management',    '40.0',  '289.0', '-249.0',   '-623%'],
      ['TOTAL',         '2,674.4', '1,171.0', '+1,503.4', '+56%']
    ],
    1.7,
    'R3 is still ongoing. 2 USs under technical revision and 5 in functional analysis with no estimated values yet.'
  );
  addFooter(s4, 4);

  // ── Slide 5: Consolidated ──
  const s5 = pptx.addSlide();
  addHeader(s5, 'R1 + R2 + R3 — Consolidated View', 'Combined effort across all releases');
  addMetrics(s5, [
    { label: 'Total Estimated',    value: '7,337.9h' },
    { label: 'Total Spent',        value: '6,965.0h' },
    { label: 'Saving (raw)',       value: '+5%',  hi: true },
    { label: 'Saving excl. UAT R1', value: '+34%', hi: true }
  ], 0.82);
  addTable(s5,
    ['Area', 'Estimated (h)', 'Spent (h)', 'Deviation (h)', '% Saving', '% w/ UAT R1'],
    [
      ['Analysis & DF', '866.8',   '708.0',   '+158.8',   '+16%',      '—'],
      ['DT / Const',    '3,272.7', '3,945.5', '-672.8',   '+48%',      '+1%'],
      ['Testing',       '2,291.6', '1,269.0', '+1,022.6', '+39%',      '—'],
      ['Deploy',        '186.0',   '0.0',     '+186.0',   '+100%',     '—'],
      ['Production',    '680.8',   '153.0',   '+527.8',   '+53%',      '—'],
      ['Management',    '40.0',    '889.5',   '-849.5',   '—',         '—'],
      ['TOTAL',         '7,337.9', '6,965.0', '+372.9',   '+5% raw',   '+34% excl. UAT R1']
    ],
    1.7,
    '* Saving excl. UAT R1 = 1-(4,854h / 7,337.9h) where 4,854h = total spent (6,965h) minus UAT R1 hours (2,111h).'
  );
  addFooter(s5, 5);

  // ── Slide 6: Notes ──
  const s6 = pptx.addSlide();
  addHeader(s6, 'Notes & Clarifications', 'Context for accurate interpretation of the metrics');
  const notes = [
    { title: 'Data Extraction Date',            text: 'Metrics extracted on Friday, January 27th. Subsequent ALMA reporting may show updated figures. This report reflects a snapshot subject to change.' },
    { title: 'Estimation Methodology – x4 Factor', text: 'All total estimations include a 4x multiplication factor. E.g. if a US is estimated at 10h, the total estimation used is 40h, covering full delivery effort.' },
    { title: 'R1 – DT Hours & UAT R1',           text: 'DT hours in R1 (3,183h) include 2,111h of UAT R1: technical debt + UAT December + R2 UAT hours logged here because the R2 ticket had not been created.' },
    { title: 'R2 – DT/Const Estimation',         text: 'R2 DT and Const estimation was duplicated. Only DT effort was used as the estimated figure. Real spent reflects both DT and Const actual hours.' },
    { title: 'R3 – Still In Progress',           text: '2 USs are under technical revision and 5 are in functional analysis with no estimated values yet. These will be incorporated once complete.' }
  ];
  notes.forEach((n, i) => {
    const y = 0.82 + i * 0.84;
    s6.addShape(pptx.ShapeType.rect, { x: 0.4, y, w: 9.2, h: 0.76, fill: { color: i % 2 === 0 ? LIGHT : WHITE }, line: { color: PURPLE, width: 1.5 } });
    s6.addText(n.title, { x: 0.6, y: y + 0.06, w: 8.8, h: 0.22, fontSize: 10, bold: true, color: PURPLE });
    s6.addText(n.text,  { x: 0.6, y: y + 0.3,  w: 8.8, h: 0.38, fontSize: 8.5, color: '444444' });
  });
  addFooter(s6, 6);

  // ── Slide 7: Gràcies ──
  const s7 = pptx.addSlide();
  s7.background = { color: PURPLE };
  s7.addText('Gràcies', { x: 0.8, y: 1.5, w: 8, h: 2, fontSize: 54, bold: true, color: WHITE });
  s7.addText('Copyright © 2025 Accenture. All rights reserved.', { x: 0.8, y: 5.1, w: 9, h: 0.28, fontSize: 9, color: PURPLE_ACCENT });

  pptx.writeFile({ fileName: 'Release_Performance_Metrics.pptx' });
}
