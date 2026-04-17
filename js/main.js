// ── Password protection ──
(function(){
  if(sessionStorage.getItem('alma_auth') === '1') {
    document.getElementById('pw-overlay').classList.add('hidden');
  }
})();

function checkPw(){
  const val = document.getElementById('pw-input').value;
  if(val === 'ALMA2026'){
    sessionStorage.setItem('alma_auth','1');
    document.getElementById('pw-overlay').classList.add('hidden');
  } else {
    document.getElementById('pw-error').textContent = 'Incorrect password. Try again.';
    document.getElementById('pw-input').value = '';
    document.getElementById('pw-input').focus();
  }
}

const titles = ["Cover", "Summary", "R1", "R2", "R3", "Consolidated", "Velocity", "Simulation", "Risks", "Notes", "Gràcies"];
const TOTAL = 11;
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
    slide.addShape(pptx.ShapeType.rect, { x: 0, y: 5.0, w: '100%', h: 0.02, fill: { color: MID } });
    slide.addText('\u276F', { x: 0.2, y: 5.08, w: 0.25, h: 0.28, fontSize: 11, bold: true, color: PURPLE });
    slide.addText('Copyright \u00A9 2026 Accenture. All rights reserved.', { x: 0.5, y: 5.1, w: 8.5, h: 0.22, fontSize: 7.5, color: GRAY });
    slide.addText(String(num), { x: 9.5, y: 5.1, w: 0.3, h: 0.22, fontSize: 7.5, color: GRAY, align: 'right' });
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

  // ── Slide 0: Cover — Full-bleed purple ──
  const s0 = pptx.addSlide();

  // Full background
  s0.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 10, h: 5.63, fill: { color: PURPLE } });
  // Slightly darker left gradient overlay for depth
  s0.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 5, h: 5.63, fill: { color: PURPLE_DARK } });
  // Blend rectangle to smooth gradient
  s0.addShape(pptx.ShapeType.rect, { x: 3.5, y: 0, w: 2.5, h: 5.63, fill: { color: PURPLE } });

  // Watermark arrows (right side, subtle)
  s0.addText('\u276F', { x: 4.8, y: -1.0, w: 7.0, h: 7.5, fontSize: 420, bold: true, color: '9353A0', align: 'center', valign: 'middle' });
  s0.addText('\u276F', { x: 6.0, y: 0.0, w: 5.0, h: 5.8, fontSize: 280, bold: false, color: 'BD96C5', align: 'center', valign: 'middle' });

  // Tag
  s0.addText('ACCENTURE \u00B7 CTTI \u00B7 APRIL 2026', { x: 0.55, y: 0.65, w: 6.5, h: 0.25, fontSize: 8, bold: true, color: PURPLE_ACCENT, charSpacing: 2 });
  // Title
  s0.addText('Release Performance\nAnalysis & Forward Simulation', { x: 0.55, y: 1.05, w: 6.8, h: 1.75, fontSize: 28, bold: true, color: WHITE, breakLine: true, valign: 'top' });
  // Divider
  s0.addShape(pptx.ShapeType.rect, { x: 0.55, y: 2.92, w: 0.5, h: 0.04, fill: { color: PURPLE_ACCENT } });
  // Subtitle
  s0.addText('R1\u2013R3 Actuals \u00B7 R4\u2013R6 Projections', { x: 0.55, y: 3.06, w: 6.8, h: 0.26, fontSize: 11, bold: true, color: 'D9C0E8' });
  // Meta
  s0.addText('Extraction date: March 24, 2026 \u00B7 Data subject to update', { x: 0.55, y: 3.38, w: 6.8, h: 0.2, fontSize: 8.5, color: 'A888BB' });

  // Badges
  s0.addShape(pptx.ShapeType.roundRect, { x: 0.55, y: 3.74, w: 2.1, h: 0.3, fill: { color: '5C2270' }, line: { color: '9B5CB0' }, rectRadius: 0.15 });
  s0.addText('+34% Saving excl. UAT R1', { x: 0.55, y: 3.74, w: 2.1, h: 0.3, fontSize: 7.5, bold: true, color: WHITE, align: 'center', valign: 'middle' });
  s0.addShape(pptx.ShapeType.roundRect, { x: 2.78, y: 3.74, w: 2.15, h: 0.3, fill: { color: '5C2270' }, line: { color: '9B5CB0' }, rectRadius: 0.15 });
  s0.addText('~2,950h Projected R4\u2013R6', { x: 2.78, y: 3.74, w: 2.15, h: 0.3, fontSize: 7.5, bold: true, color: WHITE, align: 'center', valign: 'middle' });
  s0.addShape(pptx.ShapeType.roundRect, { x: 5.06, y: 3.74, w: 1.85, h: 0.3, fill: { color: '5C2270' }, line: { color: '9B5CB0' }, rectRadius: 0.15 });
  s0.addText('3 Releases Analysed', { x: 5.06, y: 3.74, w: 1.85, h: 0.3, fontSize: 7.5, bold: true, color: WHITE, align: 'center', valign: 'middle' });

  // Footer separator
  s0.addShape(pptx.ShapeType.rect, { x: 0, y: 5.0, w: 10, h: 0.015, fill: { color: '9B5CB0' } });
  // Footer brand: arrow icon
  s0.addText('\u276F', { x: 0.28, y: 5.05, w: 0.3, h: 0.38, fontSize: 14, bold: true, color: WHITE });
  // Brand text
  s0.addText('Generalitat de Catalunya', { x: 0.7, y: 5.07, w: 4.2, h: 0.15, fontSize: 6.5, color: 'CCAADD' });
  s0.addText('Centre de Telecomunicacions i Tecnologies de la Informaci\u00F3', { x: 0.7, y: 5.22, w: 4.2, h: 0.14, fontSize: 6, bold: true, color: 'CCAADD' });
  // Copyright
  s0.addText('Copyright \u00A9 2026 Accenture. All rights reserved.', { x: 5.8, y: 5.1, w: 4.0, h: 0.26, fontSize: 7, color: 'A888BB', align: 'right' });

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
  addHeader(s2, 'R1 — Deep Dive \u00B7 Lessons Learned', 'Completed \u00B7 16 Dec 2025 \u2013 31 Jan 2026 \u00B7 UAT R1 included \u00B7 All estimates use x4 factor');
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
    null
  );
  // R1 note
  s2.addShape(pptx.ShapeType.rect, { x: 0.4, y: 3.72, w: 9.2, h: 0.3, fill: { color: 'F5F0F8' }, line: { color: PURPLE, width: 1.5 } });
  s2.addText('\u26A0\uFE0F  DT hours (3,183h) include 2,111h of UAT R1: technical debt + UAT December + R2 UAT. UAT R1: Est. 1,515h \u00B7 Spent 1,072h \u00B7 Saving +443h (+29%).', { x: 0.55, y: 3.74, w: 8.9, h: 0.26, fontSize: 7.5, color: '333333' });
  // R1 insight boxes
  s2.addShape(pptx.ShapeType.rect, { x: 0.4, y: 4.1, w: 4.5, h: 0.72, fill: { color: 'F0FDF4' }, line: { color: GREEN, width: 1.5 } });
  s2.addText('\uD83D\uDCA1', { x: 0.52, y: 4.16, w: 0.32, h: 0.28, fontSize: 12 });
  s2.addText('Key Insight', { x: 0.9, y: 4.13, w: 3.8, h: 0.2, fontSize: 9, bold: true, color: GREEN });
  s2.addText('Without UAT burden, DT/Const was ~1,072h vs 1,515h estimated \u2014 revealing +29% real team efficiency.', { x: 0.9, y: 4.33, w: 3.8, h: 0.42, fontSize: 7.5, color: '14532D' });
  s2.addShape(pptx.ShapeType.rect, { x: 5.1, y: 4.1, w: 4.5, h: 0.72, fill: { color: 'FFF7ED' }, line: { color: 'F97316', width: 1.5 } });
  s2.addText('\uD83D\uDEA9', { x: 5.22, y: 4.16, w: 0.32, h: 0.28, fontSize: 12 });
  s2.addText('Management Flag', { x: 5.6, y: 4.13, w: 3.8, h: 0.2, fontSize: 9, bold: true, color: 'F97316' });
  s2.addText('434.5h of unbudgeted management effort highlights a gap in the estimation framework that must be addressed.', { x: 5.6, y: 4.33, w: 3.8, h: 0.42, fontSize: 7.5, color: '7C2D12' });
  addFooter(s2, 2);

  // ── Slide 3: R2 ──
  const s3 = pptx.addSlide();
  addHeader(s3, 'R2 — Deep Dive \u00B7 Efficiency Gains', 'Completed \u00B7 1 Feb \u2013 16 Feb 2026 \u00B7 Technical debt resolved \u00B7 Strongest DT/Const improvement');
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
    null
  );
  // R2 note
  s3.addShape(pptx.ShapeType.rect, { x: 0.4, y: 3.72, w: 9.2, h: 0.3, fill: { color: 'F5F0F8' }, line: { color: PURPLE, width: 1.5 } });
  s3.addText('\u26A0\uFE0F  R2 DT/Const estimation was duplicated. Only DT effort was considered in the estimated figure. Real spent reflects both DT and Const hours.', { x: 0.55, y: 3.74, w: 8.9, h: 0.26, fontSize: 7.5, color: '333333' });
  // R2 insight boxes
  s3.addShape(pptx.ShapeType.rect, { x: 0.4, y: 4.1, w: 4.5, h: 0.72, fill: { color: 'F0FDF4' }, line: { color: GREEN, width: 1.5 } });
  s3.addText('\u2B50', { x: 0.52, y: 4.16, w: 0.32, h: 0.28, fontSize: 12 });
  s3.addText('DT/Const Breakthrough', { x: 0.9, y: 4.13, w: 3.8, h: 0.2, fontSize: 9, bold: true, color: GREEN });
  s3.addText('With technical debt resolved, DT/Const achieved +55% saving \u2014 strongest area improvement, confirming growing velocity.', { x: 0.9, y: 4.33, w: 3.8, h: 0.42, fontSize: 7.5, color: '14532D' });
  s3.addShape(pptx.ShapeType.rect, { x: 5.1, y: 4.1, w: 4.5, h: 0.72, fill: { color: 'F0FDF4' }, line: { color: GREEN, width: 1.5 } });
  s3.addText('\uD83D\uDCC9', { x: 5.22, y: 4.16, w: 0.32, h: 0.28, fontSize: 12 });
  s3.addText('Management Maturing', { x: 5.6, y: 4.13, w: 3.8, h: 0.2, fontSize: 9, bold: true, color: GREEN });
  s3.addText('Management overhead dropped from 434.5h to 166h \u2014 a 62% reduction showing team coordination is maturing.', { x: 5.6, y: 4.33, w: 3.8, h: 0.42, fontSize: 7.5, color: '14532D' });
  addFooter(s3, 3);

  // ── Slide 4: R3 ──
  const s4 = pptx.addSlide();
  addHeader(s4, 'R3 — Status \u00B7 In Progress', '17 Feb \u2013 6 Apr 2026 \u00B7 Partial data \u00B7 March 24, 2026 extraction');
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
    null
  );
  // Consolidated note
  s5.addShape(pptx.ShapeType.rect, { x: 0.4, y: 3.72, w: 9.2, h: 0.3, fill: { color: 'F5F0F8' }, line: { color: PURPLE, width: 1.5 } });
  s5.addText('* Saving excl. UAT R1 = 1-(4,854h / 7,337.9h) where 4,854h = total spent (6,965h) minus UAT R1 hours (2,111h).', { x: 0.55, y: 3.74, w: 8.9, h: 0.26, fontSize: 7.5, color: '333333' });
  // Consolidated insight boxes
  s5.addShape(pptx.ShapeType.rect, { x: 0.4, y: 4.1, w: 4.5, h: 0.72, fill: { color: 'F0FDF4' }, line: { color: GREEN, width: 1.5 } });
  s5.addText('\uD83D\uDCC8', { x: 0.52, y: 4.16, w: 0.32, h: 0.28, fontSize: 12 });
  s5.addText('Utilization Trend', { x: 0.9, y: 4.13, w: 3.8, h: 0.2, fontSize: 9, bold: true, color: GREEN });
  s5.addText('R1: 88% utilization (excl. UAT) \u2192 R2: 70% utilization. Declining rate signals higher team efficiency across releases.', { x: 0.9, y: 4.33, w: 3.8, h: 0.42, fontSize: 7.5, color: '14532D' });
  s5.addShape(pptx.ShapeType.rect, { x: 5.1, y: 4.1, w: 4.5, h: 0.72, fill: { color: 'FFF7ED' }, line: { color: 'F97316', width: 1.5 } });
  s5.addText('\u26A0\uFE0F', { x: 5.22, y: 4.16, w: 0.32, h: 0.28, fontSize: 12 });
  s5.addText('Management Gap', { x: 5.6, y: 4.13, w: 3.8, h: 0.2, fontSize: 9, bold: true, color: 'F97316' });
  s5.addText('889.5h spent vs 40h estimated across R1\u2013R3. 2,124% over budget \u2014 largest unbudgeted cost driver requiring immediate action.', { x: 5.6, y: 4.33, w: 3.8, h: 0.42, fontSize: 7.5, color: '7C2D12' });
  addFooter(s5, 5);

  // ── Slide 6: Velocity Trend ──
  const s6 = pptx.addSlide();
  addHeader(s6, 'Velocity Trend Analysis', '18pp efficiency improvement R1\u2192R2 \u00B7 Basis for forward simulation');
  addMetrics(s6, [
    { label: 'DT/Const Improvement', value: '26pp', hi: true },
    { label: 'Analysis & DF Trend',  value: '\u2192 Est.', hi: true },
    { label: 'Mgmt Overhead Drop',   value: '-62%', hi: true }
  ], 0.82);
  addTable(s6,
    ['Area', 'R1 Utilization', 'R2 Utilization', 'Trend', 'Insight'],
    [
      ['DT / Const',    '71%',          '45%',   '-26pp', 'Strongest improvement \u2014 debt resolved'],
      ['Testing',       '76%',          '83%',   '+7pp',  'Slight regression \u2014 monitor R3'],
      ['Analysis & DF', '79%',          '98%',   '+19pp', 'Converging to estimate \u2014 scope growing'],
      ['Management',    'Unbudgeted',   '166h',  '-62%',  'Team coordination maturing']
    ],
    1.7,
    null
  );
  s6.addShape(pptx.ShapeType.rect, { x: 0.4, y: 3.9, w: 4.5, h: 0.9, fill: { color: 'F0FDF4' }, line: { color: GREEN, width: 1.5 } });
  s6.addText('\uD83D\uDCCA', { x: 0.52, y: 3.96, w: 0.32, h: 0.3, fontSize: 12 });
  s6.addText('Overall Velocity', { x: 0.9, y: 3.94, w: 3.8, h: 0.22, fontSize: 9, bold: true, color: GREEN });
  s6.addText('The team improved efficiency by 18pp from R1 to R2, forming the basis for R4\u2013R6 projections.', { x: 0.9, y: 4.16, w: 3.8, h: 0.56, fontSize: 8, color: '14532D' });
  s6.addShape(pptx.ShapeType.rect, { x: 5.1, y: 3.9, w: 4.5, h: 0.9, fill: { color: 'FFF7ED' }, line: { color: 'F97316', width: 1.5 } });
  s6.addText('\uD83D\uDD0D', { x: 5.22, y: 3.96, w: 0.32, h: 0.3, fontSize: 12 });
  s6.addText('Watch Areas', { x: 5.6, y: 3.94, w: 3.8, h: 0.22, fontSize: 9, bold: true, color: 'F97316' });
  s6.addText('Testing showed slight regression (+7pp) and Analysis & DF is converging to estimate. Both need attention in R4.', { x: 5.6, y: 4.16, w: 3.8, h: 0.56, fontSize: 8, color: '7C2D12' });
  addFooter(s6, 6);

  // ── Slide 7: Forward Simulation ──
  const s7 = pptx.addSlide();
  addHeader(s7, 'R4\u2013R6 Forward Simulation', 'Based on current velocity trends with diminishing returns \u00B7 ~7,500h total budget');
  addMetrics(s7, [
    { label: 'Cumulative Saved R4\u2013R6', value: '~2,950h', hi: true },
    { label: 'Avg Saving R4\u2013R6',       value: '+39%',    hi: true },
    { label: 'Mgmt Budget / Release',       value: '150\u2013200h' }
  ], 0.82);
  addTable(s7,
    ['Release', 'Estimated (h)', 'Projected Spent (h)', 'Projected Saving', 'Utilization', 'Note'],
    [
      ['R4',          '~2,500', '~1,625', '+35%',    '65%',  '+5pp improvement over R2'],
      ['R5',          '~2,500', '~1,500', '+40%',    '60%',  'Team fully mature'],
      ['R6',          '~2,500', '~1,425', '+43%',    '57%',  'Diminishing returns \u2014 near plateau'],
      ['TOTAL R4\u2013R6', '~7,500', '~4,550', '~+2,950h', '~61%', 'Avg +39% saving']
    ],
    1.7,
    null
  );
  s7.addShape(pptx.ShapeType.rect, { x: 0.4, y: 3.9, w: 4.5, h: 0.9, fill: { color: 'F0FDF4' }, line: { color: GREEN, width: 1.5 } });
  s7.addText('\uD83D\uDCC8', { x: 0.52, y: 3.96, w: 0.32, h: 0.3, fontSize: 12 });
  s7.addText('Efficiency Plateau', { x: 0.9, y: 3.94, w: 3.8, h: 0.22, fontSize: 9, bold: true, color: GREEN });
  s7.addText('Projections show diminishing returns approaching R6 \u2014 team nears peak efficiency at ~57% utilization. R5 = full maturity.', { x: 0.9, y: 4.16, w: 3.8, h: 0.56, fontSize: 8, color: '14532D' });
  s7.addShape(pptx.ShapeType.rect, { x: 5.1, y: 3.9, w: 4.5, h: 0.9, fill: { color: 'FFF7ED' }, line: { color: 'F97316', width: 1.5 } });
  s7.addText('\uD83D\uDCBC', { x: 5.22, y: 3.96, w: 0.32, h: 0.3, fontSize: 12 });
  s7.addText('Management Budget', { x: 5.6, y: 3.94, w: 3.8, h: 0.22, fontSize: 9, bold: true, color: 'F97316' });
  s7.addText('Recommend formalising 150\u2013200h management allocation per release in future estimates to prevent recurring overruns.', { x: 5.6, y: 4.16, w: 3.8, h: 0.56, fontSize: 8, color: '7C2D12' });
  addFooter(s7, 7);

  // ── Slide 8: Risks & Recommendations ──
  const s8 = pptx.addSlide();
  addHeader(s8, 'Risks & Recommendations', 'Critical gaps to address to sustain projected efficiency gains');
  // Column headers
  s8.addShape(pptx.ShapeType.rect, { x: 0.4, y: 0.82, w: 4.4, h: 0.28, fill: { color: 'FEF2F2' }, line: { color: RED, width: 1 } });
  s8.addText('\u26A0\uFE0F  Key Risks', { x: 0.55, y: 0.84, w: 4.1, h: 0.22, fontSize: 9, bold: true, color: RED });
  s8.addShape(pptx.ShapeType.rect, { x: 5.1, y: 0.82, w: 4.5, h: 0.28, fill: { color: 'F0FDF4' }, line: { color: GREEN, width: 1 } });
  s8.addText('\u2705  Recommendations', { x: 5.25, y: 0.84, w: 4.1, h: 0.22, fontSize: 9, bold: true, color: GREEN });
  // Risks items
  const risks = [
    { icon: '\uD83D\uDCB0', title: 'Management Overrun',        text: '889.5h spent vs 40h estimated across R1\u2013R3 \u2014 2,124% over budget. Formalise a management allocation of 150\u2013200h per release.' },
    { icon: '\uD83D\uDD01', title: 'Technical Debt Recurrence', text: "R1's 2,111h UAT burden shows how inherited debt can mask true performance. Implement technical debt tracking per release." },
    { icon: '\uD83D\uDCCB', title: 'R3 Estimation Gaps',        text: '7 USs still without estimates \u2014 final R3 figures could shift the baseline for simulation.' },
    { icon: '\uD83D\uDE80', title: 'Deploy at 0h Actual',       text: '186h estimated but 0h spent across all releases. Review if deploy effort is captured elsewhere or if estimates should be adjusted.' }
  ];
  risks.forEach((r, i) => {
    const y = 1.18 + i * 0.86;
    const bg = i % 2 === 0 ? 'FFFFFF' : 'F3F4F6';
    s8.addShape(pptx.ShapeType.rect, { x: 0.4, y, w: 4.4, h: 0.78, fill: { color: bg }, line: { color: PURPLE, width: 1.5 } });
    s8.addShape(pptx.ShapeType.rect, { x: 0.4, y, w: 0.06, h: 0.78, fill: { color: PURPLE } });
    s8.addText(r.icon,  { x: 0.52, y: y + 0.08, w: 0.32, h: 0.3, fontSize: 13 });
    s8.addText(r.title, { x: 0.9,  y: y + 0.06, w: 3.8,  h: 0.22, fontSize: 9, bold: true, color: PURPLE });
    s8.addText(r.text,  { x: 0.9,  y: y + 0.28, w: 3.8,  h: 0.44, fontSize: 7.5, color: '444444' });
  });
  // Recommendations items
  const recs = [
    { icon: '\uD83D\uDCD0', title: 'Re-baseline Estimation Methodology', text: 'The current x4 multiplication factor overestimates by 30\u201340% for mature releases. Adjust to reflect actual team velocity.' },
    { icon: '\uD83D\uDCC5', title: 'Quarterly Capacity Review',           text: 'Conduct quarterly reviews to validate simulation assumptions against actuals and recalibrate projections. Addressing risks now will protect the projected ~2,950h in cumulative R4\u2013R6 savings.' }
  ];
  recs.forEach((r, i) => {
    const y = 1.18 + i * 1.72;
    const bg = i % 2 === 0 ? 'FFFFFF' : 'F3F4F6';
    s8.addShape(pptx.ShapeType.rect, { x: 5.1, y, w: 4.5, h: 1.64, fill: { color: bg }, line: { color: PURPLE, width: 1.5 } });
    s8.addShape(pptx.ShapeType.rect, { x: 5.1, y, w: 0.06, h: 1.64, fill: { color: PURPLE } });
    s8.addText(r.icon,  { x: 5.22, y: y + 0.12, w: 0.32, h: 0.3, fontSize: 13 });
    s8.addText(r.title, { x: 5.6,  y: y + 0.1,  w: 3.8,  h: 0.28, fontSize: 9, bold: true, color: PURPLE });
    s8.addText(r.text,  { x: 5.6,  y: y + 0.38, w: 3.8,  h: 1.18, fontSize: 8, color: '444444' });
  });
  addFooter(s8, 8);

  // ── Slide 9: Notes ──
  const s9 = pptx.addSlide();
  addHeader(s9, 'Notes & Clarifications', 'Context for accurate interpretation of the metrics');
  const notes = [
    { icon: '\uD83D\uDCCC', title: 'Data Extraction Date',                   text: 'Metrics extracted on Tuesday, March 24th. Subsequent ALMA reporting may show updated figures. This report reflects a snapshot subject to change in the next cycle.' },
    { icon: '\u2716\uFE0F', title: 'Estimation Methodology \u2013 x4 Factor', text: 'All total estimations include a 4x multiplication factor applied to the base technical estimation. E.g. if a US is technically estimated at 10h, the total estimation used is 40h, accounting for full delivery effort across all areas.' },
    { icon: '\u26A0\uFE0F', title: 'R1 \u2013 DT Hours & UAT R1',             text: 'DT hours in R1 (3,183h) include 2,111h of UAT R1: (1) technical debt from previous UATs + UAT from December, and (2) R2 UAT hours logged here because the R2 ticket hadn\'t been created and ALMA UAT reporting wasn\'t aligned with the team.' },
    { icon: '\u26A0\uFE0F', title: 'R2 \u2013 DT/Const Estimation',            text: 'R2 DT and Const estimation was duplicated. Only DT effort was used as the estimated figure. Real spent cost reflects both DT and Const actual hours.' },
    { icon: '\uD83D\uDD04', title: 'R3 \u2013 Still In Progress',              text: '2 USs are under technical revision and 5 are in functional analysis with no estimated values. These will be incorporated once complete.' },
    { icon: '\uD83D\uDCCA', title: 'Forward Simulation Assumptions',           text: 'R4\u2013R6 projections are based on observed velocity trends with diminishing returns. Assumes no new technical debt, stable team composition, and management budgets of 150\u2013200h per release.' }
  ];
  const noteItemH = 0.58;
  const noteGap   = 0.04;
  const noteStart = 0.76;
  notes.forEach((n, i) => {
    const y = noteStart + i * (noteItemH + noteGap);
    const bg = i % 2 === 0 ? LIGHT : WHITE;
    s9.addShape(pptx.ShapeType.rect, { x: 0.4, y, w: 9.2, h: noteItemH, fill: { color: bg }, line: { color: PURPLE, width: 1.5 } });
    s9.addShape(pptx.ShapeType.rect, { x: 0.4, y, w: 0.06, h: noteItemH, fill: { color: PURPLE } });
    s9.addText(n.icon,  { x: 0.52, y: y + 0.08, w: 0.35, h: 0.26, fontSize: 11 });
    s9.addText(n.title, { x: 0.95, y: y + 0.04, w: 8.5,  h: 0.18, fontSize: 8.5, bold: true, color: PURPLE });
    s9.addText(n.text,  { x: 0.95, y: y + 0.22, w: 8.5,  h: 0.32, fontSize: 7, color: '444444' });
  });
  addFooter(s9, 9);

  // ── Slide 10: Gràcies ──
  const s10 = pptx.addSlide();
  s10.background = { color: PURPLE };
  // Watermark 1 — large faint filled arrow (wm1)
  s10.addText('\u276F', { x: 5.5, y: -0.2, w: 5.5, h: 6, fontSize: 420, bold: true, color: PURPLE_DARK, align: 'center', valign: 'middle', transparency: 82 });
  // Watermark 2 — medium outline-style arrow (wm2)
  s10.addText('\u276F', { x: 6.2, y: 0.3, w: 3.5, h: 5, fontSize: 260, bold: false, color: WHITE, align: 'center', valign: 'middle', transparency: 50 });
  // Main text
  s10.addText('Gr\u00E0cies', { x: 0.8, y: 1.8, w: 6, h: 2, fontSize: 54, bold: true, color: WHITE });
  // Footer copyright
  s10.addText('Copyright \u00A9 2026 Accenture. All rights reserved.', { x: 0.8, y: 5.1, w: 9, h: 0.28, fontSize: 9, color: PURPLE_ACCENT });

  pptx.writeFile({ fileName: 'Release_Performance_Analysis.pptx' });
}
