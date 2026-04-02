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
