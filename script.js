/* ════════════════════════════════════════════════════════════════════
   script.js  v3  —  API-backed Timetable Extractor
   All Excel parsing is now done server-side (FastAPI + openpyxl).
   This file handles:
     • File upload  →  POST /upload
     • Teacher list →  GET  /teachers/{session_id}
     • Timetable    →  GET  /timetable/teacher|division|batch/...
     • Rendering the grid (unchanged visual logic)
     • PDF download (unchanged)
   ════════════════════════════════════════════════════════════════════ */

/* ─── CONFIG ────────────────────────────────────────────────────── */
const API_BASE = "http://127.0.0.1:8000";   // ← change if backend runs elsewhere

/* ─── FACULTY MAP (for display names not in API response) ─────── */
const FACULTY_MAP = {
  "UPM":  "Dr. Umesh Moharil",   "MRY":  "Dr. Meghna Yashwante",
  "MDB":  "Ms. Manisha Bhise",   "AGP":  "Dr. Amita Pal",
  "AGD":  "Dr. Anil Darekar",    "BDP":  "Dr. B D Patil",
  "PSD":  "Dr. Pratibha Desai",  "PMN":  "Dr. Poonam Nakhate",
  "MS":   "Mr. Mukesh Sharma",   "RBM":  "Mr. Rahul Mali",
  "HDV":  "Mr. Harshal Vaidya",  "PG":   "Mr. Pankaj Gaur",
  "VVK":  "Mr. Vishal Kulkarni", "RPD":  "Mr. R P Dharmale",
  "SIB":  "Mr. Sanket Barde",    "SG":   "Dr. Sandhya Gadge",
  "NG":   "Mr. Nikhil Gurav",    "PVM":  "Mrs. Pallavi Munde",
  "SK":   "Ms. Sheetal Khande",  "CJ":   "Dr. Chhaya Joshi",
  "TP":   "Mr. Tukaram Patil",   "ST":   "Ms. Shilpa Tambe",
  "NV":   "Mrs. Neha Verma",     "SM":   "Mrs. Sonali Murumkar",
  "SB":   "Ms. Swati Bagade",    "SD":   "Mr. Shankar Deshmukh",
  "MPP":  "Mr. Martand Pandagale",
};
function facultyName(code) {
  return FACULTY_MAP[String(code).trim().toUpperCase()] || String(code).trim();
}

/* ─── SLOT CONFIG (mirrors backend — kept in sync) ────────────── */
const SLOT_CONFIG = [
  { key: '8:30',  label: '8:30',  type: 'slot'  },
  { key: '9:30',  label: '9:30',  type: 'slot'  },
  { key: 'BRK',   label: 'Break', type: 'break' },
  { key: '10:45', label: '10:45', type: 'slot'  },
  { key: '11:45', label: '11:45', type: 'slot'  },
  { key: '12:45', label: '12:45', type: 'slot'  },
  { key: 'LCH',   label: 'Lunch', type: 'lunch' },
  { key: '1:30',  label: '1:30',  type: 'slot'  },
  { key: '2:30',  label: '2:30',  type: 'slot'  },
  { key: '3:30',  label: '3:30',  type: 'slot'  },
];
const LECTURE_SLOTS  = SLOT_CONFIG.filter(s => s.type === 'slot').map(s => s.key);
const FIXED_DAYS     = ['MON','TUE','WED','THU','FRI','SAT'];
const DAYS_FULL_MAP  = {
  MON:'Monday', TUE:'Tuesday', WED:'Wednesday',
  THU:'Thursday', FRI:'Friday', SAT:'Saturday',
};

const DIVISION_CONFIG = {
  A: { name: 'FE-A (Comp)',  batches: ['A1','A2','A3'] },
  B: { name: 'FE-B (Civil)', batches: ['B1','B2','B3'] },
  C: { name: 'FE-C (Comp)',  batches: ['C1','C2','C3'] },
  D: { name: 'FE-D (Mech)',  batches: ['D1','D2','D3'] },
  E: { name: 'FE-E (AI&DS)', batches: ['E1','E2','E3'] },
  F: { name: 'FE-F (Ro&AI)', batches: ['F1','F2','F3'] },
  G: { name: 'FE-G (AI&DS)', batches: ['G1','G2','G3'] },
  H: { name: 'FE-H (MTRX)',  batches: ['H1','H2','H3'] },
};

/* ─── STATE ─────────────────────────────────────────────────────── */
const state = {
  sessionId: null,
  teachers: [],          // [{code, display_name}]
  selectedTeacher: null,
  subjectColorMap: {}, colorIndex: 0, zoomLevel: 1,
  viewMode: 'teacher',   // 'teacher' | 'batch'
  selectedDiv: null, selectedBatch: null,
};

/* ─── DOM refs ──────────────────────────────────────────────────── */
const $el = id => document.getElementById(id);
const fileInput          = $el('fileInput');
const uploadBox          = $el('uploadBox');
const uploadMeta         = $el('uploadMeta');
const fileName           = $el('fileName');
const removeFile         = $el('removeFile');
const previewBtn         = $el('previewBtn');
const errorMsg           = $el('errorMsg');
const errorText          = $el('errorText');
const previewSection     = $el('previewSection');
const previewTable       = $el('previewTable');
const tableInfo          = $el('tableInfo');
const zoomInBtn          = $el('zoomIn');
const zoomOutBtn         = $el('zoomOut');
const zoomLevelEl        = $el('zoomLevel');
const extractBtn         = $el('extractBtn');
const appLayout          = $el('appLayout');
const sidebar            = $el('sidebar');
const sidebarClose       = $el('sidebarClose');
const sidebarOpen        = $el('sidebarOpen');
const teacherSearch      = $el('teacherSearch');
const teacherList        = $el('teacherList');
const contentPlaceholder = $el('contentPlaceholder');
const timetableView      = $el('timetableView');
const teacherNameDisplay = $el('teacherNameDisplay');
const colorLegend        = $el('colorLegend');
const timetableGrid      = $el('timetableGrid');
const downloadPdf        = $el('downloadPdf');
const themeToggle        = $el('themeToggle');

/* ─── Theme ─────────────────────────────────────────────────────── */
function initTheme() {
  const saved = localStorage.getItem('tt-theme') || 'light';
  document.documentElement.setAttribute('data-theme', saved);
  updateThemeIcon(saved);
}
function updateThemeIcon(theme) {
  const icon = themeToggle.querySelector('.theme-icon');
  if (icon) icon.textContent = theme === 'dark' ? '☾' : '☀︎';
}
themeToggle.addEventListener('click', () => {
  const cur = document.documentElement.getAttribute('data-theme');
  const nxt = cur === 'dark' ? 'light' : 'dark';
  document.documentElement.setAttribute('data-theme', nxt);
  localStorage.setItem('tt-theme', nxt);
  updateThemeIcon(nxt);
});
initTheme();

/* ─── Upload & File Handling ────────────────────────────────────── */
uploadBox.addEventListener('click',     () => fileInput.click());
uploadBox.addEventListener('dragover',  e  => { e.preventDefault(); uploadBox.classList.add('dragover'); });
uploadBox.addEventListener('dragleave', () => uploadBox.classList.remove('dragover'));
uploadBox.addEventListener('drop', e => {
  e.preventDefault(); uploadBox.classList.remove('dragover');
  const f = e.dataTransfer.files[0]; if (f) handleFile(f);
});
fileInput.addEventListener('change', () => { if (fileInput.files[0]) handleFile(fileInput.files[0]); });

removeFile.addEventListener('click', () => {
  if (state.sessionId) {
    fetch(`${API_BASE}/session/${state.sessionId}`, { method: 'DELETE' }).catch(() => {});
  }
  fileInput.value = ''; state.sessionId = null;
  uploadMeta.style.display = previewSection.style.display = appLayout.style.display = 'none';
  hideError();
});

function handleFile(file) {
  hideError();
  const ext = file.name.split('.').pop().toLowerCase();
  if (!['xlsx','xlsm'].includes(ext)) {
    showError('Invalid file type. Please upload a .xlsx or .xlsm file.'); return;
  }
  fileName.textContent = file.name;
  uploadMeta.style.display = 'flex';
  // Store file reference for upload on "Preview" click
  state._pendingFile = file;
}

/* ─── Preview (upload to backend, show raw data) ────────────────── */
previewBtn.addEventListener('click', async () => {
  if (!state._pendingFile) { showError('No file selected.'); return; }
  hideError();

  previewBtn.textContent = '⏳ Uploading…'; previewBtn.disabled = true;

  try {
    const formData = new FormData();
    formData.append('file', state._pendingFile);

    const res  = await fetch(`${API_BASE}/upload`, { method: 'POST', body: formData });
    const data = await res.json();
    if (!res.ok) throw new Error(data.detail || 'Upload failed');

    state.sessionId = data.session_id;
    tableInfo.textContent = `Sheet: "${data.filename}" · ${data.rows} rows × ${data.cols} cols · ${data.teachers.length} teachers found`;

    // Build a simple preview table from the metadata
    buildPreviewFromMeta(data);
    previewSection.style.display = 'block';
    previewSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
  } catch (err) {
    showError(`Upload failed: ${err.message}`);
    console.error(err);
  }

  previewBtn.textContent = 'Preview Table →'; previewBtn.disabled = false;
});

function buildPreviewFromMeta(data) {
  previewTable.innerHTML = '';
  const thead = document.createElement('thead');
  const tbody = document.createElement('tbody');

  // Header row
  const hr = document.createElement('tr');
  ['Property', 'Value'].forEach(h => {
    const th = document.createElement('th'); th.textContent = h; hr.appendChild(th);
  });
  thead.appendChild(hr);

  const rows = [
    ['Filename',          data.filename],
    ['Sheets',            data.sheets.join(', ')],
    ['Total rows',        data.rows],
    ['Total columns',     data.cols],
    ['Teachers detected', data.teachers.join(', ') || 'None'],
    ['Divisions',         data.divisions.join(', ') || 'None detected'],
  ];
  rows.forEach(([k, v]) => {
    const tr = document.createElement('tr');
    [k, v].forEach(cell => {
      const td = document.createElement('td'); td.textContent = cell; tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  previewTable.appendChild(thead);
  previewTable.appendChild(tbody);
}

/* ─── Zoom (preview table) ──────────────────────────────────────── */
zoomInBtn.addEventListener('click',  () => changeZoom( 0.1));
zoomOutBtn.addEventListener('click', () => changeZoom(-0.1));
function changeZoom(delta) {
  state.zoomLevel = Math.min(2, Math.max(0.5, state.zoomLevel + delta));
  previewTable.style.transform = `scale(${state.zoomLevel})`;
  zoomLevelEl.textContent = Math.round(state.zoomLevel * 100) + '%';
}

/* ─── Extract (fetch teachers from API) ─────────────────────────── */
extractBtn.addEventListener('click', async () => {
  if (!state.sessionId) { showError('Please upload a file first.'); return; }

  const btnText    = extractBtn.querySelector('.btn-text');
  const btnSpinner = extractBtn.querySelector('.btn-spinner');
  btnText.style.display = 'none'; btnSpinner.style.display = 'flex'; extractBtn.disabled = true;

  try {
    const res   = await fetch(`${API_BASE}/teachers/${state.sessionId}`);
    const data  = await res.json();
    if (!res.ok) throw new Error(data.detail || 'Could not fetch teachers');

    state.teachers = data.teachers;  // [{code, display_name}]
    buildTeacherList();

    appLayout.style.display = 'flex';
    contentPlaceholder.style.display = 'flex';
    timetableView.style.display = 'none';
    appLayout.scrollIntoView({ behavior: 'smooth' });

    if (!state.teachers.length) showError('No teacher codes found in this timetable.');

    if (!document.getElementById('modeTabs')) { injectBatchUI(); injectBatchCSS(); }
  } catch (err) {
    showError(`Extraction failed: ${err.message}`);
    console.error(err);
  }

  btnText.style.display = 'inline'; btnSpinner.style.display = 'none'; extractBtn.disabled = false;
});

/* ─── Teacher Sidebar ───────────────────────────────────────────── */
function buildTeacherList(filter = '') {
  teacherList.innerHTML = '';
  const q = filter.toLowerCase();
  const filtered = state.teachers
    .filter(t => {
      const code = t.code.toLowerCase();
      const name = (t.display_name || '').toLowerCase();
      return code.includes(q) || name.includes(q);
    })
    .sort((a, b) => (a.display_name || a.code).localeCompare(b.display_name || b.code));

  if (!filtered.length) {
    const li = document.createElement('li');
    li.textContent = 'No teachers found'; li.style.opacity = '.5';
    teacherList.appendChild(li); return;
  }
  filtered.forEach(t => {
    const li   = document.createElement('li');
    const name = t.display_name || t.code;
    li.innerHTML = `<span class="teacher-full-name">${name !== t.code ? `${name} (${t.code})` : t.code}</span>`;
    if (t.code === state.selectedTeacher) li.classList.add('active');
    li.addEventListener('click', () => {
      document.querySelectorAll('.teacher-list li').forEach(el => el.classList.remove('active'));
      li.classList.add('active');
      selectTeacher(t.code, name);
      if (window.innerWidth < 900) sidebar.classList.remove('open');
    });
    teacherList.appendChild(li);
  });
}
teacherSearch.addEventListener('input', () => buildTeacherList(teacherSearch.value));

/* ─── Teacher Selection & API call ─────────────────────────────── */
async function selectTeacher(code, displayName) {
  clearTimetable();
  state.selectedTeacher = code;
  state.subjectColorMap = {}; state.colorIndex = 0;
  teacherNameDisplay.textContent = `${displayName || facultyName(code)} (${code})`;
  contentPlaceholder.style.display = 'none';
  timetableView.style.display = 'block';

  showGridLoading();

  try {
    const res  = await fetch(`${API_BASE}/timetable/teacher/${state.sessionId}/${code}`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.detail || 'Failed to load timetable');
    renderTimetableGrid(data.schedule, data.slot_config || SLOT_CONFIG, data.fixed_days || FIXED_DAYS);
  } catch (err) {
    showError(`Could not load timetable: ${err.message}`);
    console.error(err);
    timetableGrid.innerHTML = '';
  }
}

/* ─── Batch Selection & API call ────────────────────────────────── */
async function selectBatch(div, batch) {
  state.selectedDiv   = div;
  state.selectedBatch = batch;
  state.selectedTeacher = null;

  document.querySelectorAll('.batch-btn').forEach(b =>
    b.classList.toggle('active', b.dataset.batch === batch)
  );
  if (window.innerWidth < 900) sidebar.classList.remove('open');

  contentPlaceholder.style.display = 'none';
  timetableView.style.display = 'block';
  teacherNameDisplay.textContent = `Division ${div} · Batch ${batch}  —  ${DIVISION_CONFIG[div]?.name || ''}`;
  state.subjectColorMap = {}; state.colorIndex = 0;
  colorLegend.innerHTML = '';
  showGridLoading();

  try {
    const res  = await fetch(`${API_BASE}/timetable/batch/${state.sessionId}/${div}/${batch}`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.detail || 'Failed to load batch timetable');
    renderTimetableGrid(data.schedule, data.slot_config || SLOT_CONFIG, data.fixed_days || FIXED_DAYS);
  } catch (err) {
    showError(`Could not load batch timetable: ${err.message}`);
    console.error(err);
    timetableGrid.innerHTML = '';
  }
}

/* ─── Render timetable grid from API response ────────────────────
   data.schedule = { MON: { slots: { "8:30": [...entries], ... } }, ... }
   entry = { subject, type, faculty_codes, class_div, span? }
   ─────────────────────────────────────────────────────────────── */
function renderTimetableGrid(schedule, slotConfig, fixedDays) {
  timetableGrid.innerHTML = '';
  colorLegend.innerHTML   = '';

  const subjectsFound = new Set();

  // ── CSS grid columns
  const colDefs = slotConfig.map(s =>
    s.type === 'slot' ? 'minmax(110px,1fr)' : s.type === 'break' ? '48px' : '56px'
  );
  timetableGrid.style.gridTemplateColumns = `120px ${colDefs.join(' ')}`;

  // ── Corner
  const corner = document.createElement('div');
  corner.className = 'grid-cell header-cell corner-cell';
  corner.innerHTML = `<span>Day&thinsp;/&thinsp;Time</span>`;
  timetableGrid.appendChild(corner);

  // ── Time headers
  slotConfig.forEach(slot => {
    const cell = document.createElement('div');
    if (slot.type === 'slot') {
      cell.className = 'grid-cell header-cell time-header-cell';
      cell.innerHTML = `<span class="time-main">${slot.label}</span><span class="time-label">Slot</span>`;
    } else {
      cell.className = `grid-cell break-header-cell ${slot.type}-cell-hdr`;
      cell.innerHTML = `<span class="break-label">${slot.label}</span>`;
    }
    timetableGrid.appendChild(cell);
  });

  // ── Day rows
  let hasAnyData = false;

  fixedDays.forEach(day => {
    const dayCell = document.createElement('div');
    dayCell.className = 'grid-cell day-cell';
    dayCell.innerHTML = `<span class="day-short">${day}</span><span class="day-full">${DAYS_FULL_MAP[day]||day}</span>`;
    timetableGrid.appendChild(dayCell);

    const daySlots = schedule[day]?.slots || {};

    // Track which slot keys to skip (absorbed into a lab span)
    const skipKeys = new Set();

    // Pre-compute spans: find entries with span > 1 from API
    // (Backend already handles merge_info; skipped keys returned omitted from slots.)
    // We need to re-compute CSS span for slots that have span > 1 on entries.

    slotConfig.forEach(slot => {
      if (slot.type !== 'slot') {
        // Break / Lunch separator
        const sep = document.createElement('div');
        sep.className = `grid-cell break-sep-cell ${slot.type}-sep`;
        sep.innerHTML = `<span class="break-sep-label">${slot.label}</span>`;
        timetableGrid.appendChild(sep);
        return;
      }

      if (skipKeys.has(slot.key)) return;

      const entries = daySlots[slot.key] || [];

      // Detect lab span (backend sets span on entries)
      const span    = entries.length > 0 ? (entries[0].span || 1) : 1;
      const cssSpan = computeCssSpan(slot.key, span, slotConfig);

      // Mark subsequent slots as skip
      if (span > 1) {
        let remaining = span - 1, idx = LECTURE_SLOTS.indexOf(slot.key) + 1;
        while (remaining > 0 && idx < LECTURE_SLOTS.length) {
          skipKeys.add(LECTURE_SLOTS[idx]);
          remaining--; idx++;
        }
      }

      if (!entries.length) {
        const empty = document.createElement('div');
        empty.className = 'grid-cell empty-cell';
        if (cssSpan > 1) empty.style.gridColumn = `span ${cssSpan}`;
        empty.innerHTML = `<div class="free-slot-inner"><div class="free-dot"></div><span class="free-text">Free</span></div>`;
        timetableGrid.appendChild(empty);
      } else {
        hasAnyData = true;
        const isLab    = entries.some(e => e.type === 'lab');
        const colorIdx = getSubjectColor(entries[0].subject || entries[0].raw || '');
        const cell     = document.createElement('div');
        cell.className = `grid-cell subject-cell color-${colorIdx}`;
        if (cssSpan > 1) cell.style.gridColumn = `span ${cssSpan}`;

        const bySubject = groupBySubject(entries);
        const typeBadge = isLab
          ? `<span class="type-badge lab-badge">Lab${span > 1 ? ` (${span}h)` : ''}</span>`
          : `<span class="type-badge lec-badge">Lecture</span>`;

        const subjectBlocks = bySubject.map(({ subject, classDivs, facultyCodes }) => {
          subjectsFound.add(subject);
          const divText = classDivs.filter(Boolean).join(', ');
          const facHtml = facultyCodes?.length
            ? `<div class="class-div batch-faculty">${facultyCodes.map(fc =>
                `<span class="fac-chip" title="${facultyName(fc)}">${fc}</span>`).join('')}</div>`
            : '';
          return `<div class="subject-entry">
            <div class="subject-name">${escHtml(subject)}</div>
            ${divText ? `<div class="class-div">${escHtml(divText)}</div>` : ''}
            ${facHtml}
          </div>`;
        }).join('<div class="entry-sep"></div>');

        cell.innerHTML = `<div class="subject-cell-inner">${subjectBlocks}<div class="cell-footer">${typeBadge}</div></div>`;
        timetableGrid.appendChild(cell);
      }
    });
  });

  buildLegend([...subjectsFound]);

  if (!hasAnyData) {
    const notice = document.createElement('div');
    notice.style.cssText = 'grid-column:1/-1;padding:14px;text-align:center;color:var(--text-muted);font-size:.83rem;border-bottom:1px solid var(--border)';
    notice.textContent = '⚠ No lectures found — all slots shown as Free.';
    timetableGrid.prepend(notice);
  }
}

/* ─── Loading skeleton ───────────────────────────────────────────── */
function showGridLoading() {
  timetableGrid.innerHTML = '';
  const colDefs = SLOT_CONFIG.map(s =>
    s.type === 'slot' ? 'minmax(110px,1fr)' : s.type === 'break' ? '48px' : '56px'
  );
  timetableGrid.style.gridTemplateColumns = `120px ${colDefs.join(' ')}`;

  for (let r = 0; r < 7; r++) {
    for (let c = 0; c < SLOT_CONFIG.length + 1; c++) {
      const sk = document.createElement('div');
      sk.className = 'grid-cell';
      sk.style.cssText = 'background:var(--bg2);border-radius:10px;animation:pulse 1.2s ease infinite;';
      timetableGrid.appendChild(sk);
    }
  }
}

/* ─── Helpers ────────────────────────────────────────────────────── */
function computeCssSpan(startKey, lectureSpan, slotConfig) {
  if (lectureSpan <= 1) return 1;
  const sc = slotConfig || SLOT_CONFIG;
  const startIdx = sc.findIndex(s => s.key === startKey);
  if (startIdx === -1) return lectureSpan;
  let left = lectureSpan, cssSpan = 0, i = startIdx;
  while (i < sc.length && left > 0) {
    cssSpan++;
    if (sc[i].type === 'slot') left--;
    i++;
  }
  return cssSpan;
}

function groupBySubject(entries) {
  const map = new Map();
  entries.forEach(e => {
    const k = e.subject || '—';
    if (!map.has(k)) map.set(k, { subject: k, classDivs: [], facultyCodes: [] });
    if (e.class_div) map.get(k).classDivs.push(e.class_div);
    if (e.faculty_codes?.length) {
      e.faculty_codes.forEach(fc => map.get(k).facultyCodes.push(fc));
    }
  });
  map.forEach(v => {
    v.classDivs    = [...new Set(v.classDivs)];
    v.facultyCodes = [...new Set(v.facultyCodes)];
  });
  return [...map.values()];
}

const COLOR_COUNT = 8;
function getSubjectColor(subject) {
  const key = String(subject).replace(/\s+/g,'').toUpperCase().slice(0, 8);
  if (state.subjectColorMap[key] === undefined) {
    state.subjectColorMap[key] = state.colorIndex % COLOR_COUNT;
    state.colorIndex++;
  }
  return state.subjectColorMap[key];
}

function buildLegend(subjects) {
  colorLegend.innerHTML = '';
  [...new Set(subjects)].slice(0, 12).forEach(subj => {
    const idx  = getSubjectColor(subj);
    const item = document.createElement('div');
    item.className = `legend-item color-${idx}`;
    item.textContent = subj;
    colorLegend.appendChild(item);
  });
}

function clearTimetable() {
  timetableGrid.innerHTML = '';
  colorLegend.innerHTML   = '';
  teacherNameDisplay.textContent = '';
}

function escHtml(s) {
  return String(s)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;')
    .replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

/* ─── PDF Download (unchanged from v2) ───────────────────────────── */
downloadPdf.addEventListener('click', async () => {
  const isBatch = state.viewMode === 'batch';
  if (!isBatch && !state.selectedTeacher) return;
  if (isBatch && (!state.selectedDiv || !state.selectedBatch)) return;

  const { jsPDF } = window.jspdf;
  const wrap = $el('timetableGridWrap');

  downloadPdf.textContent = '⏳ Generating…'; downloadPdf.disabled = true;

  const savedOverflow = wrap.style.overflow;
  const savedWidth    = wrap.style.width;
  const savedMaxWidth = wrap.style.maxWidth;
  wrap.style.overflow = 'visible';
  wrap.style.maxWidth = 'none';
  wrap.style.width    = (timetableGrid.scrollWidth + 24) + 'px';
  await delay(120);

  try {
    const bgRaw = getComputedStyle(document.documentElement).getPropertyValue('--surface').trim();
    const bg    = bgRaw || '#ffffff';

    const canvas = await html2canvas(wrap, {
      scale: 2, useCORS: true, backgroundColor: bg,
      width:  wrap.scrollWidth,
      height: wrap.scrollHeight,
      windowWidth: wrap.scrollWidth + 300,
    });

    const pdf  = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
    const pdfW = pdf.internal.pageSize.getWidth();
    const pdfH = pdf.internal.pageSize.getHeight();
    const ML=8, MR=8, MT=28, MB=8;
    const availW = pdfW - ML - MR;
    const availH = pdfH - MT - MB;

    const teacher  = state.teachers.find(t => t.code === state.selectedTeacher);
    const pdfTitle = isBatch
      ? `Division ${state.selectedDiv} · Batch ${state.selectedBatch} — ${DIVISION_CONFIG[state.selectedDiv]?.name}`
      : `Timetable: ${teacher?.display_name || facultyName(state.selectedTeacher)} (${state.selectedTeacher})`;
    const pdfFile = isBatch
      ? `Timetable_${state.selectedDiv}_${state.selectedBatch}.pdf`
      : `Timetable_${(teacher?.display_name || state.selectedTeacher).replace(/\s+/g,'_')}.pdf`;

    const addHeader = (doc, page) => {
      doc.setFontSize(14); doc.setFont('helvetica','bold'); doc.setTextColor(42,31,20);
      doc.text(pdfTitle, ML, 14);
      doc.setFontSize(8); doc.setFont('helvetica','normal'); doc.setTextColor(120,100,80);
      doc.text(`Generated by Timetable Extractor · ${new Date().toLocaleDateString()}` +
               (page > 1 ? `  [page ${page}]` : ''), ML, 21);
    };

    const scaleH  = availH / canvas.height;
    const scaledW = canvas.width * scaleH;

    if (scaledW <= availW) {
      addHeader(pdf, 1);
      pdf.addImage(canvas.toDataURL('image/png'), 'PNG', ML, MT, scaledW, availH);
    } else {
      const pxPerPage = Math.floor((availW / scaledW) * canvas.width);
      let offsetPx = 0, page = 1;
      while (offsetPx < canvas.width) {
        if (page > 1) pdf.addPage();
        addHeader(pdf, page);
        const sliceW = Math.min(pxPerPage, canvas.width - offsetPx);
        const sliceC = document.createElement('canvas');
        sliceC.width = sliceW; sliceC.height = canvas.height;
        sliceC.getContext('2d').drawImage(canvas, offsetPx, 0, sliceW, canvas.height, 0, 0, sliceW, canvas.height);
        const sliceMmW = (sliceW / canvas.width) * scaledW;
        pdf.addImage(sliceC.toDataURL('image/png'), 'PNG', ML, MT, sliceMmW, availH);
        offsetPx += sliceW; page++;
      }
    }
    pdf.save(pdfFile);
  } catch (err) {
    console.error('PDF error:', err);
    alert('PDF generation failed.\n' + err.message);
  }

  wrap.style.overflow = savedOverflow; wrap.style.width = savedWidth; wrap.style.maxWidth = savedMaxWidth;
  downloadPdf.textContent = '⬇ Download PDF'; downloadPdf.disabled = false;
});

/* ─── Sidebar (mobile) ───────────────────────────────────────────── */
sidebarOpen.addEventListener('click',  () => sidebar.classList.add('open'));
sidebarClose.addEventListener('click', () => sidebar.classList.remove('open'));

/* ─── Misc ───────────────────────────────────────────────────────── */
function showError(msg) {
  errorText.textContent = msg;
  errorMsg.style.display = 'flex';
  errorMsg.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}
function hideError() { errorMsg.style.display = 'none'; }
const delay = ms => new Promise(r => setTimeout(r, ms));


/* ══════════════════════════════════════════════════════════════════
   BATCH / DIVISION MODE  —  Sidebar UI injection
   ══════════════════════════════════════════════════════════════════ */

function injectBatchUI() {
  const strip = document.createElement('div');
  strip.id = 'modeTabs';
  strip.innerHTML = `
    <button id="tabTeacher" class="mode-tab active">👤 By Teacher</button>
    <button id="tabBatch"   class="mode-tab">🏫 By Batch</button>
  `;
  teacherSearch.parentElement.insertBefore(strip, teacherSearch);

  const panel = document.createElement('div');
  panel.id = 'batchModePanel';
  panel.style.display = 'none';
  panel.innerHTML = `
    <div id="divPanel">
      <p class="sidebar-section-label">SELECT DIVISION</p>
      <div id="divList"></div>
    </div>
    <div id="batchPanel" style="display:none">
      <p class="sidebar-section-label">SELECT BATCH</p>
      <div id="batchList"></div>
    </div>
  `;
  teacherList.parentElement.appendChild(panel);

  document.getElementById('tabTeacher').addEventListener('click', () => switchMode('teacher'));
  document.getElementById('tabBatch').addEventListener('click',   () => switchMode('batch'));

  const divList = document.getElementById('divList');
  Object.entries(DIVISION_CONFIG).forEach(([div, cfg]) => {
    const btn = document.createElement('button');
    btn.className   = 'div-btn';
    btn.dataset.div = div;
    btn.innerHTML   = `<strong>${div}</strong><span>${cfg.name}</span>`;
    btn.addEventListener('click', () => selectDivision(div));
    divList.appendChild(btn);
  });
}

function switchMode(mode) {
  state.viewMode = mode;
  const tabT = document.getElementById('tabTeacher');
  const tabB = document.getElementById('tabBatch');
  const bmp  = document.getElementById('batchModePanel');
  if (mode === 'teacher') {
    tabT.classList.add('active');   tabB.classList.remove('active');
    bmp.style.display = 'none';
    teacherSearch.style.display = '';
    teacherList.style.display   = '';
  } else {
    tabB.classList.add('active');   tabT.classList.remove('active');
    bmp.style.display = 'block';
    teacherSearch.style.display = 'none';
    teacherList.style.display   = 'none';
  }
}

function selectDivision(div) {
  state.selectedDiv   = div;
  state.selectedBatch = null;

  document.querySelectorAll('.div-btn').forEach(b =>
    b.classList.toggle('active', b.dataset.div === div)
  );

  const batchPanel = document.getElementById('batchPanel');
  const batchList  = document.getElementById('batchList');
  batchPanel.style.display = 'block';
  batchList.innerHTML = '';

  DIVISION_CONFIG[div].batches.forEach(batch => {
    const btn = document.createElement('button');
    btn.className     = 'batch-btn';
    btn.dataset.batch = batch;
    btn.textContent   = batch;
    btn.addEventListener('click', () => selectBatch(div, batch));
    batchList.appendChild(btn);
  });
}

function injectBatchCSS() {
  const style = document.createElement('style');
  style.textContent = `
    /* Skeleton pulse */
    @keyframes pulse {
      0%,100%{opacity:1} 50%{opacity:.4}
    }
    /* ─── Mode tabs ─── */
    #modeTabs {
      display: flex; gap: 6px;
      padding: 10px 12px 4px;
      border-bottom: 1px solid var(--border);
    }
    .mode-tab {
      flex: 1; padding: 7px 6px; border-radius: 8px;
      border: 1px solid var(--border, #e0d6c8);
      background: transparent;
      color: var(--text-muted, #888);
      font-size: 0.77rem; font-weight: 600; cursor: pointer;
      transition: background .18s, color .18s;
    }
    .mode-tab.active {
      background: var(--accent, #e8622a);
      color: #fff; border-color: transparent;
      box-shadow: 0 2px 8px rgba(232,98,42,.35);
    }
    #batchModePanel { padding: 0 12px 16px; overflow-y: auto; }
    .sidebar-section-label {
      font-size: 0.67rem; font-weight: 700; letter-spacing: .09em;
      color: var(--text-muted, #aaa); margin: 14px 0 6px;
      text-transform: uppercase;
    }
    #divList { display: flex; flex-direction: column; gap: 5px; }
    .div-btn {
      display: flex; align-items: center; gap: 10px;
      padding: 9px 12px; border-radius: 9px;
      border: 1px solid var(--border, #e0d6c8);
      background: transparent; cursor: pointer; text-align: left;
      color: var(--text, #333); font-size: 0.82rem;
      transition: background .15s, color .15s, border-color .15s;
      width: 100%;
    }
    .div-btn strong { font-size: 1rem; min-width: 22px; color: var(--accent, #e8622a); }
    .div-btn span   { color: var(--text-muted, #999); font-size: 0.73rem; }
    .div-btn:hover  { background: var(--accent-soft); border-color: var(--accent); }
    .div-btn.active {
      background: var(--accent); color: #fff;
      border-color: transparent;
      box-shadow: 0 2px 8px rgba(232,98,42,.3);
    }
    .div-btn.active strong { color: #fff; }
    .div-btn.active span   { color: rgba(255,255,255,.75); }
    #batchList { display: flex; flex-wrap: wrap; gap: 7px; margin-top: 4px; }
    .batch-btn {
      padding: 6px 16px; border-radius: 20px;
      border: 1px solid var(--border, #e0d6c8);
      background: transparent; cursor: pointer;
      color: var(--text, #444); font-size: 0.84rem; font-weight: 700;
      transition: background .15s, color .15s, border-color .15s;
      letter-spacing: .02em;
    }
    .batch-btn:hover  { background: var(--accent-soft); border-color: var(--accent); }
    .batch-btn.active {
      background: var(--accent); color: #fff;
      border-color: transparent;
      box-shadow: 0 2px 6px rgba(232,98,42,.3);
    }
    .batch-faculty { display: flex; flex-wrap: wrap; gap: 3px; margin-top: 3px; }
    .fac-chip {
      display: inline-block; padding: 1px 6px;
      background: rgba(0,0,0,.08); border-radius: 4px;
      font-size: 0.65rem; font-weight: 700; letter-spacing: .04em;
      color: var(--text-muted, #666); cursor: help;
    }
    [data-theme="dark"] .fac-chip {
      background: rgba(255,255,255,.12); color: rgba(255,255,255,.7);
    }
  `;
  document.head.appendChild(style);
}
