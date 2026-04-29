/* ══════════════════════════════════════════
   TIMETABLE EXTRACTOR — script.js  v4
   • Fixed canonical time slots (8:30–3:30)
   • Break / Lunch separator columns in grid
   • Full-grid PDF (no clipping, multi-page)
   • Lab merging (2-hour span)
   • Multi-entry cells
   • MON–SAT rows always rendered
══════════════════════════════════════════ */

/* ─── FACULTY CODE → NAME MAP ─── */
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

/* ═══════════════════════════════════════════════════════════════
   CANONICAL SCHEDULE STRUCTURE
   Lecture slots are 1 hr each; labs span 2 consecutive slots.
   Break  10:30–10:45  (short break between morning sessions)
   Lunch  12:45–1:30   (lunch break)

   SLOT_CONFIG drives BOTH the column headers AND cell layout.
   type: 'slot'  → actual lecture/lab column  (minmax(110px,1fr))
         'break' → narrow break separator      (48px)
         'lunch' → slightly wider lunch sep    (56px)
═══════════════════════════════════════════════════════════════ */
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

/* Lecture-only keys (used for schedule indexing & lab merge) */
const LECTURE_SLOTS = SLOT_CONFIG.filter(s => s.type === 'slot').map(s => s.key);

/* Fixed day order */
const FIXED_DAYS     = ['MON','TUE','WED','THU','FRI','SAT'];
const DAYS_FULL_MAP  = {
  MON:'Monday', TUE:'Tuesday', WED:'Wednesday',
  THU:'Thursday', FRI:'Friday', SAT:'Saturday', SUN:'Sunday'
};
const DAYS_SHORT      = ['MON','TUE','WED','THU','FRI','SAT','SUN'];
const DAYS_FULL_LIST  = ['MONDAY','TUESDAY','WEDNESDAY','THURSDAY','FRIDAY','SATURDAY','SUNDAY'];

const SKIP_TOKENS = new Set([
  'MON','TUE','WED','THU','FRI','SAT','SUN',
  'THE','AND','FOR','NOT','ARE','WAS','HAS',
  'LAB','LEC','TH','FE','SE','TE','BE',
  'DAY','TIME','AM','PM','NO','ID','PR','TD','LE','A','B','C','D',
  'CLASS','DIVISION','SLOT','BREAK','LUNCH','FREE',
]);

/* ─── DIVISION / BATCH CONFIG ─── */
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

/* ─── STATE ─── */
const state = {
  workbook: null, sheetData: [], headers: [],
  teachers: [], selectedTeacher: null,
  subjectColorMap: {}, colorIndex: 0, zoomLevel: 1,
  dayCol: 0, timeCol: 1,
  classColMap: {}, dayRowMap: {}, timeSlotLabels: [],
  /* batch mode */
  viewMode: 'teacher',   // 'teacher' | 'batch'
  selectedDiv: null,
  selectedBatch: null,
  divColMap: {},         // div letter → Set of column indices
};

/* ─── DOM ─── */
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

/* ══════════════════════════════════════════
   THEME
══════════════════════════════════════════ */
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

/* ══════════════════════════════════════════
   FILE UPLOAD / DRAG-DROP
══════════════════════════════════════════ */
uploadBox.addEventListener('click',     () => fileInput.click());
uploadBox.addEventListener('dragover',  e  => { e.preventDefault(); uploadBox.classList.add('dragover'); });
uploadBox.addEventListener('dragleave', () => uploadBox.classList.remove('dragover'));
uploadBox.addEventListener('drop', e => {
  e.preventDefault(); uploadBox.classList.remove('dragover');
  const f = e.dataTransfer.files[0]; if (f) handleFile(f);
});
fileInput.addEventListener('change', () => { if (fileInput.files[0]) handleFile(fileInput.files[0]); });
removeFile.addEventListener('click', () => {
  fileInput.value = ''; state.workbook = null; state.sheetData = [];
  uploadMeta.style.display = previewSection.style.display = appLayout.style.display = 'none';
  hideError();
});

function handleFile(file) {
  hideError();
  const ext = file.name.split('.').pop().toLowerCase();
  if (!['xlsx','xls'].includes(ext)) {
    showError('Invalid file type. Please upload a .xlsx or .xls file.'); return;
  }
  fileName.textContent = file.name; uploadMeta.style.display = 'flex';
  const reader = new FileReader();
  reader.onload = e => {
    try { state.workbook = XLSX.read(e.target.result, { type: 'array' }); }
    catch { showError('Could not read the Excel file.'); }
  };
  reader.readAsArrayBuffer(file);
}

/* ══════════════════════════════════════════
   PREVIEW
══════════════════════════════════════════ */
previewBtn.addEventListener('click', () => {
  if (!state.workbook) { showError('File not loaded yet.'); return; }
  buildPreview();
  previewSection.style.display = 'block';
  previewSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
});

function buildPreview() {
  const sheetName = state.workbook.SheetNames[0];
  const ws   = state.workbook.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  state.sheetData = data; state.headers = data[0] || [];
  tableInfo.textContent = `Sheet: "${sheetName}" · ${data.length} rows × ${state.headers.length} cols`;

  previewTable.innerHTML = '';
  const thead = document.createElement('thead');
  const tbody = document.createElement('tbody');
  const hr    = document.createElement('tr');
  state.headers.forEach(h => { const th = document.createElement('th'); th.textContent = h; hr.appendChild(th); });
  thead.appendChild(hr);
  for (let i = 1; i < Math.min(data.length, 60); i++) {
    const tr = document.createElement('tr');
    (data[i] || []).forEach(cell => { const td = document.createElement('td'); td.textContent = cell; tr.appendChild(td); });
    tbody.appendChild(tr);
  }
  previewTable.appendChild(thead); previewTable.appendChild(tbody);
}

zoomInBtn.addEventListener('click',  () => changeZoom( 0.1));
zoomOutBtn.addEventListener('click', () => changeZoom(-0.1));
function changeZoom(delta) {
  state.zoomLevel = Math.min(2, Math.max(0.5, state.zoomLevel + delta));
  previewTable.style.transform = `scale(${state.zoomLevel})`;
  zoomLevelEl.textContent = Math.round(state.zoomLevel * 100) + '%';
}

/* ══════════════════════════════════════════
   EXTRACT
══════════════════════════════════════════ */
extractBtn.addEventListener('click', async () => {
  if (!state.sheetData.length) return;
  const btnText    = extractBtn.querySelector('.btn-text');
  const btnSpinner = extractBtn.querySelector('.btn-spinner');
  btnText.style.display = 'none'; btnSpinner.style.display = 'flex'; extractBtn.disabled = true;
  await delay(600);
  try {
    parseTimetableStructure(); parseTeachers(); buildTeacherList();
    appLayout.style.display = 'flex'; contentPlaceholder.style.display = 'flex';
    timetableView.style.display = 'none';
    appLayout.scrollIntoView({ behavior: 'smooth' });
    if (!state.teachers.length) showError('No teacher codes found. Check console (F12).');
    /* Inject batch UI tabs after extract */
    if (!document.getElementById('modeTabs')) { injectBatchUI(); injectBatchCSS(); }
  } catch (err) {
    console.error(err); showError('Could not extract timetable. Please verify your file format.');
  }
  btnText.style.display = 'inline'; btnSpinner.style.display = 'none'; extractBtn.disabled = false;
});

/* ══════════════════════════════════════════
   PARSE TIMETABLE STRUCTURE
══════════════════════════════════════════ */
function parseTimetableStructure() {
  const data = state.sheetData;
  const dayCC = {}, timCC = {};
  for (let r = 0; r < data.length; r++) {
    const row = data[r] || [];
    for (let c = 0; c < row.length; c++) {
      const v = String(row[c] || '').trim().toUpperCase();
      if (DAYS_FULL_LIST.some(d => v === d) || DAYS_SHORT.some(d => v === d)) dayCC[c] = (dayCC[c]||0)+1;
      if (/^\d{1,2}[:\.]\d{2}(\s*(AM|PM))?$/i.test(v)) timCC[c] = (timCC[c]||0)+1;
    }
  }
  state.dayCol  = +(Object.keys(dayCC).sort((a,b) => dayCC[b]  - dayCC[a])[0]  ?? 0);
  state.timeCol = +(Object.keys(timCC).sort((a,b) => timCC[b] - timCC[a])[0] ?? 1);

  const timeSeen = new Set(), labels = [];
  for (const row of data) {
    const v = normaliseTime(String((row||[])[state.timeCol]||'').trim());
    if (v && !timeSeen.has(v)) { timeSeen.add(v); labels.push(v); }
  }
  state.timeSlotLabels = labels;

  const classCols = {};
  for (let r = 0; r < Math.min(data.length, 8); r++) {
    const row = data[r]||[];
    const dv  = String(row[state.dayCol]||'').trim().toUpperCase();
    if (['CLASS','DIVISION',''].includes(dv)) {
      for (let c = 0; c < row.length; c++) {
        if (c===state.dayCol||c===state.timeCol) continue;
        const v = String(row[c]||'').trim(); if (v&&!/^\d+$/.test(v)) classCols[c]=v;
      }
    }
  }
  if (!Object.keys(classCols).length) {
    (data[0]||[]).forEach((v,c) => {
      if (c!==state.dayCol&&c!==state.timeCol) { const s=String(v||'').trim(); if(s) classCols[c]=s; }
    });
  }
  state.classColMap = classCols;

  const dayRowMap = {}; let curDay = null;
  for (const row of data) {
    const dv = String((row||[])[state.dayCol]||'').trim().toUpperCase();
    const dk = matchDay(dv);
    if (dk) { curDay = dk; if (!dayRowMap[dk]) dayRowMap[dk]=[]; }
    if (curDay) {
      const t = String((row||[])[state.timeCol]||'').trim();
      if (t && /\d/.test(t)) dayRowMap[curDay].push(row);
    }
  }
  state.dayRowMap = dayRowMap;
  console.log('[Structure]', { dayCol:state.dayCol, timeCol:state.timeCol, timeSlots:state.timeSlotLabels, classCols, days:Object.keys(state.dayRowMap) });

  /* ── Build divColMap: which columns belong to which division ── */
  buildDivColMap();
}

/* ══════════════════════════════════════════
   PARSE TEACHERS
══════════════════════════════════════════ */
function parseTeachers() {
  const data = state.sheetData;
  const known = new Set(Object.keys(FACULTY_MAP).map(k => k.toUpperCase()));
  const set = new Set(), occ = {};
  for (const row of data) {
    for (const cell of (row||[])) {
      const raw = String(cell||'').trim(); if (!raw) continue;
      extractCodes(raw).forEach(code => {
        if (known.has(code)) { set.add(code); }
        else if (/^[A-Z]{2,5}$/.test(code) && !SKIP_TOKENS.has(code)) {
          occ[code]=(occ[code]||0)+1; if(occ[code]>=2) set.add(code);
        }
      });
    }
  }
  state.teachers = Array.from(set);
  console.log('[Teachers]', state.teachers);
}

/* ══════════════════════════════════════════
   TEACHER LIST (SIDEBAR)
══════════════════════════════════════════ */
function buildTeacherList(filter = '') {
  teacherList.innerHTML = '';
  const q = filter.toLowerCase();
  const filtered = state.teachers
    .filter(c => { const n=facultyName(c).toLowerCase(); return c.toLowerCase().includes(q)||n.includes(q); })
    .sort((a,b) => facultyName(a).localeCompare(facultyName(b)));
  if (!filtered.length) {
    const li=document.createElement('li'); li.textContent='No teachers found'; li.style.opacity='.5';
    teacherList.appendChild(li); return;
  }
  filtered.forEach(code => {
    const li=document.createElement('li'), name=facultyName(code);
    li.innerHTML=`<span class="teacher-full-name">${FACULTY_MAP[code]?`${name} (${code})`:code}</span>`;
    if (code===state.selectedTeacher) li.classList.add('active');
    li.addEventListener('click', () => {
      document.querySelectorAll('.teacher-list li').forEach(el=>el.classList.remove('active'));
      li.classList.add('active'); selectTeacher(code);
      if (window.innerWidth<900) sidebar.classList.remove('open');
    });
    teacherList.appendChild(li);
  });
}
teacherSearch.addEventListener('input', () => buildTeacherList(teacherSearch.value));

/* ══════════════════════════════════════════
   SELECT TEACHER
══════════════════════════════════════════ */
function selectTeacher(code) {
  clearTimetable();
  state.selectedTeacher = code; state.subjectColorMap = {}; state.colorIndex = 0;
  teacherNameDisplay.textContent   = `${facultyName(code)} (${code})`;
  contentPlaceholder.style.display = 'none';
  timetableView.style.display      = 'block';
  buildTimetableGrid(code);
}
function clearTimetable() {
  timetableGrid.innerHTML = ''; colorLegend.innerHTML = ''; teacherNameDisplay.textContent = '';
}

const COLOR_COUNT = 8;
function getSubjectColor(subject) {
  const key = String(subject).replace(/\s+/g,'').toUpperCase().slice(0,8);
  if (state.subjectColorMap[key]===undefined) { state.subjectColorMap[key]=state.colorIndex%COLOR_COUNT; state.colorIndex++; }
  return state.subjectColorMap[key];
}

/* ══════════════════════════════════════════
   BUILD TIMETABLE GRID
══════════════════════════════════════════ */
function buildTimetableGrid(teacher) {
  const data = state.sheetData, dayCol = state.dayCol, timeCol = state.timeCol;

  /* ── 1. Initialise empty schedule ── */
  const schedule = {};
  for (const day of FIXED_DAYS) {
    schedule[day] = {};
    for (const key of LECTURE_SLOTS) schedule[day][key] = [];
  }

  const subjectsFound = new Set();
  let curDay = null;

  /* ── 2. Fill schedule from Excel data ── */
  for (let r = 0; r < data.length; r++) {
    const row    = data[r] || [];
    const dayVal = String(row[dayCol]  || '').trim().toUpperCase();
    const timeRaw= String(row[timeCol] || '').trim();

    const dk = matchDay(dayVal);
    if (dk) curDay = dk;
    if (!curDay || !FIXED_DAYS.includes(curDay)) continue;

    const timeNorm = normaliseTime(timeRaw);
    if (!timeNorm) continue;

    const slotKey = resolveSlotKey(timeNorm);
    if (!slotKey) continue;

    for (let c = 0; c < row.length; c++) {
      if (c===dayCol||c===timeCol) continue;
      const raw = String(row[c]||'').trim(); if (!raw) continue;
      splitCellSegments(raw).forEach(seg => {
        if (!containsTeacher(seg, teacher)) return;
        const entry = parseCellSegment(seg, teacher); if (!entry) return;
        entry.classDiv = state.classColMap[c] || '';
        schedule[curDay][slotKey].push(entry);
        if (entry.subject) subjectsFound.add(entry.subject);
      });
    }
  }

  /* ── 3. Lab merging ── */
  const mergeInfo = {};
  for (const day of FIXED_DAYS) {
    mergeInfo[day] = {};
    let i = 0;
    while (i < LECTURE_SLOTS.length) {
      const key = LECTURE_SLOTS[i], entries = schedule[day][key];
      if (!entries.length) { i++; continue; }
      const isLab = entries.some(e => e.type==='lab'), subjKey = entries[0]?.subject||'';
      if (isLab) {
        let span = 1;
        while (i+span < LECTURE_SLOTS.length) {
          const k2 = LECTURE_SLOTS[i+span], ent2 = schedule[day][k2];
          if (!ent2.length) break;
          if (ent2.some(e=>e.type==='lab') && ent2[0]?.subject===subjKey) { mergeInfo[day][k2]={skip:true}; span++; }
          else break;
        }
        mergeInfo[day][key] = { span };
      }
      i++;
    }
  }

  /* ── 4. CSS grid columns ── */
  const colDefs = SLOT_CONFIG.map(s =>
    s.type==='slot' ? 'minmax(110px,1fr)' : s.type==='break' ? '48px' : '56px'
  );
  timetableGrid.style.gridTemplateColumns = `120px ${colDefs.join(' ')}`;
  timetableGrid.innerHTML = '';

  /* ── 5. Header row ── */
  const corner = document.createElement('div');
  corner.className = 'grid-cell header-cell corner-cell';
  corner.innerHTML = `<span>Day&thinsp;/&thinsp;Time</span>`;
  timetableGrid.appendChild(corner);

  for (const slot of SLOT_CONFIG) {
    const cell = document.createElement('div');
    if (slot.type === 'slot') {
      cell.className = 'grid-cell header-cell time-header-cell';
      cell.innerHTML = `<span class="time-main">${slot.label}</span><span class="time-label">Slot</span>`;
    } else {
      cell.className = `grid-cell break-header-cell ${slot.type}-cell-hdr`;
      cell.innerHTML = `<span class="break-label">${slot.label}</span>`;
    }
    timetableGrid.appendChild(cell);
  }

  /* ── 6. Day rows ── */
  let hasAnyData = false;

  for (const day of FIXED_DAYS) {
    const dayCell = document.createElement('div');
    dayCell.className = 'grid-cell day-cell';
    dayCell.innerHTML = `<span class="day-short">${day}</span><span class="day-full">${DAYS_FULL_MAP[day]||day}</span>`;
    timetableGrid.appendChild(dayCell);

    for (const slot of SLOT_CONFIG) {
      /* Break / Lunch separator */
      if (slot.type !== 'slot') {
        const sep = document.createElement('div');
        sep.className = `grid-cell break-sep-cell ${slot.type}-sep`;
        sep.innerHTML = `<span class="break-sep-label">${slot.label}</span>`;
        timetableGrid.appendChild(sep);
        continue;
      }

      const mi      = mergeInfo[day]?.[slot.key];
      if (mi?.skip) continue;   /* absorbed by lab merge */

      const entries = schedule[day][slot.key] || [];
      const span    = mi?.span || 1;
      const cssSpan = computeCssSpan(slot.key, span);

      if (!entries.length) {
        /* FREE cell */
        const empty = document.createElement('div');
        empty.className = 'grid-cell empty-cell';
        if (cssSpan > 1) empty.style.gridColumn = `span ${cssSpan}`;
        empty.innerHTML = `<div class="free-slot-inner"><div class="free-dot"></div><span class="free-text">Free</span></div>`;
        timetableGrid.appendChild(empty);
      } else {
        /* Subject cell */
        hasAnyData = true;
        const isLab    = entries.some(e => e.type==='lab');
        const colorIdx = getSubjectColor(entries[0].subject || entries[0].raw || '');
        const cell     = document.createElement('div');
        cell.className = `grid-cell subject-cell color-${colorIdx}`;
        if (cssSpan > 1) cell.style.gridColumn = `span ${cssSpan}`;

        const bySubject    = groupBySubject(entries);
        const typeBadge    = isLab
          ? `<span class="type-badge lab-badge">Lab${span>1?` (${span}h)`:''}</span>`
          : `<span class="type-badge lec-badge">Lecture</span>`;
        const subjectBlocks = bySubject.map(({ subject, classDivs }) => {
          const divText = classDivs.filter(Boolean).join(', ');
          return `<div class="subject-entry"><div class="subject-name">${escHtml(subject)}</div>${divText?`<div class="class-div">${escHtml(divText)}</div>`:''}</div>`;
        }).join('<div class="entry-sep"></div>');

        cell.innerHTML = `<div class="subject-cell-inner">${subjectBlocks}<div class="cell-footer">${typeBadge}</div></div>`;
        timetableGrid.appendChild(cell);
      }
    }
  }

  buildLegend([...subjectsFound]);

  if (!hasAnyData) {
    const notice = document.createElement('div');
    notice.style.cssText = 'grid-column:1/-1;padding:14px;text-align:center;color:var(--text-muted);font-size:.83rem;border-bottom:1px solid var(--border)';
    notice.textContent = '⚠ No lectures found for this teacher — all slots shown as Free.';
    timetableGrid.prepend(notice);
  }
}

/* ══════════════════════════════════════════
   SLOT KEY RESOLVER
   Maps a normalised Excel time → nearest
   canonical LECTURE_SLOTS key (±20 min tol.)
══════════════════════════════════════════ */
function resolveSlotKey(timeNorm) {
  const mins = parseTimeToMins(timeNorm);
  let best = null, bestDiff = Infinity;
  for (const key of LECTURE_SLOTS) {
    const diff = Math.abs(parseTimeToMins(key) - mins);
    if (diff < bestDiff) { bestDiff = diff; best = key; }
  }
  return bestDiff <= 20 ? best : null;
}

/* ══════════════════════════════════════════
   CSS SPAN CALCULATOR
   A lab spanning N lecture slots also needs
   to span any break/lunch columns in between.
══════════════════════════════════════════ */
function computeCssSpan(startKey, lectureSpan) {
  if (lectureSpan <= 1) return 1;
  const startIdx = SLOT_CONFIG.findIndex(s => s.key === startKey);
  if (startIdx === -1) return lectureSpan;
  let left = lectureSpan, cssSpan = 0, i = startIdx;
  while (i < SLOT_CONFIG.length && left > 0) {
    cssSpan++;
    if (SLOT_CONFIG[i].type === 'slot') left--;
    i++;
  }
  return cssSpan;
}

/* ══════════════════════════════════════════
   CELL PARSING
══════════════════════════════════════════ */
function splitCellSegments(raw) {
  const lines = raw.split(/[\n\r]+/).map(s => s.trim()).filter(Boolean);
  const result = [];
  lines.forEach(line => splitOnCommaOutsideParens(line).forEach(p => { if (p.trim()) result.push(p.trim()); }));
  return result.length ? result : [raw.trim()];
}

function splitOnCommaOutsideParens(str) {
  const parts = []; let depth = 0, start = 0;
  for (let i = 0; i < str.length; i++) {
    if (str[i]==='(') depth++;
    else if (str[i]===')') depth--;
    else if (str[i]===',' && depth===0) { parts.push(str.slice(start,i)); start=i+1; }
  }
  parts.push(str.slice(start)); return parts;
}

function containsTeacher(text, code) {
  const esc = code.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');
  return new RegExp(`(?:^|[^A-Za-z])${esc}(?:[^A-Za-z]|$)`,'i').test(text);
}

function parseCellSegment(segment, teacherCode) {
  let s = segment.trim();
  const isLab = /\blab\b|\bpr\b|\bpractical\b/i.test(s);
  const esc = teacherCode.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');
  s = s.replace(new RegExp(`\\(\\s*${esc}\\s*\\)`,'gi'),'');
  s = s.replace(new RegExp(`(?:^|\\s)${esc}(?:\\s|$)`,'gi'),' ');
  s = s.replace(/\b[A-Z]\d+\b/g,'').replace(/\b\d{2,}\b/g,'')
       .replace(/\b(Lab|Pr|Lec|Practical|TD|LE)\b/gi,'')
       .replace(/[()]/g,'').replace(/[-\/,]+$/,'').replace(/^[-\/,]+/,'')
       .replace(/\s{2,}/g,' ').trim();
  if (!s || s.length<2 || SKIP_TOKENS.has(s.toUpperCase())) return null;
  return { subject:s, raw:segment, type: isLab?'lab':'lec' };
}

function extractCodes(raw) {
  return [...raw.matchAll(/\b([A-Z]{2,5})\b/g)].map(m => m[1]);
}

function groupBySubject(entries) {
  const map = new Map();
  entries.forEach(e => {
    const k = e.subject||'—';
    if (!map.has(k)) map.set(k,{subject:k,classDivs:[]});
    if (e.classDiv) map.get(k).classDivs.push(e.classDiv);
  });
  map.forEach(v => { v.classDivs=[...new Set(v.classDivs)]; });
  return [...map.values()];
}

/* ══════════════════════════════════════════
   HELPERS
══════════════════════════════════════════ */
function normaliseTime(raw) {
  const m = raw.match(/(\d{1,2})[:\.](\d{2})/); if (!m) return '';
  let h = parseInt(m[1]); const min = m[2];
  if (/pm/i.test(raw) && h<12) h+=12;
  if (/am/i.test(raw) && h===12) h=0;
  return `${h}:${min}`;
}
function parseTimeToMins(t) {
  const m = String(t).match(/(\d{1,2}):(\d{2})/); if (!m) return 9999;
  let h=parseInt(m[1]); const min=parseInt(m[2]);
  if (h<7) h+=12; return h*60+min;
}
function matchDay(val) {
  const v = String(val).trim().toUpperCase();
  const si = DAYS_SHORT.findIndex(d => v===d); if (si>=0) return DAYS_SHORT[si];
  const fi = DAYS_FULL_LIST.findIndex(d => v===d||v.startsWith(d)); if (fi>=0) return DAYS_SHORT[fi];
  return null;
}
function escHtml(s) {
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
function buildLegend(subjects) {
  colorLegend.innerHTML = '';
  [...new Set(subjects)].slice(0,12).forEach(subj => {
    const idx = getSubjectColor(subj), item = document.createElement('div');
    item.className = `legend-item color-${idx}`; item.textContent = subj;
    colorLegend.appendChild(item);
  });
}

/* ══════════════════════════════════════════
   DOWNLOAD PDF — FULL GRID, NO CLIPPING
   Steps:
   1. Temporarily remove overflow:auto and
      stretch wrapper to full scroll width.
   2. Capture with html2canvas at scale:2.
   3. Fit to A4-landscape; if canvas width
      exceeds page width → split into pages.
   4. Restore wrapper styles.
══════════════════════════════════════════ */
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

    const pdfTitle = isBatch
      ? `Division ${state.selectedDiv} · Batch ${state.selectedBatch} — ${DIVISION_CONFIG[state.selectedDiv].name}`
      : `Timetable: ${facultyName(state.selectedTeacher)} (${state.selectedTeacher})`;
    const pdfFile = isBatch
      ? `Timetable_${state.selectedDiv}_${state.selectedBatch}.pdf`
      : `Timetable_${facultyName(state.selectedTeacher).replace(/\s+/g,'_')}.pdf`;

    const addHeader = (doc, page) => {
      doc.setFontSize(14); doc.setFont('helvetica','bold'); doc.setTextColor(42,31,20);
      doc.text(pdfTitle, ML, 14);
      doc.setFontSize(8); doc.setFont('helvetica','normal'); doc.setTextColor(120,100,80);
      doc.text(`Generated by Timetable Extractor · ${new Date().toLocaleDateString()}` + (page>1?`  [page ${page}]`:''), ML, 21);
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

/* ══════════════════════════════════════════
   SIDEBAR (MOBILE)
══════════════════════════════════════════ */
sidebarOpen.addEventListener('click',  () => sidebar.classList.add('open'));
sidebarClose.addEventListener('click', () => sidebar.classList.remove('open'));

/* ══════════════════════════════════════════
   MISC
══════════════════════════════════════════ */
function showError(msg) { errorText.textContent=msg; errorMsg.style.display='flex'; errorMsg.scrollIntoView({behavior:'smooth',block:'nearest'}); }
function hideError() { errorMsg.style.display='none'; }
const delay = ms => new Promise(r => setTimeout(r, ms));
/* ══════════════════════════════════════════════════════════════════
   BATCH / DIVISION MODE  —  Complete Implementation
   Adds a "By Batch" tab in the sidebar. After clicking a division
   then a batch, builds a full individual timetable for that batch.
   • Whole-division lectures  → detected via column header mapping
   • Batch-specific labs/lecs → detected via batch code in cell text
   • Exact subject + faculty codes shown as in master timetable
══════════════════════════════════════════════════════════════════ */

/* ── 1. Build divColMap: division letter → Set<colIndex> ──────── */
function buildDivColMap() {
  const data = state.sheetData;
  state.divColMap = {};
  Object.keys(DIVISION_CONFIG).forEach(d => state.divColMap[d] = new Set());

  /* Scan first 5 rows for column headers */
  for (let r = 0; r < Math.min(data.length, 5); r++) {
    const row = data[r] || [];
    row.forEach((cell, c) => {
      if (c === state.dayCol || c === state.timeCol) return;
      const v = String(cell || '').trim().toUpperCase();
      Object.keys(DIVISION_CONFIG).forEach(div => {
        if (
          v === div ||
          v === `FE-${div}` ||
          v === `FE ${div}` ||
          v.startsWith(`FE-${div} `) ||
          v.startsWith(`FE ${div} `) ||
          v.includes(`(${div})`) ||
          v.startsWith(`${div} (`) ||
          v === `FE - ${div}`
        ) {
          state.divColMap[div].add(c);
        }
      });
    });
  }

  /* Fallback: if no headers found, try classColMap values */
  Object.entries(state.classColMap).forEach(([col, label]) => {
    const v = String(label).trim().toUpperCase();
    Object.keys(DIVISION_CONFIG).forEach(div => {
      if (v.includes(div) || v.startsWith(`FE-${div}`) || v === div) {
        state.divColMap[div].add(Number(col));
      }
    });
  });

  console.log('[DivColMap]', Object.fromEntries(
    Object.entries(state.divColMap).map(([k,v]) => [k,[...v]])
  ));
}

/* ── 2. Does a cell belong to this batch? ─────────────────────── */
function cellBelongsToBatch(rawCell, div, batch, colIndex) {
  const raw = rawCell.trim();
  if (!raw) return false;

  /* Case A: batch code explicitly in cell text (e.g. A1, B2) */
  const batchEsc = batch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  if (new RegExp(`(?:^|[^A-Za-z0-9])${batchEsc}(?:[^A-Za-z0-9]|$)`, 'i').test(raw)) return true;

  /* Case B: no specific batch codes in cell → whole-division lecture
     Qualify: column must belong to this div AND no OTHER div's batch code is in cell */
  const hasBatchCode = /\b[A-H]\d\b/i.test(raw);
  if (!hasBatchCode) {
    const colBelongsToDiv = state.divColMap[div]?.has(colIndex);
    if (colBelongsToDiv) return true;

    /* Also accept if cell has NO batch codes AND column is in classColMap
       with a label matching this division */
    const colLabel = String(state.classColMap[colIndex] || '').toUpperCase();
    if (colLabel && (colLabel.includes(div) || colLabel.startsWith('FE-' + div))) return true;
  }

  return false;
}

/* ── 3. Parse one cell segment for a batch ────────────────────── */
function parseBatchCellSegment(seg, div, batch) {
  let s = seg.trim();
  if (!s || s.length < 2) return null;

  const isLab = /\blab\b|\bpr\b|\bpractical\b/i.test(s);

  /* Collect ALL caps tokens */
  const allCodes    = [...s.matchAll(/\b([A-Z]{2,5})\b/g)].map(m => m[1]);
  const batchCodes  = allCodes.filter(c => /^[A-H]\d$/i.test(c));
  const facultyCodes= allCodes.filter(c =>
    !(/^[A-H]\d$/i.test(c)) &&
    !SKIP_TOKENS.has(c) &&
    (FACULTY_MAP[c] || c.length >= 2)
  );

  /* If cell names batches but ours isn't among them → skip */
  if (batchCodes.length > 0 && !batchCodes.map(b=>b.toUpperCase()).includes(batch.toUpperCase())) return null;

  /* Build subject string: remove batch codes + code-only paren groups */
  let subj = s
    .replace(/\b[A-H]\d\b/gi, '')
    .replace(/\(([^)]*)\)/g, (m, inner) => {
      const toks = inner.trim().split(/[\s,·]+/);
      const allCodeTokens = toks.every(t => /^[A-Z]{2,5}$/.test(t) || /^[A-H]\d$/i.test(t) || t === '');
      return allCodeTokens ? '' : m;
    })
    .replace(/[-,;\/·•]+\s*$/, '')
    .replace(/^\s*[-,;\/·•]+/, '')
    .replace(/\s{2,}/g, ' ')
    .trim();

  /* Remove faculty codes from subject string (they'll show separately) */
  facultyCodes.forEach(fc => {
    const fe = fc.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    subj = subj.replace(new RegExp(`(?:^|\\s)${fe}(?:\\s|$)`, 'g'), ' ').trim();
  });

  subj = subj.replace(/[()]/g,'').replace(/\s{2,}/g,' ').trim();
  if (!subj || subj.length < 2 || SKIP_TOKENS.has(subj.toUpperCase())) return null;

  /* Filter faculty codes — remove batch/skip tokens */
  const cleanFaculty = facultyCodes.filter(fc =>
    !SKIP_TOKENS.has(fc) && !/^[A-H]\d$/i.test(fc)
  );

  return {
    subject: subj,
    facultyCodes: cleanFaculty,
    type: isLab ? 'lab' : 'lec',
    raw: s,
  };
}

/* ── 4. Core: Build & render batch timetable ──────────────────── */
function buildBatchTimetable(div, batch) {
  const data    = state.sheetData;
  const dayCol  = state.dayCol;
  const timeCol = state.timeCol;

  /* Init empty schedule */
  const schedule = {};
  FIXED_DAYS.forEach(day => {
    schedule[day] = {};
    LECTURE_SLOTS.forEach(key => { schedule[day][key] = []; });
  });

  let curDay = null;

  for (let r = 0; r < data.length; r++) {
    const row     = data[r] || [];
    const dayVal  = String(row[dayCol]  || '').trim().toUpperCase();
    const timeRaw = String(row[timeCol] || '').trim();

    const dk = matchDay(dayVal);
    if (dk) curDay = dk;
    if (!curDay || !FIXED_DAYS.includes(curDay)) continue;

    const timeNorm = normaliseTime(timeRaw);
    if (!timeNorm) continue;
    const slotKey = resolveSlotKey(timeNorm);
    if (!slotKey) continue;

    for (let c = 0; c < row.length; c++) {
      if (c === dayCol || c === timeCol) continue;
      const raw = String(row[c] || '').trim();
      if (!raw) continue;

      if (!cellBelongsToBatch(raw, div, batch, c)) continue;

      splitCellSegments(raw).forEach(seg => {
        const entry = parseBatchCellSegment(seg, div, batch);
        if (!entry) return;
        /* Dedupe: skip if same subject already added for this slot */
        const exists = schedule[curDay][slotKey].some(e => e.subject === entry.subject && e.type === entry.type);
        if (!exists) schedule[curDay][slotKey].push(entry);
      });
    }
  }

  /* Lab merging — same logic as teacher view */
  const mergeInfo = {};
  FIXED_DAYS.forEach(day => {
    mergeInfo[day] = {};
    let i = 0;
    while (i < LECTURE_SLOTS.length) {
      const key     = LECTURE_SLOTS[i];
      const entries = schedule[day][key];
      if (!entries.length) { i++; continue; }
      const isLab = entries.some(e => e.type === 'lab');
      if (isLab) {
        const subjKey = entries[0]?.subject || '';
        let span = 1;
        while (i + span < LECTURE_SLOTS.length) {
          const k2   = LECTURE_SLOTS[i + span];
          const ent2 = schedule[day][k2];
          if (!ent2.length) break;
          if (ent2.some(e => e.type === 'lab') && ent2[0]?.subject === subjKey) {
            mergeInfo[day][k2] = { skip: true }; span++;
          } else break;
        }
        mergeInfo[day][key] = { span };
      }
      i++;
    }
  });

  /* ── Render grid ── */
  state.subjectColorMap = {}; state.colorIndex = 0;
  const subjectsFound = new Set();

  const colDefs = SLOT_CONFIG.map(s =>
    s.type === 'slot' ? 'minmax(110px,1fr)' : s.type === 'break' ? '48px' : '56px'
  );
  timetableGrid.style.gridTemplateColumns = `120px ${colDefs.join(' ')}`;
  timetableGrid.innerHTML = '';
  colorLegend.innerHTML   = '';

  /* Corner header */
  const corner = document.createElement('div');
  corner.className = 'grid-cell header-cell corner-cell';
  corner.innerHTML = `<span>Day&thinsp;/&thinsp;Time</span>`;
  timetableGrid.appendChild(corner);

  /* Time slot headers */
  SLOT_CONFIG.forEach(slot => {
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

  /* Day rows */
  let hasAnyData = false;
  FIXED_DAYS.forEach(day => {
    /* Day label cell */
    const dayCell = document.createElement('div');
    dayCell.className = 'grid-cell day-cell';
    dayCell.innerHTML = `<span class="day-short">${day}</span><span class="day-full">${DAYS_FULL_MAP[day] || day}</span>`;
    timetableGrid.appendChild(dayCell);

    SLOT_CONFIG.forEach(slot => {
      /* Break/Lunch separator */
      if (slot.type !== 'slot') {
        const sep = document.createElement('div');
        sep.className = `grid-cell break-sep-cell ${slot.type}-sep`;
        sep.innerHTML = `<span class="break-sep-label">${slot.label}</span>`;
        timetableGrid.appendChild(sep);
        return;
      }

      const mi = mergeInfo[day]?.[slot.key];
      if (mi?.skip) return;  // absorbed into lab span

      const entries = schedule[day][slot.key] || [];
      const span    = mi?.span || 1;
      const cssSpan = computeCssSpan(slot.key, span);

      if (!entries.length) {
        /* FREE cell */
        const empty = document.createElement('div');
        empty.className = 'grid-cell empty-cell';
        if (cssSpan > 1) empty.style.gridColumn = `span ${cssSpan}`;
        empty.innerHTML = `<div class="free-slot-inner"><div class="free-dot"></div><span class="free-text">Free</span></div>`;
        timetableGrid.appendChild(empty);
      } else {
        hasAnyData = true;
        const isLab    = entries.some(e => e.type === 'lab');
        const colorIdx = getSubjectColor(entries[0].subject || '');
        const cell     = document.createElement('div');
        cell.className = `grid-cell subject-cell color-${colorIdx}`;
        if (cssSpan > 1) cell.style.gridColumn = `span ${cssSpan}`;

        /* Cell content: subject name + faculty codes (exact as in master) */
        const subjectBlocks = entries.map(e => {
          subjectsFound.add(e.subject);
          const facultyHtml = e.facultyCodes?.length
            ? `<div class="class-div batch-faculty">${
                e.facultyCodes.map(fc =>
                  `<span class="fac-chip" title="${facultyName(fc)}">${fc}</span>`
                ).join('')
              }</div>`
            : '';
          return `<div class="subject-entry">
            <div class="subject-name">${escHtml(e.subject)}</div>
            ${facultyHtml}
          </div>`;
        }).join('<div class="entry-sep"></div>');

        const typeBadge = isLab
          ? `<span class="type-badge lab-badge">Lab${span > 1 ? ` (${span}h)` : ''}</span>`
          : `<span class="type-badge lec-badge">Lecture</span>`;

        cell.innerHTML = `<div class="subject-cell-inner">${subjectBlocks}<div class="cell-footer">${typeBadge}</div></div>`;
        timetableGrid.appendChild(cell);
      }
    });
  });

  buildLegend([...subjectsFound]);

  if (!hasAnyData) {
    const notice = document.createElement('div');
    notice.style.cssText = 'grid-column:1/-1;padding:16px;text-align:center;color:var(--text-muted);font-size:.85rem;border-bottom:1px solid var(--border)';
    notice.innerHTML = `⚠ No entries found for <strong>${batch}</strong>.<br>
      <small>Tip: Make sure batch codes like "<strong>${batch}</strong>" appear inside the cells of your Excel sheet, or that column headers clearly show Division <strong>${div}</strong>.</small>`;
    timetableGrid.prepend(notice);
  }
}

/* ── 5. Inject UI: tabs + division/batch panel into sidebar ───── */
function injectBatchUI() {
  /* Tab strip — insert above teacher search */
  const strip = document.createElement('div');
  strip.id    = 'modeTabs';
  strip.innerHTML = `
    <button id="tabTeacher" class="mode-tab active">👤 By Teacher</button>
    <button id="tabBatch"   class="mode-tab">🏫 By Batch</button>
  `;
  teacherSearch.parentElement.insertBefore(strip, teacherSearch);

  /* Batch panel */
  const panel = document.createElement('div');
  panel.id    = 'batchModePanel';
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

  /* Tab click handlers */
  document.getElementById('tabTeacher').addEventListener('click', () => switchMode('teacher'));
  document.getElementById('tabBatch').addEventListener('click',   () => switchMode('batch'));

  /* Populate division buttons */
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

/* ── 6. Switch between Teacher / Batch mode ───────────────────── */
function switchMode(mode) {
  state.viewMode = mode;
  const tabT = document.getElementById('tabTeacher');
  const tabB = document.getElementById('tabBatch');
  const bmp  = document.getElementById('batchModePanel');

  if (mode === 'teacher') {
    tabT.classList.add('active');   tabB.classList.remove('active');
    bmp.style.display              = 'none';
    teacherSearch.style.display    = '';
    teacherList.style.display      = '';
  } else {
    tabB.classList.add('active');   tabT.classList.remove('active');
    bmp.style.display              = 'block';
    teacherSearch.style.display    = 'none';
    teacherList.style.display      = 'none';
  }
}

/* ── 7. Division selected → show batch buttons ────────────────── */
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

/* ── 8. Batch selected → render timetable ─────────────────────── */
function selectBatch(div, batch) {
  state.selectedDiv   = div;
  state.selectedBatch = batch;
  state.selectedTeacher = null;

  document.querySelectorAll('.batch-btn').forEach(b =>
    b.classList.toggle('active', b.dataset.batch === batch)
  );
  if (window.innerWidth < 900) sidebar.classList.remove('open');

  contentPlaceholder.style.display = 'none';
  timetableView.style.display      = 'block';
  teacherNameDisplay.textContent   =
    `Division ${div} · Batch ${batch}   —   ${DIVISION_CONFIG[div].name}`;

  buildBatchTimetable(div, batch);
}

/* ── 9. Inject CSS for batch UI elements ──────────────────────── */
function injectBatchCSS() {
  const style = document.createElement('style');
  style.textContent = `
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
      background: var(--accent, #7c3aed);
      color: #fff; border-color: transparent;
      box-shadow: 0 2px 8px rgba(124,58,237,.35);
    }

    /* ─── Batch panel ─── */
    #batchModePanel { padding: 0 12px 16px; overflow-y: auto; }
    .sidebar-section-label {
      font-size: 0.67rem; font-weight: 700; letter-spacing: .09em;
      color: var(--text-muted, #aaa); margin: 14px 0 6px;
      text-transform: uppercase;
    }

    /* ─── Division buttons ─── */
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
    .div-btn strong {
      font-size: 1rem; min-width: 22px; color: var(--accent, #7c3aed);
    }
    .div-btn span { color: var(--text-muted, #999); font-size: 0.73rem; }
    .div-btn:hover { background: var(--accent-soft, #ede9fe); border-color: var(--accent, #7c3aed); }
    .div-btn.active {
      background: var(--accent, #7c3aed); color: #fff;
      border-color: transparent;
      box-shadow: 0 2px 8px rgba(124,58,237,.3);
    }
    .div-btn.active strong { color: #fff; }
    .div-btn.active span   { color: rgba(255,255,255,.75); }

    /* ─── Batch buttons ─── */
    #batchList { display: flex; flex-wrap: wrap; gap: 7px; margin-top: 4px; }
    .batch-btn {
      padding: 6px 16px; border-radius: 20px;
      border: 1px solid var(--border, #e0d6c8);
      background: transparent; cursor: pointer;
      color: var(--text, #444); font-size: 0.84rem; font-weight: 700;
      transition: background .15s, color .15s, border-color .15s;
      letter-spacing: .02em;
    }
    .batch-btn:hover { background: var(--accent-soft, #ede9fe); border-color: var(--accent, #7c3aed); }
    .batch-btn.active {
      background: var(--accent, #7c3aed); color: #fff;
      border-color: transparent;
      box-shadow: 0 2px 6px rgba(124,58,237,.3);
    }

    /* ─── Faculty chip in batch cells ─── */
    .batch-faculty { display: flex; flex-wrap: wrap; gap: 3px; margin-top: 3px; }
    .fac-chip {
      display: inline-block; padding: 1px 6px;
      background: rgba(0,0,0,.08); border-radius: 4px;
      font-size: 0.65rem; font-weight: 700; letter-spacing: .04em;
      color: var(--text-muted, #666);
      cursor: help;
    }
    [data-theme="dark"] .fac-chip {
      background: rgba(255,255,255,.12); color: rgba(255,255,255,.7);
    }
  `;
  document.head.appendChild(style);
}
