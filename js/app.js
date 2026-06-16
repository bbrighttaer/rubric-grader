/* ── Rubric Grader — app.js ── */

// ── Theme ──
(function () {
  const saved = localStorage.getItem('rg-theme');
  const prefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
  if (saved === 'dark' || (!saved && prefersDark)) {
    document.documentElement.setAttribute('data-theme', 'dark');
  }
})();

document.addEventListener('DOMContentLoaded', () => {
  const btn   = document.getElementById('theme-toggle');
  const icon  = document.getElementById('theme-icon');
  const label = document.getElementById('theme-label');

  function applyTheme(dark) {
    document.documentElement.setAttribute('data-theme', dark ? 'dark' : 'light');
    icon.textContent  = dark ? '☀️' : '🌙';
    label.textContent = dark ? 'Light' : 'Dark';
    localStorage.setItem('rg-theme', dark ? 'dark' : 'light');
  }

  // Sync button label on load
  applyTheme(document.documentElement.getAttribute('data-theme') === 'dark');

  btn.addEventListener('click', () => {
    applyTheme(document.documentElement.getAttribute('data-theme') !== 'dark');
  });
});

// State
let rubric = null;   // { name, scales:[{title,value}], criteria:[{title,desc,weight,descriptors:[]}] }
let selections = {}; // criterionIndex → { scaleIndex, comment }

// ── DOM refs ──
const fileInput      = document.getElementById('file-input');
const dropZone       = document.getElementById('drop-zone');
const fileNameEl     = document.getElementById('file-name');
const studentSection = document.getElementById('student-section');
const gradingSection = document.getElementById('grading-section');
const resultSection  = document.getElementById('result-section');
const criteriaList   = document.getElementById('criteria-list');
const rubricInfo     = document.getElementById('rubric-info');
const finalScoreEl   = document.getElementById('final-score');
const scoreLabelEl   = document.getElementById('score-label');
const scoreGradeEl   = document.getElementById('score-grade-label');
const scorePercentEl = document.getElementById('score-percent');
const feedbackOutput = document.getElementById('feedback-output');
const copyBtn        = document.getElementById('copy-btn');
const resetBtn       = document.getElementById('reset-btn');
const reloadBtn      = document.getElementById('reload-btn');

// ── File upload ──
fileInput.addEventListener('change', e => handleFile(e.target.files[0]));

dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragleave', ()  => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  const file = e.dataTransfer.files[0];
  if (file) handleFile(file);
});

function handleFile(file) {
  if (!file) return;
  if (!file.name.match(/\.(xlsx?|rbc)$/i)) {
    alert('Please select an .xls, .xlsx, or .rbc file.');
    return;
  }
  fileNameEl.textContent = file.name;
  const reader = new FileReader();
  if (file.name.match(/\.rbc$/i)) {
    reader.onload = e => parseRbcRubric(e.target.result, file.name);
    reader.readAsText(file);
  } else {
    reader.onload = e => parseRubric(e.target.result, file.name);
    reader.readAsArrayBuffer(file);
  }
}

// ── Parse Turnitin .rbc rubric (JSON format) ──
function parseRbcRubric(text, fileName) {
  let data;
  try { data = JSON.parse(text); }
  catch (e) { alert('Could not parse .rbc file — invalid JSON.'); return; }

  const rawScales    = data.RubricScale       || [];
  const rawCriteria  = data.RubricCriterion   || [];
  const rawCritScales= data.RubricCriterionScale || [];

  if (!rawScales.length || !rawCriteria.length) {
    alert('No rubric data found in this .rbc file.');
    return;
  }

  // Build scale lookup: id → { title, value }
  const scaleById = {};
  rawScales.sort((a, b) => a.position - b.position)
    .forEach(s => { scaleById[s.id] = { title: s.name, value: s.value }; });

  // Build criterion-scale descriptor lookup: criterionId+scaleValueId → description
  const descMap = {};
  rawCritScales.forEach(cs => {
    descMap[`${cs.criterion}_${cs.scale_value}`] = cs.description || '';
  });

  // Build scales array in position order
  const scales = rawScales
    .sort((a, b) => a.position - b.position)
    .map(s => ({ title: s.name, value: s.value }));

  // Build criteria in position order
  const criteria = rawCriteria
    .sort((a, b) => a.position - b.position)
    .map(crit => {
      const descriptors = (crit.criterion_scales || []).map(scaleId => {
        const scale = scaleById[scaleId];
        if (!scale) return '';
        return descMap[`${crit.id}_${scaleId}`] || '';
      });
      return {
        title:      crit.name        || '',
        desc:       crit.description || '',
        weight:     crit.value       || null,
        weightStr:  crit.value != null ? crit.value + '%' : '',
        descriptors,
      };
    });

  const maxScore = scales[0].value !== null ? scales[0].value : 100;
  rubric = { name: fileName.replace(/\.rbc$/i, ''), scales, criteria, maxScore };
  selections = {};

  renderGrading();
  collapseUploadSection();
  studentSection.classList.remove('hidden');
  gradingSection.classList.remove('hidden');
  resultSection.classList.add('hidden');
  window.scrollTo({ top: studentSection.offsetTop - 20, behavior: 'smooth' });
}

// ── Parse Turnitin XLS rubric ──
// Format:
//   Row 0:  blank | Scale titles…
//   Row 1:  blank | Scale values (numeric)…
//   Row 2+: "Title\n\nDesc\n\nWeight%" | descriptor per scale…
function parseRubric(arrayBuffer, fileName) {
  const wb = XLSX.read(arrayBuffer, { type: 'arraybuffer' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  if (rows.length < 3) {
    alert('This file does not look like a valid Turnitin rubric.');
    return;
  }

  // Scale titles (row 0, cols 1+)
  const scaleTitleRow = rows[0];
  // Scale values (row 1, cols 1+) — could be numbers or strings
  const scaleValueRow = rows[1];

  // Determine how many scale columns exist
  let numScales = 0;
  for (let c = 1; c < scaleTitleRow.length; c++) {
    if (String(scaleTitleRow[c]).trim() !== '') numScales = c;
  }
  numScales += 1; // convert last index to count — but re-derive cleanly:
  numScales = 0;
  for (let c = 1; c < scaleTitleRow.length; c++) {
    if (String(scaleTitleRow[c]).trim() !== '') numScales++;
    else break;
  }

  const scales = [];
  for (let c = 1; c <= numScales; c++) {
    const title = String(scaleTitleRow[c] || '').trim();
    const rawVal = scaleValueRow[c];
    const value = rawVal !== '' && rawVal !== undefined ? parseFloat(rawVal) : null;
    if (title) scales.push({ title, value });
  }

  if (scales.length === 0) {
    alert('Could not find any scale columns in this rubric file.');
    return;
  }

  // Criteria rows (row 2+)
  const criteria = [];
  for (let r = 2; r < rows.length; r++) {
    const row = rows[r];
    const cell0 = String(row[0] || '').trim();
    if (!cell0) continue;

    // Parse criterion cell: "Title\n\nDescription\n\nWeight%"
    const parts = cell0.split(/\n\n+/);
    const title  = parts[0].trim();
    const desc   = parts.length > 2 ? parts.slice(1, -1).join('\n\n').trim() : (parts[1] || '').trim();
    const weightStr = parts.length > 1 ? parts[parts.length - 1].trim() : '';
    const weightMatch = weightStr.match(/([\d.]+)%/);
    const weight = weightMatch ? parseFloat(weightMatch[1]) : null;

    const descriptors = [];
    for (let c = 1; c <= scales.length; c++) {
      descriptors.push(String(row[c] || '').trim());
    }

    if (title) criteria.push({ title, desc, weight, weightStr, descriptors });
  }

  if (criteria.length === 0) {
    alert('No criteria found in this rubric file.');
    return;
  }

  // Determine max possible score (sum of highest scale value per criterion)
  // If scale values are the shared 30/22.5/15/7.5/0 style, max = scales[0].value
  // We'll use scales[0].value as the overall max for display
  const maxScore = scales[0].value !== null ? scales[0].value : 100;

  rubric = { name: fileName.replace(/\.xlsx?$/i, ''), scales, criteria, maxScore };
  selections = {};

  renderGrading();
  collapseUploadSection();
  studentSection.classList.remove('hidden');
  gradingSection.classList.remove('hidden');
  resultSection.classList.add('hidden');
  window.scrollTo({ top: studentSection.offsetTop - 20, behavior: 'smooth' });
}

// ── Render grading UI ──
function renderGrading() {
  // Rubric meta
  rubricInfo.innerHTML =
    `<span>Scales: <strong>${rubric.scales.map(s => s.title).join(' → ')}</strong></span>` +
    `<span>Criteria: <strong>${rubric.criteria.length}</strong></span>` +
    `<span>Max score: <strong>${rubric.maxScore}</strong></span>`;

  criteriaList.innerHTML = '';

  rubric.criteria.forEach((crit, ci) => {
    const card = document.createElement('div');
    card.className = 'criterion-card';
    card.id = `crit-${ci}`;

    // Header
    const header = document.createElement('div');
    header.className = 'criterion-header';
    header.innerHTML =
      `<div>
        <div class="criterion-title">${escHtml(crit.title)}</div>
        ${crit.desc ? `<div class="criterion-desc">${escHtml(crit.desc)}</div>` : ''}
      </div>
      ${crit.weight !== null ? `<div class="criterion-weight">${crit.weight.toFixed(2)}%</div>` : ''}`;
    card.appendChild(header);

    // Scale radio options
    const optionsWrap = document.createElement('div');
    optionsWrap.className = 'scale-options';

    rubric.scales.forEach((scale, si) => {
      const id = `crit-${ci}-scale-${si}`;
      const opt = document.createElement('div');
      opt.className = 'scale-option';

      const descriptor = crit.descriptors[si] || '';
      const valueDisplay = scale.value !== null ? `${scale.value}` : '';

      opt.innerHTML =
        `<input type="radio" name="crit-${ci}" id="${id}" value="${si}" />
        <label for="${id}">
          <span class="scale-name">${escHtml(scale.title)}</span>
          <span class="scale-value">${valueDisplay ? valueDisplay + ' pts' : ''}</span>
          <span class="descriptor-text">${descriptor ? escHtml(descriptor) : '<em style="color:#9ca3af">No descriptor provided.</em>'}</span>
        </label>`;

      opt.querySelector('input').addEventListener('change', () => {
        if (!selections[ci]) selections[ci] = { scaleIndex: si, adjustedVal: scale.value, comment: '' };
        else { selections[ci].scaleIndex = si; selections[ci].adjustedVal = scale.value; }
        showAdjuster(ci, si);
        updateProgress();
      });

      optionsWrap.appendChild(opt);
    });

    // Adjuster row — shown after a level is selected
    const adjuster = document.createElement('div');
    adjuster.className = 'score-adjuster hidden';
    adjuster.id = `adjuster-${ci}`;
    adjuster.innerHTML =
      `<label class="adjuster-label">Adjust score:</label>
      <input type="range" class="adjuster-range" id="range-${ci}" />
      <input type="number" class="adjuster-number" id="number-${ci}" step="0.01" />
      <span class="adjuster-bounds" id="bounds-${ci}"></span>`;

    adjuster.querySelector(`#range-${ci}`).addEventListener('input', e => {
      const val = parseFloat(e.target.value);
      adjuster.querySelector(`#number-${ci}`).value = val;
      selections[ci].adjustedVal = val;
    });
    adjuster.querySelector(`#number-${ci}`).addEventListener('input', e => {
      let val = parseFloat(e.target.value);
      const range = adjuster.querySelector(`#range-${ci}`);
      val = Math.min(Math.max(val, parseFloat(range.min)), parseFloat(range.max));
      range.value = val;
      selections[ci].adjustedVal = val;
    });

    optionsWrap.appendChild(adjuster);
    card.appendChild(optionsWrap);

    // Comment box
    const commentArea = document.createElement('div');
    commentArea.className = 'comment-area';
    commentArea.innerHTML =
      `<label for="comment-${ci}">Marker comment (optional)</label>
      <textarea id="comment-${ci}" placeholder="Add specific feedback for this criterion…" rows="2"></textarea>`;
    commentArea.querySelector('textarea').addEventListener('input', e => {
      if (!selections[ci]) selections[ci] = { scaleIndex: null, comment: e.target.value };
      else selections[ci].comment = e.target.value;
    });
    card.appendChild(commentArea);

    criteriaList.appendChild(card);
  });

  // Progress bar
  const progressWrap = document.createElement('div');
  progressWrap.className = 'grading-progress';
  progressWrap.innerHTML =
    `<span id="progress-text">0 of ${rubric.criteria.length} graded</span>
    <div class="progress-bar-wrap"><div class="progress-bar-fill" id="progress-fill" style="width:0%"></div></div>`;
  criteriaList.prepend(progressWrap);

  // Submit button
  const submitWrap = document.createElement('div');
  submitWrap.className = 'grade-submit-wrap';
  const submitBtn = document.createElement('button');
  submitBtn.className = 'btn btn-grade';
  submitBtn.id = 'submit-grade';
  submitBtn.textContent = 'Calculate Final Grade';
  submitBtn.disabled = true;
  submitBtn.addEventListener('click', calculateResult);
  submitWrap.appendChild(submitBtn);
  criteriaList.appendChild(submitWrap);
}

function showAdjuster(ci, si) {
  const adjuster = document.getElementById(`adjuster-${ci}`);
  const range    = document.getElementById(`range-${ci}`);
  const number   = document.getElementById(`number-${ci}`);
  const bounds   = document.getElementById(`bounds-${ci}`);

  const scales    = rubric.scales;
  const upperVal  = scales[si].value !== null ? scales[si].value : 0;
  // Lower bound: value of the next (lower) scale level, or 0 if last
  const lowerVal  = si + 1 < scales.length && scales[si + 1].value !== null
    ? scales[si + 1].value : 0;

  const defaultVal = upperVal;
  const step = upperVal === lowerVal ? 0.01 : Math.round(((upperVal - lowerVal) / 20) * 100) / 100 || 0.01;

  range.min   = lowerVal;
  range.max   = upperVal;
  range.step  = step;
  range.value = defaultVal;

  number.min   = lowerVal;
  number.max   = upperVal;
  number.value = defaultVal;

  bounds.textContent = `(${lowerVal} – ${upperVal})`;

  selections[ci].adjustedVal = defaultVal;
  adjuster.classList.remove('hidden');
}

function updateProgress() {
  const graded = rubric.criteria.filter((_, ci) =>
    selections[ci] && selections[ci].scaleIndex !== null && selections[ci].scaleIndex !== undefined
  ).length;
  const total = rubric.criteria.length;
  const pct = Math.round((graded / total) * 100);

  document.getElementById('progress-text').textContent = `${graded} of ${total} graded`;
  document.getElementById('progress-fill').style.width = pct + '%';
  document.getElementById('submit-grade').disabled = graded < total;
}

// ── Calculate and display result ──
function calculateResult() {
  // User-supplied total marks overrides the rubric's max scale value.
  const totalMarksInput = parseFloat(document.getElementById('total-marks').value);
  const maxScaleVal = rubric.scales[0].value !== null ? rubric.scales[0].value : 100;
  const maxScore = !isNaN(totalMarksInput) && totalMarksInput > 0 ? totalMarksInput : maxScaleVal;

  // Each criterion's selected scale value is a ratio of the max scale value (e.g. 75/100 = 0.75).
  // When weights are present: score = sum( weight_i/100 * scaleVal_i/maxScaleVal * maxScore )
  // When weights are absent:  score = (average ratio) * maxScore
  let score = 0;
  const hasWeights = rubric.criteria.every(c => c.weight !== null);

  rubric.criteria.forEach((crit, ci) => {
    const sel      = selections[ci];
    const scaleVal = sel.adjustedVal !== undefined && sel.adjustedVal !== null
      ? sel.adjustedVal
      : (rubric.scales[sel.scaleIndex].value !== null ? rubric.scales[sel.scaleIndex].value : 0);
    const ratio    = scaleVal / maxScaleVal;

    if (hasWeights) {
      score += (crit.weight / 100) * ratio * maxScore;
    } else {
      score += ratio * maxScore / rubric.criteria.length;
    }
  });

  score = Math.round(score * 100) / 100;
  const pct = maxScore > 0 ? Math.round((score / maxScore) * 100) : 0;

  // Grade label
  let gradeLabel, gradeClass;
  if (pct >= 85)      { gradeLabel = 'High Distinction'; gradeClass = 'grade-high'; }
  else if (pct >= 75) { gradeLabel = 'Distinction';      gradeClass = 'grade-high'; }
  else if (pct >= 65) { gradeLabel = 'Credit';           gradeClass = 'grade-mid';  }
  else if (pct >= 50) { gradeLabel = 'Pass';             gradeClass = 'grade-mid';  }
  else                { gradeLabel = 'Fail';             gradeClass = 'grade-low';  }

  // Update score display
  const circle = document.querySelector('.score-circle');
  circle.className = 'score-circle ' + gradeClass;
  finalScoreEl.textContent = Number.isInteger(score) ? score : score.toFixed(2);
  scoreLabelEl.textContent = `/ ${maxScore}`;
  scoreGradeEl.textContent = gradeLabel;
  scorePercentEl.textContent = `${pct}%`;

  // Build feedback
  buildFeedback(score, maxScore, maxScaleVal, pct, gradeLabel);

  resultSection.classList.remove('hidden');
  window.scrollTo({ top: resultSection.offsetTop - 20, behavior: 'smooth' });
}

function buildFeedback(score, maxScore, maxScaleVal, pct, gradeLabel) {
  const hasWeights = rubric.criteria.every(c => c.weight !== null);

  let lines = [];
  lines.push(`ASSESSMENT FEEDBACK`);
  lines.push(`${'─'.repeat(50)}`);
  lines.push('');

  rubric.criteria.forEach((crit, ci) => {
    const sel        = selections[ci];
    const scale      = rubric.scales[sel.scaleIndex];
    const descriptor = crit.descriptors[sel.scaleIndex] || '';
    const comment    = (sel.comment || '').trim();

    const rawVal = sel.adjustedVal !== undefined && sel.adjustedVal !== null
      ? sel.adjustedVal
      : (scale.value !== null ? scale.value : 0);
    const ratio = maxScaleVal > 0 ? rawVal / maxScaleVal : 0;

    let scaledScore;
    if (hasWeights && crit.weight !== null) {
      scaledScore = (crit.weight / 100) * ratio * maxScore;
    } else {
      scaledScore = ratio * maxScore / rubric.criteria.length;
    }
    scaledScore = Math.round(scaledScore * 100) / 100;

    lines.push(`▸ ${crit.title}`);
    lines.push(`  Grade     : ${scale.title} (${scaledScore})`);
    if (descriptor) lines.push(`  Standard  : ${descriptor}`);
    if (comment)    lines.push(`  Comment   : ${comment}`);
    lines.push('');
  });

  lines.push(`${'─'.repeat(50)}`);
  lines.push(`FINAL SCORE : ${score % 1 === 0 ? score : score.toFixed(2)} / ${maxScore}`);

  // Render as styled HTML inside the feedback div
  const html = lines.map(line => {
    if (line.startsWith('ASSESSMENT FEEDBACK') || line.startsWith('FINAL SCORE')) {
      return `<div class="fb-total">${escHtml(line)}</div>`;
    } else if (line.startsWith('─')) {
      return `<div class="fb-divider">${escHtml(line)}</div>`;
    } else if (line.startsWith('▸')) {
      return `<div class="fb-criterion">${escHtml(line)}</div>`;
    } else if (line.includes('Grade     :')) {
      return `<div class="fb-level">${escHtml(line)}</div>`;
    } else if (line.includes('Comment   :')) {
      return `<div class="fb-comment">${escHtml(line)}</div>`;
    } else if (line === '') {
      return `<div>&nbsp;</div>`;
    } else {
      return `<div>${escHtml(line)}</div>`;
    }
  }).join('');

  feedbackOutput.innerHTML = html;

  // Store plain text for copying
  feedbackOutput.dataset.plain = lines.join('\n');
}

// ── Copy feedback ──
copyBtn.addEventListener('click', () => {
  const text = feedbackOutput.dataset.plain || feedbackOutput.innerText;
  navigator.clipboard.writeText(text).then(() => showToast('Feedback copied!')).catch(() => {
    // Fallback for file:// protocol
    const ta = document.createElement('textarea');
    ta.value = text;
    ta.style.position = 'fixed'; ta.style.opacity = '0';
    document.body.appendChild(ta);
    ta.select();
    document.execCommand('copy');
    document.body.removeChild(ta);
    showToast('Feedback copied!');
  });
});

// ── Reset / Reload ──
resetBtn.addEventListener('click', () => {
  selections = {};
  document.getElementById('student-name').value = '';
  document.getElementById('student-id').value = '';
  document.getElementById('assessment-name').value = '';
  resultSection.classList.add('hidden');
  // Re-render grading (clear selections)
  renderGrading();
  window.scrollTo({ top: studentSection.offsetTop - 20, behavior: 'smooth' });
});

reloadBtn.addEventListener('click', () => {
  rubric = null; selections = {};
  fileInput.value = '';
  fileNameEl.textContent = 'No file selected';
  expandUploadSection();
  studentSection.classList.add('hidden');
  gradingSection.classList.add('hidden');
  resultSection.classList.add('hidden');
  criteriaList.innerHTML = '';
  rubricInfo.innerHTML = '';
  window.scrollTo({ top: 0, behavior: 'smooth' });
});

// ── Upload section collapse / expand ──
function collapseUploadSection() {
  const section = document.getElementById('upload-section');
  const name = rubric ? rubric.name : (document.getElementById('file-name') || {}).textContent || '';
  section.innerHTML =
    `<h2>1. Load Rubric</h2>
    <div class="upload-collapsed">
      <span class="upload-collapsed-icon">✓</span>
      <span class="upload-collapsed-name">${escHtml(name)}</span>
      <button class="btn btn-outline upload-change-btn" id="change-rubric-btn">Change rubric</button>
    </div>`;
  document.getElementById('change-rubric-btn').addEventListener('click', () => {
    // Rebuild the original upload UI
    section.innerHTML =
      `<h2>1. Load Rubric</h2>
      <div class="upload-area" id="drop-zone">
        <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path stroke-linecap="round" stroke-linejoin="round" d="M3 16.5v2.25A2.25 2.25 0 005.25 21h13.5A2.25 2.25 0 0021 18.75V16.5m-13.5-9L12 3m0 0l4.5 4.5M12 3v13.5"/></svg>
        <p>Drag &amp; drop your <strong>.xls</strong> or <strong>.rbc</strong> rubric file here, or</p>
        <label class="btn btn-primary" for="file-input">Browse file</label>
        <input type="file" id="file-input" accept=".xls,.xlsx,.rbc" />
        <p class="file-hint" id="file-name">No file selected</p>
      </div>`;
    reattachUploadListeners();
    // Also trigger full reload state
    reloadBtn.click();
  });
}

function expandUploadSection() {
  const section = document.getElementById('upload-section');
  section.innerHTML =
    `<h2>1. Load Rubric</h2>
    <div class="upload-area" id="drop-zone">
      <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path stroke-linecap="round" stroke-linejoin="round" d="M3 16.5v2.25A2.25 2.25 0 005.25 21h13.5A2.25 2.25 0 0021 18.75V16.5m-13.5-9L12 3m0 0l4.5 4.5M12 3v13.5"/></svg>
      <p>Drag &amp; drop your <strong>.xls</strong> rubric file here, or</p>
      <label class="btn btn-primary" for="file-input">Browse file</label>
      <input type="file" id="file-input" accept=".xls,.xlsx" />
      <p class="file-hint" id="file-name">No file selected</p>
    </div>`;
  reattachUploadListeners();
}

function reattachUploadListeners() {
  const fi = document.getElementById('file-input');
  const dz = document.getElementById('drop-zone');
  fi.addEventListener('change', e => handleFile(e.target.files[0]));
  dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('drag-over'); });
  dz.addEventListener('dragleave', () => dz.classList.remove('drag-over'));
  dz.addEventListener('drop', e => {
    e.preventDefault();
    dz.classList.remove('drag-over');
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
  });
}

// ── Toast ──
function showToast(msg) {
  let toast = document.querySelector('.toast');
  if (!toast) {
    toast = document.createElement('div');
    toast.className = 'toast';
    document.body.appendChild(toast);
  }
  toast.textContent = msg;
  toast.classList.add('show');
  setTimeout(() => toast.classList.remove('show'), 2200);
}

// ── Helpers ──
function escHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
