const TIME_OFF_VALUES = new Set(['TIME OFF', 'REQ VAC']);
const SHUTTLE_COMBO_LABEL = '10:30am - 6:30pm (c)';
const CREW_SUGGESTION_REGEX = /^\s*\d{1,2}:\d{2}(?:am|pm)\s*-\s*\d{1,2}:\d{2}(?:am|pm)\s*$/i;
const SHIFT_TIME_REGEX_GLOBAL = /(\d{1,2})(?::(\d{2}))?\s*(am|pm)/ig;
const CREW_SHIFT_CUTOFF_MINUTES = (17 * 60) + 45;
const CUSTOM_SHIFT_VALUE = '__custom__';
const CANONICAL_SHIFT_LABELS = [
  '5AM–12PM',
  '6AM–12PM',
  '7AM–12PM',
  'AM (6:00AM–2:00PM)',
  'AM (6:15AM–2:15PM)',
  'PM (2:00PM–10:00PM)',
  'PM (2:15PM–10:15PM)',
  'Audit (10:00PM–6:00AM)',
  'Audit (10:15PM–6:15AM)',
  'AM (3:30AM–11:30AM)',
  'Midday (10:30AM–6:30PM)',
  'PM (5:30PM–1:30AM)',
  'Crew (5:45PM–1:45AM)',
  'Crew (8:00PM–12:00AM)',
  'Crew (9:00PM–1:00AM)',
  '8AM–4:30PM',
];

function reqVacLabelForDateKey(dateKey) {
  if (!dateKey) {
    return 'REQ VAC';
  }
  const parts = dateKey.split('-');
  if (parts.length !== 3) {
    return 'REQ VAC';
  }
  const month = parseInt(parts[1], 10);
  const day = parseInt(parts[2], 10);
  if (Number.isNaN(month) || Number.isNaN(day)) {
    return 'REQ VAC';
  }
  return `REQ VAC ${month}/${day}`;
}

function reqOffLabelForDateKey(dateKey) {
  if (!dateKey) {
    return 'REQ OFF';
  }
  const parts = dateKey.split('-');
  if (parts.length !== 3) {
    return 'REQ OFF';
  }
  const month = parseInt(parts[1], 10);
  const day = parseInt(parts[2], 10);
  if (Number.isNaN(month) || Number.isNaN(day)) {
    return 'REQ OFF';
  }
  return `REQ OFF ${month}/${day}`;
}

function applyTimeOffLabelsToSelect(selectEl) {
  if (!selectEl) {
    return;
  }
  const cell = selectEl.closest('.cell');
  const dateKey = cell ? cell.getAttribute('data-date') : '';
  const vacLabel = reqVacLabelForDateKey(dateKey);
  const offLabel = reqOffLabelForDateKey(dateKey);
  Array.from(selectEl.options).forEach(opt => {
    if (opt.value === 'REQ VAC') {
      opt.textContent = vacLabel;
    } else if (opt.value === 'TIME OFF') {
      const trimmed = (opt.textContent || '').trim().toUpperCase();
      if (trimmed.startsWith('REQ VAC')) {
        opt.textContent = vacLabel;
      } else {
        opt.textContent = offLabel;
      }
    }
  });
}

function ensureTimeOffOptions(selectEl) {
  if (![...selectEl.options].some(o => o.value === 'TIME OFF')) {
    const opt = document.createElement('option');
    opt.value = 'TIME OFF';
    opt.textContent = 'REQ OFF';
    selectEl.insertBefore(opt, selectEl.firstChild);
  }
  if (![...selectEl.options].some(o => o.value === 'REQ VAC')) {
    const opt = document.createElement('option');
    opt.value = 'REQ VAC';
    opt.textContent = 'REQ VAC';
    selectEl.insertBefore(opt, selectEl.firstChild);
  }
  applyTimeOffLabelsToSelect(selectEl);
}

function showToast(msg) {
  const el = document.getElementById('toast');
  if (!el) return;
  el.textContent = msg;
  el.classList.add('show');
  setTimeout(() => el.classList.remove('show'), 1600);
}

function applyThemePreference(isDark) {
  const root = document.documentElement;
  if (isDark) {
    root.classList.add('dark-mode');
  } else {
    root.classList.remove('dark-mode');
  }
  if (document.body) {
    document.body.classList.toggle('dark-mode', isDark);
  }
  const toggle = document.getElementById('theme-toggle');
  if (toggle) {
    toggle.setAttribute('aria-pressed', isDark ? 'true' : 'false');
  }
}

function initThemeToggle() {
  const toggle = document.getElementById('theme-toggle');
  if (!toggle) return;

  let storedTheme = null;
  try {
    storedTheme = localStorage.getItem('theme');
  } catch (err) {
    storedTheme = null;
  }

  let isDark;
  if (storedTheme === 'dark') {
    isDark = true;
  } else if (storedTheme === 'light') {
    isDark = false;
  } else {
    isDark = document.documentElement.classList.contains('dark-mode');
  }
  applyThemePreference(isDark);

  toggle.addEventListener('click', () => {
    const next = !document.documentElement.classList.contains('dark-mode');
    applyThemePreference(next);
    try {
      localStorage.setItem('theme', next ? 'dark' : 'light');
    } catch (err) {
      /* localStorage may be unavailable; ignore */
    }
  });
}

function isSuggestedCrewValue(value) {
  return typeof value === 'string' && CREW_SUGGESTION_REGEX.test(value);
}

function shiftTimePoints(value) {
  if (typeof value !== 'string') return [];
  const normalized = value.replace(/\u2013/g, '-');
  const regex = new RegExp(SHIFT_TIME_REGEX_GLOBAL.source, 'ig');
  const points = [];
  let match;
  while ((match = regex.exec(normalized))) {
    let hour = parseInt(match[1], 10);
    const minute = match[2] ? parseInt(match[2], 10) : 0;
    const period = (match[3] || '').toLowerCase();
    if (Number.isNaN(hour) || Number.isNaN(minute)) continue;
    hour = hour % 12;
    if (period === 'pm') hour += 12;
    points.push((hour * 60) + minute);
  }
  return points;
}

function shiftWindowMinutes(value) {
  const points = shiftTimePoints(value);
  if (points.length < 2) return null;
  let start = points[0];
  let end = points[1];
  if (end <= start) end += 24 * 60;
  return [start, end];
}

function windowOverlapMinutes(aStart, aEnd, bStart, bEnd) {
  return Math.max(0, Math.min(aEnd, bEnd) - Math.max(aStart, bStart));
}

function shiftStartMinutes(value) {
  const points = shiftTimePoints(value);
  return points.length ? points[0] : null;
}

function basicSelectClass(value) {
  if (!value) return '';
  if (value === 'Set') return 'select-gray';
  if (TIME_OFF_VALUES.has(value)) return 'select-yellow';
  if (value === SHUTTLE_COMBO_LABEL) return 'select-orange';
  if (isSuggestedCrewValue(value)) return 'select-red';
  if (value === '5AM–12PM') return 'select-gold';
  if (value === '6AM–12PM') return 'select-blue';
  if (value === '7AM–12PM') return 'select-purple';
  if (value.startsWith('Audit')) return 'select-red';
  if (value.startsWith('Crew')) return 'select-red';
  if (value.startsWith('Midday')) return 'select-blue';
  if (value.startsWith('PM (2')) return 'select-blue';
  if (value.startsWith('PM')) return 'select-purple';
  if (value.startsWith('AM')) return 'select-green';
  if (value === '8AM–4:30PM') return 'select-green';
  const startMinutes = shiftStartMinutes(value);
  if (startMinutes !== null && startMinutes >= CREW_SHIFT_CUTOFF_MINUTES) return 'select-red';
  return '';
}

function buildShiftClassWindows() {
  return CANONICAL_SHIFT_LABELS.map(label => {
    const window = shiftWindowMinutes(label);
    if (!window) return null;
    const className = basicSelectClass(label);
    if (!className) return null;
    return { start: window[0], end: window[1], className };
  }).filter(Boolean);
}

const SHIFT_CLASS_WINDOWS = buildShiftClassWindows();

function matchShiftClassByOverlap(value) {
  const window = shiftWindowMinutes(value);
  if (!window) return '';
  const [start, end] = window;
  let bestClass = '';
  let bestOverlap = 0;
  let bestStart = Number.POSITIVE_INFINITY;
  SHIFT_CLASS_WINDOWS.forEach(ref => {
    const overlap = windowOverlapMinutes(start, end, ref.start, ref.end);
    if (overlap >= 300 && (overlap > bestOverlap || (overlap === bestOverlap && ref.start < bestStart))) {
      bestOverlap = overlap;
      bestStart = ref.start;
      bestClass = ref.className;
    }
  });
  return bestClass;
}

function selectClassForValue(_section, value) {
  const cls = basicSelectClass(value);
  if (cls) return cls;
  return matchShiftClassByOverlap(value);
}

function updateSelectClass(selectEl, section, value) {
  selectEl.classList.remove('select-gold', 'select-green', 'select-blue', 'select-red', 'select-gray', 'select-yellow', 'select-purple', 'select-orange');
  const cls = selectClassForValue(section, value);
  if (cls) selectEl.classList.add(cls);
}

function initSelectColors() {
  document.querySelectorAll('.shift-select').forEach(sel => {
    const cell = sel.closest('.cell');
    const section = cell ? cell.getAttribute('data-section') : '';
    updateSelectClass(sel, section, sel.value);
    applyTimeOffLabelsToSelect(sel);
  });
}


function updateCoverageUI(data) {
  const {
    counts,
    missing,
    required,
    variant_counts,
    shuttle_missing,
    bb_missing,
    bb_order_warnings,
    maintenance_missing,
    fd_duplicates,
  } = data;

  // Update Front Desk headers using FD missing map
  document
    .querySelectorAll('.schedule-table[aria-label="Front Desk"] .row.header .cell.head[data-date]')
    .forEach(h => {
      const dk = h.getAttribute('data-date');
      if (!dk) return;
      const countSpan = document.querySelector(`.count[data-count-for="${dk}"]`);
      if (countSpan) countSpan.textContent = String(counts?.[dk] || 0);
      h.classList.remove('missing', 'duplicate');
      if (missing?.[dk]) {
        h.classList.add('missing');
      } else if (fd_duplicates?.[dk]) {
        h.classList.add('duplicate');
      }
    });

  // Update Shuttle headers using Shuttle missing map
  document
    .querySelectorAll('.schedule-table[aria-label="Shuttle"] .row.header .cell.head[data-date]')
    .forEach(h => {
      const dk = h.getAttribute('data-date');
      if (!dk) return;
      if (shuttle_missing?.[dk]) h.classList.add('missing'); else h.classList.remove('missing');
    });

  // Update Breakfast Bar headers using Breakfast missing map
  document
    .querySelectorAll('.schedule-table[aria-label="Breakfast Bar"] .row.header .cell.head[data-date]')
    .forEach(h => {
      const dk = h.getAttribute('data-date');
      if (!dk) return;
      h.classList.remove('missing', 'order-warning');
      if (bb_missing?.[dk]) {
        h.classList.add('missing');
      } else if (bb_order_warnings?.[dk]) {
        h.classList.add('order-warning');
      }
    });
  const breakfastOrderWarning = document.getElementById('breakfast-order-warning');
  if (breakfastOrderWarning) {
    const showBreakfastOrderWarning = Object.values(bb_order_warnings || {}).some(Boolean);
    breakfastOrderWarning.hidden = !showBreakfastOrderWarning;
  }

  // Update Maintenance headers using Maintenance missing map
  document
    .querySelectorAll('.schedule-table[aria-label="Maintenance"] .row.header .cell.head[data-date]')
    .forEach(h => {
      const dk = h.getAttribute('data-date');
      if (!dk) return;
      if (maintenance_missing?.[dk]) h.classList.add('missing'); else h.classList.remove('missing');
    });

  // Missing list with detailed Front Desk variant information
  const list = document.getElementById('missing-list');
  if (list) {
    const missingDates = Object.entries(missing || {}).filter(([, v]) => v).map(([k]) => k);
    if (missingDates.length === 0) {
      list.innerHTML = '';
    } else {
      const fmt = (iso) => {
        const d = new Date(iso);
        return d.toLocaleDateString(undefined, { weekday: 'short', month: 'short', day: 'numeric' });
      };

      const missingDetails = missingDates.map(dateKey => {
        const dateStr = fmt(dateKey);
        const variants = (variant_counts && variant_counts[dateKey]) || {};
        const missingVariants = [];
        if ((variants.AM || 0) < 2) missingVariants.push(`AM (${variants.AM || 0}/2)`);
        if ((variants.PM || 0) < 2) missingVariants.push(`PM (${variants.PM || 0}/2)`);
        if ((variants.Audit || 0) < 2) missingVariants.push(`Audit (${variants.Audit || 0}/2)`);
        return `${dateStr}: ${missingVariants.join(', ')}`;
      });

      list.innerHTML = `Missing coverage: ${missingDetails.join('; ')}`;
    }
  }
}

function updateConflictsUI() {
  const cells = Array.from(document.querySelectorAll('.schedule-table .row:not(.header) .cell'));
  cells.forEach(function(c){ c.classList.remove('conflict'); });
  const groups = {};
  cells.forEach(function(cell){
    const empId = cell.getAttribute('data-employee-id');
    const empName = cell.getAttribute('data-employee');
    const empKey = empId || empName;
    const dk = cell.getAttribute('data-date');
    if (!empKey || !dk) return;
    const sel = cell.querySelector('select');
    const val = sel ? sel.value : null;
    const active = !!val && val !== 'Set' && !TIME_OFF_VALUES.has(val);
    const key = empKey + '||' + dk;
    if (!groups[key]) groups[key] = [];
    groups[key].push({ cell: cell, active: active });
  });
  Object.keys(groups).forEach(function(key){
    const arr = groups[key];
    let activeCount = 0;
    arr.forEach(function(it){ if (it.active) activeCount++; });
    if (activeCount > 1) {
      arr.forEach(function(it){ if (it.active) it.cell.classList.add('conflict'); });
    }
  });
}

function applySuggestionOptionState(sel, mode = 'selected') {
  const options = sel ? sel.querySelectorAll('option[data-suggestion-label]') : null;
  if (!options || !options.length) return;
  options.forEach(opt => {
    const label = opt.getAttribute('data-suggestion-label');
    const plain = opt.getAttribute('data-suggestion-text') || label;
    if (mode === 'menu') {
      if (label) opt.textContent = label;
    } else {
      if (opt.selected) {
        if (plain) opt.textContent = plain;
      } else if (label) {
        opt.textContent = label;
      }
    }
  });
}

function minutesFromTimeStr(value) {
  if (!value) return null;
  const parts = value.split(':');
  if (parts.length < 2) return null;
  const hour = Number(parts[0]);
  const minute = Number(parts[1]);
  if (Number.isNaN(hour) || Number.isNaN(minute)) return null;
  return ((hour % 24) * 60) + minute;
}

function formatSuggestionClock(minutes) {
  const total = ((minutes % (24 * 60)) + (24 * 60)) % (24 * 60);
  const hour = Math.floor(total / 60);
  const minute = total % 60;
  const suffix = hour < 12 ? 'am' : 'pm';
  const displayHour = (hour % 12) || 12;
  const body = `${displayHour}:${String(minute).padStart(2, '0')}`;
  return `${body}${suffix}`;
}

function formatSuggestionRange(startMinutes, endMinutes) {
  const start = formatSuggestionClock(startMinutes);
  const end = formatSuggestionClock(endMinutes);
  return `${start} - ${end}`;
}

function computeShuttleSuggestion(dateKey) {
  if (!dateKey) return '';
  const cells = document.querySelectorAll(`.aircrew-table .cell[data-date="${dateKey}"]`);
  const minutes = [];
  cells.forEach(cell => {
    let list = [];
    try {
      list = JSON.parse(cell.dataset.times || '[]');
    } catch (e) {
      list = [];
    }
    list.forEach(value => {
      const mins = minutesFromTimeStr(value);
      if (mins === null) return;
      minutes.push(mins);
    });
  });
  if (!minutes.length) return '';
  const ordered = minutes.slice().sort((a, b) => a - b);
  const extended = ordered.concat(ordered[0] + 24 * 60);
  let maxGap = -1;
  let gapIdx = 0;
  for (let i = 0; i < ordered.length; i += 1) {
    const gap = extended[i + 1] - extended[i];
    if (gap > maxGap) {
      maxGap = gap;
      gapIdx = i;
    }
  }
  let startTotal = extended[gapIdx + 1];
  let endTotal = extended[gapIdx];
  if (endTotal < startTotal) {
    endTotal += 24 * 60;
  }
  const buffer = 60;
  startTotal = Math.max(startTotal - buffer, 0);
  endTotal = Math.min(endTotal + buffer, startTotal + (18 * 60));
  return formatSuggestionRange(startTotal, endTotal);
}

function ensureShuttleSuggestionOption(sel, suggestion) {
  if (!sel) return;
  let opt = sel.querySelector('option[data-suggestion="1"]');
  if (!suggestion) {
    if (opt) {
      const wasSelected = opt.selected;
      opt.remove();
      if (wasSelected) {
        sel.selectedIndex = 0;
      }
    }
    applySuggestionOptionState(sel, document.activeElement === sel ? 'menu' : 'selected');
    return;
  }
  if (!opt) {
    opt = document.createElement('option');
    opt.setAttribute('data-suggestion', '1');
    sel.appendChild(opt);
  } else {
    sel.appendChild(opt); // ensure last
  }
  opt.value = suggestion;
  opt.setAttribute('data-suggestion-text', suggestion);
  opt.setAttribute('data-suggestion-label', `${suggestion} (Suggested)`);
  opt.textContent = document.activeElement === sel ? `${suggestion} (Suggested)` : suggestion;
  applySuggestionOptionState(sel, document.activeElement === sel ? 'menu' : 'selected');
}

function updateShuttleSuggestionForDate(dateKey) {
  if (!dateKey) return;
  const suggestion = computeShuttleSuggestion(dateKey);
  const shuttleCells = document.querySelectorAll(`.schedule-table[aria-label="Shuttle"] .cell[data-date="${dateKey}"]`);
  shuttleCells.forEach(cell => {
    const sel = cell.querySelector('select.shift-select');
    ensureShuttleSuggestionOption(sel, suggestion);
    if (suggestion) {
      cell.setAttribute('data-suggestion', suggestion);
    } else {
      cell.removeAttribute('data-suggestion');
    }
  });
}

function initShuttleSuggestions() {
  const dateKeys = new Set();
  document.querySelectorAll('.schedule-table[aria-label="Shuttle"] .cell[data-date]').forEach(cell => {
    const dk = cell.getAttribute('data-date');
    if (dk) dateKeys.add(dk);
    const sel = cell.querySelector('select.shift-select');
    const suggestion = cell.getAttribute('data-suggestion');
    if (suggestion) {
      ensureShuttleSuggestionOption(sel, suggestion);
    }
  });
  dateKeys.forEach(dk => updateShuttleSuggestionForDate(dk));
}

function wireShiftSelects() {
  document.querySelectorAll('.shift-select').forEach(sel => {
    sel.dataset.prevValue = sel.value;
    applySuggestionOptionState(sel, 'selected');
    sel.addEventListener('focus', () => {
      applySuggestionOptionState(sel, 'menu');
      sel.dataset.prevValue = sel.value;
    });
    sel.addEventListener('blur', () => applySuggestionOptionState(sel, 'selected'));
    sel.addEventListener('change', async (e) => {
      const cell = sel.closest('.cell');
      const section = cell.getAttribute('data-section');
      const employee = cell.getAttribute('data-employee');
      const dateKey = cell.getAttribute('data-date');
      const previousValue = sel.dataset.prevValue || 'Set';
      let value = sel.value;
      if (sel.dataset.allowCustom === '1' && value === CUSTOM_SHIFT_VALUE) {
        const custom = window.prompt('Enter the time window (e.g., 5:30am - 1:30pm)');
        if (!custom || !custom.trim()) {
          sel.value = previousValue;
          applySuggestionOptionState(sel, 'selected');
          return;
        }
        const cleaned = custom.trim();
        let existing = Array.from(sel.options).find(opt => opt.value === cleaned);
        if (!existing) {
          existing = document.createElement('option');
          existing.value = cleaned;
          existing.textContent = cleaned;
          existing.dataset.custom = '1';
          sel.appendChild(existing);
        }
        value = cleaned;
        sel.value = cleaned;
        updateSelectClass(sel, section, value);
      }
      try {
        const res = await fetch('/assign', {
          method: 'POST', headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ section, employee, date: dateKey, value, week_id: window.currentWeekId }),
        });
        const data = await res.json();
        if (!data.ok) {
          if (data.code === 'timeoff') {
            // Force TIME OFF but keep dropdown enabled so user can change
            ensureTimeOffOptions(sel);
            const desired = TIME_OFF_VALUES.has(data.value) ? data.value : 'TIME OFF';
            sel.value = desired;
            updateSelectClass(sel, section, sel.value);
            updateCoverageUI(data);
            updateConflictsUI();
            showToast('Blocked: approved time off');
            return;
          }
          throw new Error(data.error || 'Save failed');
        }
        updateSelectClass(sel, section, value);
        updateCoverageUI(data);
        updateConflictsUI();
        showToast('Saved');
        applySuggestionOptionState(sel, 'selected');
        if (previousValue !== value) {
          const prevOption = Array.from(sel.options).find(opt => opt.value === previousValue && opt.dataset.custom === '1');
          if (prevOption) {
            prevOption.remove();
          }
        }
        sel.dataset.prevValue = value;
      } catch (err) {
        console.error(err);
        const message = err && typeof err.message === 'string' && err.message.trim() ? err.message : 'Save failed';
        showToast(message);
        sel.value = previousValue;
        updateSelectClass(sel, section, sel.value);
        applySuggestionOptionState(sel, 'selected');
      }
    });
  });
}

function wireTimeOff() {
  document.querySelectorAll('.timeoff-item').forEach(item => {
    const id = Number(item.getAttribute('data-id'));
    const toggle = item.querySelector('.timeoff-toggle');
    const typeToggle = item.querySelector('.timeoff-type-toggle');
    const status = item.querySelector('[data-status]');
    const deleteBtn = item.querySelector('.timeoff-delete');
    
    toggle.addEventListener('change', async () => {
      try {
        const res = await fetch('/timeoff/toggle', {
          method: 'POST', headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ id, approved: toggle.checked }),
        });
        const data = await res.json();
        if (!data.ok) throw new Error(data.error || 'Failed');
        if (data.item.approved) {
          status.textContent = '✔️';
          status.classList.remove('pending');
          status.classList.add('approved');
        } else {
          status.textContent = 'pending';
          status.classList.remove('approved');
          status.classList.add('pending');
        }
        // Update cells for this employee over the date range
        const name = data.item.name;
        const from = new Date(data.item.from);
        const to = new Date(data.item.to);
        const allCells = document.querySelectorAll(`.cell[data-employee="${name}"] select`);
        allCells.forEach(sel => {
          const cell = sel.closest('.cell');
          const dk = cell.getAttribute('data-date');
          const d = new Date(dk);
          if (d >= from && d <= to) {
            if (data.item.approved) {
              ensureTimeOffOptions(sel);
              const desired = data.item.vacation ? 'REQ VAC' : 'TIME OFF';
              sel.value = desired;
              updateSelectClass(sel, cell.getAttribute('data-section'), sel.value);
            } else {
              if (TIME_OFF_VALUES.has(sel.value)) sel.value = 'Set';
              updateSelectClass(sel, cell.getAttribute('data-section'), sel.value);
            }
          }
        });
        // Update coverage counts (server included counts/missing)
        updateCoverageUI(data);
        showToast('Time off updated');
      } catch (err) {
        console.error(err);
        showToast('Update failed');
        toggle.checked = !toggle.checked; // revert
      }
    });
    
    // Delete button functionality
    if (deleteBtn) {
      deleteBtn.addEventListener('click', async () => {
        if (!confirm('Are you sure you want to delete this time off request?')) {
          return;
        }
        
        try {
          const res = await fetch(`/timeoff/delete/${id}`, {
            method: 'POST', headers: { 'Content-Type': 'application/json' },
          });
          const data = await res.json();
          if (!data.ok) throw new Error(data.error || 'Failed');
          
          // Remove the item from the DOM
          item.remove();
          
          // Update coverage counts
          updateCoverageUI(data);
          updateConflictsUI();
          showToast('Time off request deleted');
        } catch (err) {
          console.error(err);
          showToast('Delete failed');
        }
      });
    }

    // Wire vacation type toggle persistence
    if (typeToggle) {
      typeToggle.addEventListener('change', async () => {
        try {
          const res = await fetch('/timeoff/vacation', {
            method: 'POST', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ id, vacation: typeToggle.checked }),
          });
          const data = await res.json();
          if (!data.ok) throw new Error(data.error || 'Failed');
          const payload = data.item;
          if (toggle.checked) {
            const name = payload.name;
            const from = new Date(payload.from);
            const to = new Date(payload.to);
            const selects = document.querySelectorAll(`.cell[data-employee="${name}"] select`);
            selects.forEach(sel => {
              const cell = sel.closest('.cell');
              const dk = cell.getAttribute('data-date');
              if (!dk) return;
              const d = new Date(dk);
              if (d >= from && d <= to && payload.approved) {
                ensureTimeOffOptions(sel);
                sel.value = payload.vacation ? 'REQ VAC' : 'TIME OFF';
                updateSelectClass(sel, cell.getAttribute('data-section'), sel.value);
              }
            });
            updateConflictsUI();
          }
          showToast('Vacation updated');
        } catch (err) {
          console.error(err);
          showToast('Update failed');
          // revert UI state if failed
          typeToggle.checked = !typeToggle.checked;
        }
      });
    }
  });
}

function applyTimeOffUIUpdate(name, fromIso, toIso, approved, vacation) {
  if (!name || !fromIso || !toIso) return;
  const from = new Date(fromIso);
  const to = new Date(toIso);
  const allCells = document.querySelectorAll(`.cell[data-employee="${name}"] select`);
  allCells.forEach(sel => {
    const cell = sel.closest('.cell');
    const dk = cell.getAttribute('data-date');
    if (!dk) return;
    const d = new Date(dk);
    if (d >= from && d <= to) {
      if (approved) {
        ensureTimeOffOptions(sel);
        sel.value = vacation ? 'REQ VAC' : 'TIME OFF';
      } else {
        if (TIME_OFF_VALUES.has(sel.value)) sel.value = 'Set';
      }
      updateSelectClass(sel, cell.getAttribute('data-section'), sel.value);
    }
  });
}

const AIRCREW_TIME_DEFAULT = '18:00';
let aircrewTimePicker = null;

function formatAircrewTimeLabel(value) {
  if (!value) return '';
  const [h, m] = value.split(':').map(Number);
  if (Number.isNaN(h) || Number.isNaN(m)) return value;
  const suffix = h < 12 ? 'am' : 'pm';
  const displayHour = (h % 12) || 12;
  return `${displayHour}:${String(m).padStart(2, '0')}${suffix}`;
}

function normalizeAircrewTimeValue(value) {
  if (!value) return null;
  const [rawHour, rawMinute] = value.split(':');
  const hour = Number(rawHour);
  const minute = Number(rawMinute);
  if (Number.isNaN(hour) || Number.isNaN(minute)) return null;
  const h = Math.min(23, Math.max(0, hour));
  const m = Math.min(59, Math.max(0, minute));
  return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
}

const AIRCREW_RECENT_STORAGE_KEY = 'workScheduler.aircrewRecentTimes';
const AIRCREW_RECENT_MAX = 6;
let aircrewRecentTimes = [];
const aircrewRecentsSubscribers = [];

function loadAircrewRecentTimes() {
  if (typeof window === 'undefined' || !window.localStorage) {
    return [];
  }
  try {
    const raw = window.localStorage.getItem(AIRCREW_RECENT_STORAGE_KEY);
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed)) {
      return parsed
        .map(item => normalizeAircrewTimeValue(item))
        .filter(Boolean);
    }
  } catch (err) {
    console.warn('Unable to load aircrew recents', err);
  }
  return [];
}

function saveAircrewRecentTimes(list) {
  if (typeof window === 'undefined' || !window.localStorage) return;
  try {
    window.localStorage.setItem(AIRCREW_RECENT_STORAGE_KEY, JSON.stringify(list));
  } catch (err) {
    console.warn('Unable to save aircrew recents', err);
  }
}

function subscribeAircrewRecents(fn) {
  if (typeof fn !== 'function') return;
  aircrewRecentsSubscribers.push(fn);
  try {
    fn([...aircrewRecentTimes]);
  } catch (err) {
    console.warn('Aircrew recents subscriber error', err);
  }
}

function notifyAircrewRecents() {
  aircrewRecentsSubscribers.forEach(fn => {
    try {
      fn([...aircrewRecentTimes]);
    } catch (err) {
      console.warn('Aircrew recents subscriber error', err);
    }
  });
}

function recordAircrewRecentTime(value) {
  const normalized = normalizeAircrewTimeValue(value);
  if (!normalized) return;
  if (aircrewRecentTimes.includes(normalized)) {
    return;
  }
  const next = [normalized, ...aircrewRecentTimes].slice(0, AIRCREW_RECENT_MAX);
  aircrewRecentTimes = next;
  saveAircrewRecentTimes(aircrewRecentTimes);
  notifyAircrewRecents();
}

function clearAircrewRecentTimes() {
  if (!aircrewRecentTimes.length) return;
  aircrewRecentTimes = [];
  saveAircrewRecentTimes(aircrewRecentTimes);
  notifyAircrewRecents();
}

aircrewRecentTimes = loadAircrewRecentTimes();

function renderAircrewCell(cell, times) {
  if (!cell) return;
  const chipsWrap = cell.querySelector('.aircrew-chips');
  if (!chipsWrap) return;
  chipsWrap.innerHTML = '';
  const list = Array.isArray(times) ? times : [];
  list.forEach(time => {
    const chip = document.createElement('span');
    chip.className = 'aircrew-chip';
    chip.setAttribute('data-time', time);
    const label = document.createElement('span');
    label.className = 'aircrew-chip-label';
    label.textContent = formatAircrewTimeLabel(time);
    const removeBtn = document.createElement('button');
    removeBtn.type = 'button';
    removeBtn.className = 'aircrew-chip-remove';
    removeBtn.setAttribute('aria-label', `Remove ${label.textContent}`);
    removeBtn.textContent = '×';
    chip.appendChild(label);
    chip.appendChild(removeBtn);
    chipsWrap.appendChild(chip);
  });
  cell.dataset.times = JSON.stringify(list);
  cell.classList.toggle('empty', list.length === 0);
}

async function postAircrewUpdate(action, carrier, dateKey, extra = {}) {
  if (!window.currentWeekId) {
    throw new Error('Week not found');
  }
  const payload = {
    carrier,
    date: dateKey,
    action,
    week_id: window.currentWeekId,
    ...extra,
  };
  const res = await fetch('/aircrew/arrival', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok || !data?.ok) {
    throw new Error(data?.error || 'Unable to update aircrew arrivals');
  }
  return data;
}

function applyAircrewCells(carrier, cells) {
  if (!cells) return;
  const updatedDates = new Set();
  Object.entries(cells).forEach(([dateKey, times]) => {
    const cell = document.querySelector(`.aircrew-table .cell[data-carrier="${carrier}"][data-date="${dateKey}"]`);
    if (!cell) return;
    renderAircrewCell(cell, Array.isArray(times) ? times : []);
    updatedDates.add(dateKey);
  });
  updatedDates.forEach(updateShuttleSuggestionForDate);
}

function renderOccupancyCell(cell, value) {
  if (!cell) return;
  const input = cell.querySelector('.occupancy-input');
  const hint = cell.querySelector('.occupancy-cell-hint');
  let cleaned = '';
  if (value !== null && value !== undefined && !Number.isNaN(Number(value))) {
    const numericValue = Number(value);
    const clamped = Math.max(0, Math.min(100, numericValue));
    cleaned = clamped.toFixed(2);
  }
  if (input) {
    input.value = cleaned;
  }
  if (hint) {
    hint.textContent = cleaned ? `Saved ${cleaned}%` : 'Not set';
  }
  if (cleaned) {
    cell.dataset.value = cleaned;
    cell.classList.remove('empty');
  } else {
    delete cell.dataset.value;
    cell.classList.add('empty');
  }
}

function applyOccupancyCells(cells) {
  if (!cells) return;
  Object.entries(cells).forEach(([dateKey, value]) => {
    const cell = document.querySelector(`.occupancy-cell[data-date="${dateKey}"]`);
    if (cell) {
      renderOccupancyCell(cell, value);
    }
  });
}

async function postOccupancyUpdate(dateKey, value) {
  if (!window.currentWeekId) {
    throw new Error('Week not found');
  }
  const payload = {
    date: dateKey,
    value: value === null || value === undefined ? null : value,
    week_id: window.currentWeekId,
  };
  const res = await fetch('/occupancy', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok || !data?.ok) {
    throw new Error(data?.error || 'Unable to update occupancy');
  }
  return data;
}

function wireOccupancyInputs() {
  const inputs = document.querySelectorAll('.occupancy-input');
  if (!inputs.length) return;
  inputs.forEach(input => {
    const cell = input.closest('.occupancy-cell');
    if (!cell) return;
    const dateKey = cell.getAttribute('data-date');
    if (!dateKey) return;

    const setBusy = (busy) => {
      if (busy) {
        input.setAttribute('disabled', 'disabled');
      } else {
        input.removeAttribute('disabled');
      }
    };

    const saveValue = async () => {
      const raw = input.value.trim();
      let value = null;
      if (raw !== '') {
        const num = Number(raw);
        if (Number.isNaN(num)) {
          showToast('Enter a number between 0 and 100');
          renderOccupancyCell(cell, cell.dataset.value || null);
          return;
        }
        const clamped = Math.max(0, Math.min(100, num));
        value = clamped;
        const normalized = clamped.toFixed(2);
        if (input.value !== normalized) {
          input.value = normalized;
        }
      }
      if (cell.dataset.saving === '1') return;
      cell.dataset.saving = '1';
      setBusy(true);
      try {
        const data = await postOccupancyUpdate(dateKey, value);
        renderOccupancyCell(cell, data.value);
        showToast(value === null ? 'Occupancy cleared' : 'Occupancy updated');
      } catch (err) {
        console.error(err);
        const message = err && err.message ? err.message : 'Unable to update occupancy';
        showToast(message);
      } finally {
        delete cell.dataset.saving;
        setBusy(false);
      }
    };

    input.addEventListener('change', saveValue);
    input.addEventListener('keydown', (evt) => {
      if (evt.key === 'Enter') {
        evt.preventDefault();
        saveValue();
      }
    });
  });
}

function wireOccupancyUpload() {
  const trigger = document.getElementById('occupancy-upload-trigger');
  const input = document.getElementById('occupancy-upload-input');
  const status = document.getElementById('occupancy-upload-status');
  if (!trigger || !input) return;

  const setStatus = (text, state) => {
    if (!status) return;
    if (state) {
      status.setAttribute('data-state', state);
    } else {
      status.removeAttribute('data-state');
    }
    status.textContent = text || '';
  };

  const resetInputs = () => {
    input.value = '';
    input.disabled = false;
    trigger.disabled = false;
  };

  trigger.addEventListener('click', () => {
    if (trigger.disabled) return;
    input.click();
  });

  input.addEventListener('change', async () => {
    if (!input.files || !input.files.length) return;
    const file = input.files[0];
    const formData = new FormData();
    formData.append('file', file);
    if (window.currentWeekId) {
      formData.append('week_id', window.currentWeekId);
    }
    trigger.disabled = true;
    input.disabled = true;
    setStatus('Uploading…', 'info');
    try {
      const res = await fetch('/occupancy/import', {
        method: 'POST',
        body: formData,
      });
      const data = await res.json().catch(() => ({}));
      if (!res.ok || !data?.ok) {
        throw new Error(data?.error || 'Upload failed');
      }
      if (data.week_cells) {
        applyOccupancyCells(data.week_cells);
      }
      const message = data.message || `Updated ${data.updated_cells || 0} days`;
      setStatus('Upload complete', 'success');
      showToast(message);
    } catch (err) {
      console.error(err);
      const msg = err && err.message ? err.message : 'Upload failed';
      setStatus(msg, 'error');
      showToast(msg);
    } finally {
      resetInputs();
      setTimeout(() => setStatus('', ''), 4000);
    }
  });
}

function wireScheduleTemplateUpload() {
  const form = document.getElementById('schedule-template-upload-form');
  const trigger = document.getElementById('schedule-template-upload-trigger');
  const input = document.getElementById('schedule-template-upload-input');
  if (!form || !trigger || !input) return;
  trigger.addEventListener('click', () => {
    if (trigger.disabled) return;
    input.click();
  });
  input.addEventListener('change', () => {
    if (!input.files || !input.files.length) return;
    trigger.disabled = true;
    trigger.textContent = 'Uploading…';
    form.submit();
  });
}

function initAircrewTimePicker() {
  const modal = document.getElementById('aircrew-time-modal');
  if (!modal) return null;
  const overlay = modal.querySelector('.aircrew-time-overlay');
  const closeBtn = modal.querySelector('.aircrew-time-close');
  const cancelBtn = modal.querySelector('.aircrew-time-cancel');
  const confirmBtn = modal.querySelector('.aircrew-time-confirm');
  const input = modal.querySelector('#aircrew-time-input');
  const recentsWrap = modal.querySelector('#aircrew-time-recents');
  const recentsEmpty = modal.querySelector('#aircrew-time-recents-empty');
  const clearRecentsBtn = modal.querySelector('#aircrew-time-recents-clear');

  const state = {
    current: null,
    lastValue: aircrewRecentTimes[0] || AIRCREW_TIME_DEFAULT,
  };

  const setValue = (value) => {
    const normalized = normalizeAircrewTimeValue(value) || AIRCREW_TIME_DEFAULT;
    state.lastValue = normalized;
    if (input) {
      input.value = normalized;
    }
  };

  const close = () => {
    modal.setAttribute('aria-hidden', 'true');
    document.body.classList.remove('aircrew-time-open');
    state.current = null;
  };

  const confirmSelection = () => {
    const raw = input?.value || state.lastValue || AIRCREW_TIME_DEFAULT;
    const normalized = normalizeAircrewTimeValue(raw);
    if (!normalized) {
      showToast('Enter a valid time (HH:MM)');
      input?.focus();
      return;
    }
    state.lastValue = normalized;
    const cb = state.current?.onConfirm;
    close();
    if (typeof cb === 'function') {
      cb({ time24: normalized });
    }
  };

  const open = ({ onConfirm, prefill } = {}) => {
    state.current = { onConfirm };
    const fallback = state.lastValue || aircrewRecentTimes[0] || AIRCREW_TIME_DEFAULT;
    setValue(prefill || fallback);
    modal.setAttribute('aria-hidden', 'false');
    document.body.classList.add('aircrew-time-open');
    window.setTimeout(() => {
      if (input) {
        input.focus();
        input.select();
      } else {
        confirmBtn?.focus();
      }
    }, 15);
  };

  overlay?.addEventListener('click', close);
  closeBtn?.addEventListener('click', close);
  cancelBtn?.addEventListener('click', close);
  confirmBtn?.addEventListener('click', confirmSelection);
  const renderRecents = (recents) => {
    if (!recentsWrap) return;
    recentsWrap.innerHTML = '';
    if (!recents.length) {
      recentsWrap.hidden = true;
      if (recentsEmpty) recentsEmpty.hidden = false;
      if (clearRecentsBtn) clearRecentsBtn.hidden = true;
      return;
    }
    recentsWrap.hidden = false;
    if (recentsEmpty) recentsEmpty.hidden = true;
    if (clearRecentsBtn) clearRecentsBtn.hidden = false;
    recents.forEach((time) => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'aircrew-time-preset';
      btn.setAttribute('data-value', time);
      btn.textContent = formatAircrewTimeLabel(time);
      btn.addEventListener('click', () => {
        // Selecting a preset immediately confirms to avoid the extra Add click
        setValue(time);
        confirmSelection();
      });
      recentsWrap.appendChild(btn);
    });
  };
  subscribeAircrewRecents(renderRecents);
  clearRecentsBtn?.addEventListener('click', () => {
    if (!aircrewRecentTimes.length) return;
    clearAircrewRecentTimes();
    showToast('Recent quick times removed');
  });
  input?.addEventListener('input', () => {
    const normalized = normalizeAircrewTimeValue(input.value);
    if (normalized) {
      state.lastValue = normalized;
    }
  });
  modal.addEventListener('keydown', (evt) => {
    if (modal.getAttribute('aria-hidden') === 'true') return;
    if (evt.key === 'Escape') {
      evt.preventDefault();
      close();
    } else if (evt.key === 'Enter') {
      evt.preventDefault();
      confirmSelection();
    }
  });

  return { open, close };
}

function wireAircrewArrivals() {
  const cells = document.querySelectorAll('.aircrew-table .cell[data-carrier]');
  if (!cells.length) return;
  cells.forEach(cell => {
    const carrier = cell.getAttribute('data-carrier');
    const dateKey = cell.getAttribute('data-date');
    const trigger = cell.querySelector('.aircrew-wheel-trigger');

    const setBusy = (busy) => {
      if (busy) {
        cell.dataset.saving = '1';
      } else {
        delete cell.dataset.saving;
      }
    };

    const addArrival = async (timeValue) => {
      if (!timeValue) return;
      if (cell.dataset.saving === '1') return;
      setBusy(true);
      try {
        const data = await postAircrewUpdate('add', carrier, dateKey, { time: timeValue });
        applyAircrewCells(carrier, data.cells);
        recordAircrewRecentTime(timeValue);
        showToast('Arrival added');
      } catch (err) {
        console.error(err);
        showToast(err && err.message ? err.message : 'Unable to add arrival');
      } finally {
        setBusy(false);
      }
    };

    if (trigger && aircrewTimePicker) {
      trigger.addEventListener('click', () => {
        if (cell.dataset.saving === '1') return;
        aircrewTimePicker.open({
          onConfirm: ({ time24 }) => addArrival(time24),
        });
      });
    }

    cell.addEventListener('click', async (evt) => {
      const removeBtn = evt.target.closest('.aircrew-chip-remove');
      if (!removeBtn) return;
      const chip = removeBtn.closest('.aircrew-chip');
      const timeValue = chip ? chip.getAttribute('data-time') : null;
      if (!timeValue || cell.dataset.saving === '1') return;
      setBusy(true);
      try {
        const data = await postAircrewUpdate('remove', carrier, dateKey, { time: timeValue });
        applyAircrewCells(carrier, data.cells);
        showToast('Arrival removed');
      } catch (err) {
        console.error(err);
        showToast(err && err.message ? err.message : 'Unable to remove arrival');
      } finally {
        setBusy(false);
      }
    });

  });
}


function initLiveUpdates() {
  try {
    const es = new EventSource('/events');
    es.addEventListener('message', (evt) => {
      if (!evt?.data) return;
      let payload;
      try { payload = JSON.parse(evt.data); } catch (_) { return; }
      if (payload?.type === 'timeoff' && payload?.item) {
        const it = payload.item;
        // Update UI for affected employee/date range
        applyTimeOffUIUpdate(it.name, it.from, it.to, !!it.approved, !!it.vacation);
        // Update coverage UI if included
        if (payload.counts) {
          updateCoverageUI(payload);
        }
        updateConflictsUI();
      } else if (payload?.type === 'aircrew') {
        if (payload.week_id && window.currentWeekId && Number(payload.week_id) !== Number(window.currentWeekId)) {
          return;
        }
        const batch = Array.isArray(payload.batch) ? payload.batch : [payload];
        batch.forEach(item => {
          if (!item || !item.carrier || !item.date) return;
          const times = Array.isArray(item.times) ? item.times : [];
          const cell = document.querySelector(`.aircrew-table .cell[data-carrier="${item.carrier}"][data-date="${item.date}"]`);
          if (cell && !document.body.classList.contains('aircrew-time-open')) {
            renderAircrewCell(cell, times);
          }
          updateShuttleSuggestionForDate(item.date);
        });
      } else if (payload?.type === 'occupancy') {
        if (payload.week_id && window.currentWeekId && Number(payload.week_id) !== Number(window.currentWeekId)) {
          return;
        }
        const entries = Array.isArray(payload.items) ? payload.items : [payload];
        const cells = {};
        entries.forEach(item => {
          if (!item || !item.date) return;
          cells[item.date] = item.value;
        });
        applyOccupancyCells(cells);
      }
    });
  } catch (e) {
    // SSE may be unsupported; fail silently
    console.warn('Live updates unavailable', e);
  }
}

function wireWeekNav() {
  const prev = document.getElementById('prev-week');
  const next = document.getElementById('next-week');
  const currentWeekId = window.currentWeekId;
  
  if (prev && currentWeekId) {
    prev.addEventListener('click', () => {
      window.location.href = `/week/${currentWeekId}/prev`;
    });
  }
  
  if (next && currentWeekId) {
    next.addEventListener('click', () => {
      window.location.href = `/week/${currentWeekId}/next`;
    });
  }
}

function initUndoCountdown() {
  const banner = document.querySelector('[data-undo-deadline]');
  if (!banner) return;
  const deadlineAttr = banner.getAttribute('data-undo-deadline');
  if (!deadlineAttr) return;
  const deadlineMs = parseInt(deadlineAttr, 10);
  if (!deadlineMs) return;
  const secondsEl = banner.querySelector('[data-undo-seconds]');
  const undoButton = banner.querySelector('button');

  const updateCountdown = () => {
    const remainingMs = deadlineMs - Date.now();
    const remaining = Math.max(0, Math.round(remainingMs / 1000));
    if (secondsEl) {
      secondsEl.textContent = remaining.toString();
    }
    if (remaining <= 0) {
      if (undoButton) {
        undoButton.disabled = true;
        undoButton.classList.remove('secondary');
        undoButton.classList.add('ghost');
        undoButton.textContent = 'Undo expired';
      }
      banner.removeAttribute('data-undo-deadline');
      return true;
    }
    return false;
  };

  if (updateCountdown()) return;
  const timer = setInterval(() => {
    if (updateCountdown()) {
      clearInterval(timer);
    }
  }, 1000);
}

function wireGenerateSchedule() {
  const form = document.getElementById('generate-form');
  if (!form) return;
  
  // Check if schedule_generated is available from the template
  const scheduleGenerated = window.scheduleGenerated || false;
  
  form.addEventListener('submit', (e) => {
    if (scheduleGenerated) {
      const confirmed = confirm(
        'This will override the existing schedule with a new generated schedule. ' +
        'All current shift assignments will be replaced. Are you sure you want to continue?'
      );
      if (!confirmed) {
        e.preventDefault();
      }
    }
  });
}

function confirmGenerateSchedule() {
  return confirm(
    'This will generate a new schedule for the current week. ' +
    'If a schedule already exists, it will be overridden with new assignments. ' +
    'Are you sure you want to continue?'
  );
}

function wireEmployeeSorting() {
  const bodies = document.querySelectorAll('.employee-table-body[data-section-id]');
  if (!bodies.length) return;

  let draggedRow = null;
  let dragSourceBody = null;
  let startOrder = null;

  const orderFor = (tbody) =>
    Array.from(tbody.querySelectorAll('tr[data-employee-id]'))
      .map(row => row.getAttribute('data-employee-id'))
      .join(',');

  const restoreOrder = (tbody, orderString) => {
    if (!orderString) return;
    const ids = orderString.split(',').filter(Boolean);
    if (!ids.length) return;
    const rowsById = {};
    tbody.querySelectorAll('tr[data-employee-id]').forEach(row => {
      rowsById[row.getAttribute('data-employee-id')] = row;
    });
    ids.forEach(id => {
      const row = rowsById[id];
      if (row) tbody.appendChild(row);
    });
  };

  bodies.forEach(tbody => {
    tbody.querySelectorAll('tr[data-employee-id]').forEach(row => {
      row.addEventListener('dragstart', (e) => {
        draggedRow = row;
        dragSourceBody = tbody;
        startOrder = orderFor(tbody);
        row.classList.add('dragging');
        if (e.dataTransfer) {
          e.dataTransfer.effectAllowed = 'move';
          e.dataTransfer.setData('text/plain', row.getAttribute('data-employee-id') || '');
        }
      });

      row.addEventListener('dragend', () => {
        row.classList.remove('dragging');
        if (dragSourceBody && startOrder) {
          restoreOrder(dragSourceBody, startOrder);
        }
        draggedRow = null;
        dragSourceBody = null;
        startOrder = null;
      });
    });

    tbody.addEventListener('dragover', (e) => {
      if (!draggedRow || dragSourceBody !== tbody) return;
      e.preventDefault();
      const targetRow = e.target.closest('tr[data-employee-id]');
      if (!targetRow || targetRow === draggedRow) return;
      const rect = targetRow.getBoundingClientRect();
      const before = e.clientY < rect.top + rect.height / 2;
      tbody.insertBefore(draggedRow, before ? targetRow : targetRow.nextSibling);
    });

    tbody.addEventListener('drop', async (e) => {
      if (!draggedRow || dragSourceBody !== tbody) return;
      e.preventDefault();
      draggedRow.classList.remove('dragging');
      const newOrder = orderFor(tbody);
      const sectionId = Number(tbody.getAttribute('data-section-id'));
      const previousOrder = startOrder;
      draggedRow = null;
      dragSourceBody = null;
      startOrder = null;
      if (!sectionId || newOrder === previousOrder) {
        return;
      }
      const ids = newOrder
        .split(',')
        .filter(Boolean)
        .map(id => Number(id));
      try {
        const res = await fetch('/admin/employees/reorder', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ section_id: sectionId, employee_ids: ids }),
        });
        const data = await res.json().catch(() => ({ ok: false }));
        if (!res.ok || !data?.ok) {
          throw new Error(data?.error || 'Save failed');
        }
        showToast('Order saved');
      } catch (err) {
        console.error(err);
        const message = err && typeof err.message === 'string' && err.message.trim()
          ? err.message
          : 'Unable to save order';
        showToast(message);
        restoreOrder(tbody, previousOrder);
      }
    });
  });
}

function wireEmployeeRoleChanges() {
  document.querySelectorAll('.employee-role-select[data-employee-id]').forEach(sel => {
    const employeeId = Number(sel.getAttribute('data-employee-id'));
    sel.addEventListener('change', async () => {
      if (!employeeId) return;
      const previous = sel.getAttribute('data-current-role');
      const sectionId = Number(sel.value);
      try {
        const res = await fetch(`/admin/employees/${employeeId}/role`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ section_id: sectionId }),
        });
        const data = await res.json().catch(() => ({ ok: false }));
        if (!res.ok || !data?.ok) {
          throw new Error(data?.error || 'Unable to update role');
        }
        sel.setAttribute('data-current-role', String(sectionId));
        const allRoles = Array.isArray(window.employeeRoleOptions) ? window.employeeRoleOptions : [];
        const row = sel.closest('tr');
        const secondarySelect = row ? row.querySelector('.employee-secondary-select') : null;
        if (secondarySelect) {
          const previousSelected = secondarySelect.getAttribute('data-selected-secondary') || '';
          const filteredSelected = previousSelected === String(sectionId) ? '' : previousSelected;
          secondarySelect.innerHTML = '';

          const noneOption = document.createElement('option');
          noneOption.value = '';
          noneOption.textContent = 'None';
          if (!filteredSelected) {
            noneOption.selected = true;
          }
          secondarySelect.appendChild(noneOption);

          allRoles.forEach(role => {
            if (!role || role.id === undefined || role.name === undefined) return;
            if (Number(role.id) === sectionId) return;
            const opt = document.createElement('option');
            opt.value = String(role.id);
            opt.textContent = role.name;
            if (filteredSelected && String(role.id) === filteredSelected) {
              opt.selected = true;
            }
            secondarySelect.appendChild(opt);
          });

          // Ensure selection persists or falls back to none
          if (filteredSelected) {
            const match = secondarySelect.querySelector(`option[value="${filteredSelected}"]`);
            if (match) {
              match.selected = true;
            } else {
              secondarySelect.value = '';
            }
          } else {
            secondarySelect.value = '';
          }
          secondarySelect.setAttribute('data-selected-secondary', filteredSelected);
        }
        showToast('Role updated');
      } catch (err) {
        console.error(err);
        const message = err && typeof err.message === 'string' && err.message.trim()
          ? err.message
          : 'Unable to update role';
        showToast(message);
        if (previous !== null) {
          sel.value = previous;
        }
      }
    });
  });
}

function wireEmployeeSecondaryRoles() {
  document.querySelectorAll('.employee-secondary-select[data-employee-id]').forEach(sel => {
    const employeeId = Number(sel.getAttribute('data-employee-id'));
    sel.addEventListener('change', async () => {
      if (!employeeId) return;
      const previous = sel.getAttribute('data-selected-secondary') || '';
      const selectedValue = sel.value;
      const payload = {
        secondary_role: selectedValue ? Number(selectedValue) : null,
      };
      try {
        const res = await fetch(`/admin/employees/${employeeId}/roles`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload),
        });
        const data = await res.json().catch(() => ({ ok: false }));
        if (!res.ok || !data?.ok) {
          throw new Error(data?.error || 'Unable to update secondary role');
        }
        const updated = data.secondary_role;
        sel.setAttribute('data-selected-secondary', updated ? String(updated) : '');
        if (updated) {
          sel.value = String(updated);
        } else {
          sel.value = '';
        }
        showToast('Secondary role updated');
      } catch (err) {
        console.error(err);
        const message = err && typeof err.message === 'string' && err.message.trim()
          ? err.message
          : 'Unable to update secondary role';
        showToast(message);
        sel.value = previous ? previous : '';
        sel.setAttribute('data-selected-secondary', previous);
      }
    });
  });
}

function templateSlotStatusLabel(slotData) {
  if (!slotData || !slotData.has_data) {
    return 'Empty slot';
  }
  const weekLabel = slotData.saved_week_label || 'week';
  if (slotData.updated_label) {
    return `Saved from ${weekLabel} • ${slotData.updated_label}`;
  }
  return `Saved from ${weekLabel}`;
}

function applyTemplateSlotState(slotEl, slotData) {
  if (!slotEl || !slotData) return;
  slotEl.setAttribute('data-has-data', slotData.has_data ? '1' : '0');
  if (slotData.saved_week_label) {
    slotEl.setAttribute('data-week-label', slotData.saved_week_label);
  } else {
    slotEl.removeAttribute('data-week-label');
  }
  if (slotData.updated_label) {
    slotEl.setAttribute('data-updated-label', slotData.updated_label);
  } else {
    slotEl.removeAttribute('data-updated-label');
  }
  const status = slotEl.querySelector('[data-slot-status]');
  if (status) {
    status.textContent = templateSlotStatusLabel(slotData);
  }
  const loadBtn = slotEl.querySelector('.template-slot-load');
  if (loadBtn) {
    loadBtn.disabled = !slotData.has_data;
  }
}

function setTemplateSlotBusy(slotEl, busy) {
  if (!slotEl) return;
  slotEl.classList.toggle('busy', !!busy);
  slotEl.querySelectorAll('button').forEach(btn => {
    if (busy) {
      btn.setAttribute('disabled', 'disabled');
    } else {
      btn.removeAttribute('disabled');
      if (btn.classList.contains('template-slot-load')) {
        const hasData = slotEl.getAttribute('data-has-data') === '1';
        if (!hasData) {
          btn.setAttribute('disabled', 'disabled');
        }
      }
    }
  });
}

function initTemplateControls() {
  const container = document.querySelector('[data-template-slots]');
  if (!container) return;

  const runAction = async (slotEl, action) => {
    if (!slotEl || !action) return;
    const slot = Number(slotEl.getAttribute('data-slot'));
    if (!slot || !window.currentWeekId) {
      showToast('Week not ready');
      return;
    }
    if (action === 'load' && !window.confirm('Load this template over the selected week? Existing assignments will be replaced.')) {
      return;
    }
    setTemplateSlotBusy(slotEl, true);
    try {
      const res = await fetch(`/schedule-templates/${action}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ slot, week_id: window.currentWeekId }),
      });
      const data = await res.json().catch(() => ({}));
      if (!res.ok || !data?.ok) {
        throw new Error(data?.error || 'Request failed');
      }
      if (data.slot) {
        applyTemplateSlotState(slotEl, data.slot);
      }
      if (action === 'save') {
        showToast('Template saved');
      } else {
        showToast('Template applied');
        setTimeout(() => window.location.reload(), 600);
      }
    } catch (err) {
      console.error(err);
      const message = err && err.message ? err.message : 'Unable to update template';
      showToast(message);
    } finally {
      setTemplateSlotBusy(slotEl, false);
    }
  };

  container.querySelectorAll('[data-action]').forEach(btn => {
    btn.addEventListener('click', () => {
      const slotEl = btn.closest('.template-slot');
      if (!slotEl) return;
      const action = btn.getAttribute('data-action');
      if (!action) return;
      if (action === 'load' && btn.disabled) return;
      runAction(slotEl, action);
    });
  });
}

function initWhatsappPasteModal() {
  const trigger = document.getElementById('whatsapp-paste-trigger');
  const modal = document.getElementById('whatsapp-paste-modal');
  if (!trigger || !modal) return;

  const dropzone = modal.querySelector('[data-dropzone]');
  const previewImg = modal.querySelector('[data-preview]');
  const captionInput = modal.querySelector('[data-caption]');
  const feedbackEl = modal.querySelector('[data-feedback]');
  const sendBtn = modal.querySelector('[data-action="send-whatsapp"]');
  const clearBtn = modal.querySelector('[data-action="clear-image"]');
  const closeButtons = modal.querySelectorAll('[data-action="close-whatsapp"]');
  const MAX_BYTES = 8 * 1024 * 1024;
  let currentFile = null;
  let objectUrl = null;
  let pasteListenerAttached = false;

  function setFeedback(message, state) {
    if (!feedbackEl) return;
    feedbackEl.textContent = message || '';
    if (state) {
      feedbackEl.setAttribute('data-state', state);
    } else {
      feedbackEl.removeAttribute('data-state');
    }
  }

  function resetPreview() {
    currentFile = null;
    if (objectUrl) {
      URL.revokeObjectURL(objectUrl);
      objectUrl = null;
    }
    if (previewImg) {
      previewImg.src = '';
      previewImg.setAttribute('hidden', 'hidden');
    }
    if (dropzone) {
      dropzone.classList.remove('has-image', 'dragover');
    }
    if (sendBtn) {
      sendBtn.disabled = true;
    }
    if (clearBtn) {
      clearBtn.disabled = true;
    }
    setFeedback('', null);
  }

  function closeModal() {
    modal.setAttribute('aria-hidden', 'true');
    detachPasteListener();
    resetPreview();
    if (captionInput) {
      captionInput.value = '';
    }
    document.removeEventListener('keydown', onKeydown, true);
  }

  function openModal() {
    modal.setAttribute('aria-hidden', 'false');
    if (dropzone) {
      dropzone.focus();
    }
    attachPasteListener();
    document.addEventListener('keydown', onKeydown, true);
    setFeedback('Paste with Ctrl/Cmd+V or drop an image.', 'info');
  }

  function onKeydown(event) {
    if (event.key === 'Escape') {
      event.preventDefault();
      closeModal();
    }
  }

  function attachPasteListener() {
    if (pasteListenerAttached) return;
    window.addEventListener('paste', handlePasteEvent);
    pasteListenerAttached = true;
  }

  function detachPasteListener() {
    if (!pasteListenerAttached) return;
    window.removeEventListener('paste', handlePasteEvent);
    pasteListenerAttached = false;
  }

  function handleFiles(fileList) {
    if (!fileList || fileList.length === 0) {
      setFeedback('Paste an image first.', 'error');
      return;
    }
    const match = Array.from(fileList).find(file => file && file.type && file.type.startsWith('image/'));
    if (!match) {
      setFeedback('Only image files are supported.', 'error');
      return;
    }
    if (match.size > MAX_BYTES) {
      setFeedback('Image is too large (max 8 MB).', 'error');
      return;
    }
    currentFile = match;
    if (objectUrl) {
      URL.revokeObjectURL(objectUrl);
    }
    objectUrl = URL.createObjectURL(match);
    if (previewImg) {
      previewImg.src = objectUrl;
      previewImg.removeAttribute('hidden');
    }
    if (dropzone) {
      dropzone.classList.add('has-image');
    }
    if (sendBtn) {
      sendBtn.disabled = false;
    }
    if (clearBtn) {
      clearBtn.disabled = false;
    }
    setFeedback('Ready to send.', 'info');
  }

  function handlePasteEvent(event) {
    if (modal.getAttribute('aria-hidden') === 'true') {
      return;
    }
    const files = event.clipboardData && event.clipboardData.files;
    if (files && files.length > 0) {
      event.preventDefault();
      handleFiles(files);
    }
  }

  function handleDrop(event) {
    event.preventDefault();
    if (dropzone) {
      dropzone.classList.remove('dragover');
    }
    if (event.dataTransfer?.files?.length) {
      handleFiles(event.dataTransfer.files);
    }
  }

  function handleDragOver(event) {
    event.preventDefault();
    if (dropzone) {
      dropzone.classList.add('dragover');
    }
  }

  function handleDragLeave(event) {
    if (dropzone && event.target === dropzone) {
      dropzone.classList.remove('dragover');
    }
  }

  async function sendToWhatsapp() {
    if (!currentFile || !sendBtn) {
      setFeedback('Paste an image first.', 'error');
      return;
    }
    const formData = new FormData();
    const filename = currentFile.name || `pasted-${Date.now()}.png`;
    formData.append('image', currentFile, filename);
    const caption = captionInput ? captionInput.value.trim() : '';
    if (caption) {
      formData.append('caption', caption);
    }
    if (window.currentWeekId) {
      formData.append('week_id', window.currentWeekId);
    }
    const originalText = sendBtn.textContent;
    sendBtn.disabled = true;
    sendBtn.textContent = 'Sending…';
    setFeedback('Sending to WhatsApp…', 'info');
    try {
      const response = await fetch('/api/whatsapp/send-image', {
        method: 'POST',
        body: formData,
      });
      let data = null;
      try {
        data = await response.json();
      } catch (err) {
        data = null;
      }
      if (!response.ok || !data?.ok) {
        const message = data?.error || 'Unable to send image.';
        throw new Error(message);
      }
      setFeedback(data?.message || 'Image sent.', 'success');
      sendBtn.textContent = 'Sent';
      setTimeout(() => {
        closeModal();
      }, 1200);
    } catch (error) {
      console.error(error);
      const message = error && error.message ? error.message : 'Unable to send image.';
      setFeedback(message, 'error');
    } finally {
      sendBtn.disabled = false;
      sendBtn.textContent = originalText || 'Send to WhatsApp';
    }
  }

  if (trigger) {
    trigger.addEventListener('click', () => {
      openModal();
    });
  }
  closeButtons.forEach(btn => {
    btn.addEventListener('click', closeModal);
  });
  if (dropzone) {
    dropzone.addEventListener('paste', handlePasteEvent);
    dropzone.addEventListener('drop', handleDrop);
    dropzone.addEventListener('dragover', handleDragOver);
    dropzone.addEventListener('dragleave', handleDragLeave);
  }
  if (clearBtn) {
    clearBtn.addEventListener('click', resetPreview);
  }
  if (sendBtn) {
    sendBtn.addEventListener('click', sendToWhatsapp);
  }
}

// Make functions globally available
window.confirmGenerateSchedule = confirmGenerateSchedule;

document.addEventListener('DOMContentLoaded', () => {
  initThemeToggle();
  wireShiftSelects();
  initShuttleSuggestions();
  wireTimeOff();
  initLiveUpdates();
  wireWeekNav();
  wireGenerateSchedule();
  wireEmployeeSorting();
  wireEmployeeRoleChanges();
  wireEmployeeSecondaryRoles();
  initSelectColors();
  updateConflictsUI();
  initUndoCountdown();
  aircrewTimePicker = initAircrewTimePicker();
  window.aircrewTimePicker = aircrewTimePicker;
  wireAircrewArrivals();
  wireOccupancyInputs();
  wireOccupancyUpload();
  wireScheduleTemplateUpload();
  initTemplateControls();
  initWhatsappPasteModal();
});
