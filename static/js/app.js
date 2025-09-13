function showToast(msg) {
  const el = document.getElementById('toast');
  if (!el) return;
  el.textContent = msg;
  el.classList.add('show');
  setTimeout(() => el.classList.remove('show'), 1600);
}

function selectClassForValue(section, value) {
  if (!value || value === 'Set') return 'select-gray';
  if (value === 'TIME OFF') return 'select-yellow';
  if (section === 'Breakfast Bar') {
    if (value === '5AM–12PM') return 'select-green';
    if (value === '6AM–12PM') return 'select-blue';
    if (value === '7AM–12PM') return 'select-purple';
  }
  if (section === 'Front Desk') {
    if (value.startsWith('Audit')) return 'select-red';
    if (value.startsWith('AM')) return 'select-green';
    if (value.startsWith('PM')) return 'select-blue';
  }
  if (section === 'Shuttle') {
    if (value.startsWith('AM')) return 'select-green';
    if (value.startsWith('Midday')) return 'select-blue';
    if (value.startsWith('PM')) return 'select-purple';
    if (value.startsWith('Crew')) return 'select-red';
  }
  return '';
}

function updateSelectClass(selectEl, section, value) {
  selectEl.classList.remove('select-green', 'select-blue', 'select-red', 'select-gray', 'select-yellow', 'select-purple');
  const cls = selectClassForValue(section, value);
  if (cls) selectEl.classList.add(cls);
}

function initSelectColors() {
  document.querySelectorAll('.shift-select').forEach(sel => {
    const cell = sel.closest('.cell');
    const section = cell ? cell.getAttribute('data-section') : '';
    updateSelectClass(sel, section, sel.value);
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
      if (bb_missing?.[dk]) h.classList.add('missing'); else h.classList.remove('missing');
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
    const emp = cell.getAttribute('data-employee');
    const dk = cell.getAttribute('data-date');
    if (!emp || !dk) return;
    const sel = cell.querySelector('select');
    const val = sel ? sel.value : null;
    const active = !!val && val !== 'Set' && val !== 'TIME OFF';
    const key = emp + '||' + dk;
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

function wireShiftSelects() {
  document.querySelectorAll('.shift-select').forEach(sel => {
    sel.addEventListener('change', async (e) => {
      const cell = sel.closest('.cell');
      const section = cell.getAttribute('data-section');
      const employee = cell.getAttribute('data-employee');
      const dateKey = cell.getAttribute('data-date');
      const value = sel.value;
      try {
        const res = await fetch('/assign', {
          method: 'POST', headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ section, employee, date: dateKey, value, week_id: window.currentWeekId }),
        });
        const data = await res.json();
        if (!data.ok) {
          if (data.code === 'timeoff') {
            // Force TIME OFF but keep dropdown enabled so user can change
            sel.value = data.value || 'TIME OFF';
            updateSelectClass(sel, section, sel.value);
            updateCoverageUI(data);
            updateConflictsUI();
            showToast('Blocked: approved time off');
            return;
          }
          throw new Error(data.error || 'Failed');
        }
        updateSelectClass(sel, section, value);
        updateCoverageUI(data);
        updateConflictsUI();
        showToast('Saved');
      } catch (err) {
        console.error(err);
        showToast('Save failed');
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
              // Ensure TIME OFF option exists, then set + disable
              if (![...sel.options].some(o => o.value === 'TIME OFF')) {
                const opt = document.createElement('option');
                opt.value = 'TIME OFF';
                opt.textContent = 'TIME OFF';
                sel.insertBefore(opt, sel.firstChild);
              }
              sel.value = 'TIME OFF';
              updateSelectClass(sel, cell.getAttribute('data-section'), sel.value);
            } else {
              // Re-enable and revert TIME OFF to Set
              if (sel.value === 'TIME OFF') sel.value = 'Set';
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

function applyTimeOffUIUpdate(name, fromIso, toIso, approved) {
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
        // Ensure TIME OFF option exists, then set
        if (![...sel.options].some(o => o.value === 'TIME OFF')) {
          const opt = document.createElement('option');
          opt.value = 'TIME OFF';
          opt.textContent = 'TIME OFF';
          sel.insertBefore(opt, sel.firstChild);
        }
        sel.value = 'TIME OFF';
      } else {
        if (sel.value === 'TIME OFF') sel.value = 'Set';
      }
      updateSelectClass(sel, cell.getAttribute('data-section'), sel.value);
    }
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
        applyTimeOffUIUpdate(it.name, it.from, it.to, !!it.approved);
        // Update coverage UI if included
        if (payload.counts) {
          updateCoverageUI(payload);
        }
        updateConflictsUI();
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

// Make functions globally available
window.confirmGenerateSchedule = confirmGenerateSchedule;

document.addEventListener('DOMContentLoaded', () => {
  wireShiftSelects();
  wireTimeOff();
  initLiveUpdates();
  wireWeekNav();
  wireGenerateSchedule();
  initSelectColors();
  updateConflictsUI();
});
