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
  if (section === 'Maintenance') {
    if (value === '8AM–4:30PM') return 'select-green';
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
      if (bb_missing?.[dk]) h.classList.add('missing'); else h.classList.remove('missing');
    });

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
          throw new Error(data.error || 'Save failed');
        }
        updateSelectClass(sel, section, value);
        updateCoverageUI(data);
        updateConflictsUI();
        showToast('Saved');
      } catch (err) {
        console.error(err);
        const message = err && typeof err.message === 'string' && err.message.trim() ? err.message : 'Save failed';
        showToast(message);
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

// Make functions globally available
window.confirmGenerateSchedule = confirmGenerateSchedule;

document.addEventListener('DOMContentLoaded', () => {
  wireShiftSelects();
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
});
