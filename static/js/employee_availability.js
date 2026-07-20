(function () {
  const form = document.getElementById('employee-availability-form');
  const gridScroll = document.querySelector('.availability-time-grid-scroll');
  if (!form || !gridScroll) return;

  const PIXELS_PER_HOUR = 48;
  const SNAP_MINUTES = 15;
  const MIN_RANGE_MINUTES = 45;
  const dayStart = Number(gridScroll.dataset.dayStart);
  const dayEnd = Number(gridScroll.dataset.dayEnd);
  const aiUrl = gridScroll.dataset.aiUrl;
  const columns = Array.from(document.querySelectorAll('.availability-day-column'));
  const inputContainer = document.getElementById('availability-range-inputs');
  const commandInput = document.getElementById('availability-command');
  const applyButton = document.getElementById('apply-availability-command');
  const status = document.getElementById('availability-command-status');
  const discardButton = document.getElementById('discard-availability-changes');
  const unsaved = document.getElementById('availability-unsaved');
  const initialRanges = JSON.parse(document.getElementById('initial-availability-ranges').textContent || '[]');
  const initialFields = new Map(
    Array.from(form.elements)
      .filter((element) => element.name && element.name !== 'availability_range')
      .map((element) => [element, fieldValue(element)])
  );
  let ranges = normalizeRanges(initialRanges);
  let gesture = null;

  gridScroll.addEventListener('wheel', (event) => {
    if (event.shiftKey || Math.abs(event.deltaX) >= Math.abs(event.deltaY) || event.deltaY === 0) return;
    const multiplier = event.deltaMode === WheelEvent.DOM_DELTA_LINE
      ? 16
      : event.deltaMode === WheelEvent.DOM_DELTA_PAGE
        ? window.innerHeight
        : 1;
    window.scrollBy({top: event.deltaY * multiplier, left: 0, behavior: 'auto'});
    event.preventDefault();
  }, {passive: false});

  function fieldValue(element) {
    if (element.type === 'checkbox' || element.type === 'radio') return element.checked;
    return element.value;
  }

  function snap(value) {
    return Math.round(value / SNAP_MINUTES) * SNAP_MINUTES;
  }

  function clamp(value, minimum, maximum) {
    return Math.min(maximum, Math.max(minimum, value));
  }

  function formatTime(totalMinutes) {
    const minutesInDay = ((totalMinutes % 1440) + 1440) % 1440;
    const hour24 = Math.floor(minutesInDay / 60);
    const minute = minutesInDay % 60;
    const suffix = hour24 < 12 ? 'AM' : 'PM';
    const hour12 = hour24 % 12 || 12;
    return `${hour12}:${String(minute).padStart(2, '0')} ${suffix}`;
  }

  function rangeKey(range) {
    return `${range.day}:${range.start}:${range.end}`;
  }

  function normalizeRanges(source) {
    const byDay = new Map();
    source.forEach((item) => {
      const range = {
        day: clamp(Number(item.day), 0, 6),
        start: clamp(snap(Number(item.start)), dayStart, dayEnd - MIN_RANGE_MINUTES),
        end: clamp(snap(Number(item.end)), dayStart + MIN_RANGE_MINUTES, dayEnd),
      };
      if (range.end - range.start < MIN_RANGE_MINUTES) {
        range.end = Math.min(range.start + MIN_RANGE_MINUTES, dayEnd);
        range.start = Math.min(range.start, range.end - MIN_RANGE_MINUTES);
      }
      if (!byDay.has(range.day)) byDay.set(range.day, []);
      byDay.get(range.day).push(range);
    });
    const result = [];
    byDay.forEach((dayRanges) => {
      dayRanges.sort((a, b) => a.start - b.start);
      dayRanges.forEach((range) => {
        const previous = result[result.length - 1];
        if (previous && previous.day === range.day && range.start <= previous.end) {
          previous.end = Math.max(previous.end, range.end);
        } else {
          result.push({...range});
        }
      });
    });
    return result.sort((a, b) => a.day - b.day || a.start - b.start);
  }

  function rangesSignature(source) {
    return normalizeRanges(source).map(rangeKey).join('|');
  }

  function markChanged() {
    const fieldChanged = Array.from(initialFields.entries()).some(([element, value]) => fieldValue(element) !== value);
    const rangeChanged = rangesSignature(ranges) !== rangesSignature(initialRanges);
    const changed = fieldChanged || rangeChanged;
    form.classList.toggle('has-changes', changed);
    unsaved.textContent = changed ? 'Unsaved changes' : '';
  }

  function blockPosition(range, block) {
    const top = ((range.start - dayStart) / 60) * PIXELS_PER_HOUR;
    const height = ((range.end - range.start) / 60) * PIXELS_PER_HOUR;
    block.style.top = `${top}px`;
    block.style.height = `${height}px`;
  }

  function createBlock(range, preview = false) {
    const block = document.createElement('div');
    block.className = `availability-range-block${preview ? ' is-preview' : ''}`;
    block.dataset.key = rangeKey(range);
    block.tabIndex = preview ? -1 : 0;
    block.innerHTML = `
      <button class="availability-resize-handle top" type="button" aria-label="Adjust start time"></button>
      <span class="availability-range-title">Available</span>
      <span class="availability-range-time"></span>
      <button class="availability-range-remove" type="button" aria-label="Remove availability">&times;</button>
      <button class="availability-resize-handle bottom" type="button" aria-label="Adjust end time"></button>
    `;
    block.querySelector('.availability-range-time').textContent = `${formatTime(range.start)} – ${formatTime(range.end)}`;
    blockPosition(range, block);
    if (!preview) wireBlock(block, range);
    return block;
  }

  function render() {
    columns.forEach((column) => column.querySelectorAll('.availability-range-block').forEach((block) => block.remove()));
    inputContainer.replaceChildren();
    ranges.forEach((range) => {
      columns[range.day].appendChild(createBlock(range));
      const input = document.createElement('input');
      input.type = 'hidden';
      input.name = 'availability_range';
      input.value = `${range.day}::${range.start}::${range.end}`;
      inputContainer.appendChild(input);
    });
    columns.forEach((column) => column.classList.toggle('has-ranges', !!column.querySelector('.availability-range-block')));
    markChanged();
  }

  function minuteAtPointer(column, clientY) {
    const rect = column.getBoundingClientRect();
    return clamp(snap(dayStart + ((clientY - rect.top) / PIXELS_PER_HOUR * 60)), dayStart, dayEnd);
  }

  function creationRange(day, anchor, current = anchor) {
    if (current < anchor || anchor > dayEnd - MIN_RANGE_MINUTES) {
      const end = clamp(Math.max(anchor, dayStart + MIN_RANGE_MINUTES), dayStart + MIN_RANGE_MINUTES, dayEnd);
      return {
        day,
        start: clamp(Math.min(current, end - MIN_RANGE_MINUTES), dayStart, end - MIN_RANGE_MINUTES),
        end,
      };
    }
    const start = clamp(anchor, dayStart, dayEnd - MIN_RANGE_MINUTES);
    return {
      day,
      start,
      end: clamp(Math.max(current, start + MIN_RANGE_MINUTES), start + MIN_RANGE_MINUTES, dayEnd),
    };
  }

  function replaceRange(original, replacement) {
    ranges = ranges.filter((range) => range !== original);
    if (replacement) ranges.push(replacement);
    ranges = normalizeRanges(ranges);
    render();
  }

  function wireBlock(block, range) {
    block.querySelector('.availability-range-remove').addEventListener('click', (event) => {
      event.stopPropagation();
      replaceRange(range, null);
    });
    block.addEventListener('keydown', (event) => {
      if (event.key === 'Delete' || event.key === 'Backspace') {
        event.preventDefault();
        replaceRange(range, null);
      }
    });
    block.addEventListener('pointerdown', (event) => {
      if (event.button !== 0 || event.target.closest('.availability-range-remove')) return;
      event.preventDefault();
      const mode = event.target.classList.contains('top') ? 'resize-start'
        : event.target.classList.contains('bottom') ? 'resize-end' : 'move';
      gesture = {
        mode,
        range,
        column: columns[range.day],
        pointerStart: minuteAtPointer(columns[range.day], event.clientY),
        originalStart: range.start,
        originalEnd: range.end,
        block,
      };
      block.setPointerCapture(event.pointerId);
      block.classList.add('is-dragging');
    });
  }

  columns.forEach((column) => {
    column.addEventListener('pointerdown', (event) => {
      if (event.button !== 0 || event.target.closest('.availability-range-block')) return;
      event.preventDefault();
      const start = minuteAtPointer(column, event.clientY);
      const range = creationRange(Number(column.dataset.day), start);
      const block = createBlock(range, true);
      column.appendChild(block);
      gesture = {mode: 'create', range, column, pointerStart: start, block};
      column.setPointerCapture(event.pointerId);
    });
  });

  document.addEventListener('pointermove', (event) => {
    if (!gesture) return;
    const current = minuteAtPointer(gesture.column, event.clientY);
    if (gesture.mode === 'create') {
      Object.assign(gesture.range, creationRange(gesture.range.day, gesture.pointerStart, current));
    } else if (gesture.mode === 'resize-start') {
      gesture.range.start = clamp(current, dayStart, gesture.originalEnd - MIN_RANGE_MINUTES);
    } else if (gesture.mode === 'resize-end') {
      gesture.range.end = clamp(current, gesture.originalStart + MIN_RANGE_MINUTES, dayEnd);
    } else {
      const duration = gesture.originalEnd - gesture.originalStart;
      const delta = current - gesture.pointerStart;
      const nextStart = clamp(gesture.originalStart + delta, dayStart, dayEnd - duration);
      gesture.range.start = nextStart;
      gesture.range.end = nextStart + duration;
    }
    gesture.block.querySelector('.availability-range-time').textContent = `${formatTime(gesture.range.start)} – ${formatTime(gesture.range.end)}`;
    blockPosition(gesture.range, gesture.block);
  });

  document.addEventListener('pointerup', () => {
    if (!gesture) return;
    const completed = gesture;
    gesture.block.classList.remove('is-dragging');
    gesture = null;
    if (completed.mode === 'create') ranges.push(completed.range);
    ranges = normalizeRanges(ranges);
    render();
  });

  function applyAiOperations(operations) {
    operations.forEach((operation) => {
      const days = new Set((operation.days || []).map(Number).filter((day) => day >= 0 && day <= 6));
      if (!days.size) return;
      if (operation.action === 'clear' || operation.action === 'set') {
        ranges = ranges.filter((range) => !days.has(range.day));
      }
      if (operation.action === 'set' || operation.action === 'add') {
        const start = Number(operation.start);
        const end = Number(operation.end);
        if (start >= dayStart && end <= dayEnd && end - start >= MIN_RANGE_MINUTES) {
          days.forEach((day) => ranges.push({day, start, end}));
        }
      }
    });
    ranges = normalizeRanges(ranges);
    render();
  }

  async function interpretWithAi(command) {
    if (!aiUrl) throw new Error('Natural-language interpretation is not available.');
    const response = await fetch(aiUrl, {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({instruction: command, current_ranges: ranges}),
    });
    let payload = null;
    try {
      payload = await response.json();
    } catch (error) {
      payload = null;
    }
    if (!response.ok || !payload?.ok) {
      throw new Error(payload?.error || 'CLIProxy could not interpret that instruction.');
    }
    applyAiOperations(payload.operations || []);
    const source = payload.interpreter === 'local_fallback'
      ? 'Applied by the outage fallback.'
      : 'Interpreted by CLIProxy.';
    const proxyNotice = payload.proxy_notice ? `${payload.proxy_notice} ` : '';
    status.textContent = `${proxyNotice}${payload.summary || 'Availability updated.'} ${source}`;
  }

  async function applyCommand() {
    const command = (commandInput.value || '').trim();
    if (!command) {
      status.textContent = 'Enter an availability change first.';
      commandInput.focus();
      return;
    }
    const originalText = applyButton.textContent;
    applyButton.disabled = true;
    applyButton.textContent = 'Interpreting…';
    status.textContent = 'Asking CLIProxy to interpret the instruction…';
    try {
      await interpretWithAi(command);
    } catch (error) {
      status.textContent = error?.message || 'CLIProxy could not interpret that instruction.';
    } finally {
      applyButton.disabled = false;
      applyButton.textContent = originalText;
    }
  }

  form.addEventListener('input', markChanged);
  form.addEventListener('change', markChanged);
  applyButton.addEventListener('click', applyCommand);
  commandInput.addEventListener('keydown', (event) => {
    if (event.key === 'Enter') {
      event.preventDefault();
      applyCommand();
    }
  });
  discardButton.addEventListener('click', () => {
    ranges = normalizeRanges(initialRanges);
    form.reset();
    status.textContent = '';
    render();
  });

  render();
})();
