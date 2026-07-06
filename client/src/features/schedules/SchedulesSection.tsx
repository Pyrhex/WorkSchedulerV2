import { useCallback, useEffect, useMemo, useState } from 'react';

import { api } from '../../api/client';
import { Badge, Button, Card, Input, Select } from '../../components/ui';
import type { AircrewCarrierArrivals, Assignment, Week, WeekSummary } from '../../types/api';

const neutralShiftOptions = ['Set', 'OFF', 'N/A', 'TIME OFF', 'REQ VAC'];
const sectionShiftOptions: Record<string, string[]> = {
  'Breakfast Bar': [...neutralShiftOptions, '5AM-12PM', '5AM–12PM', '6AM–12PM', '7AM–12PM'],
  'Front Desk': [
    ...neutralShiftOptions,
    'AM (6:00AM–2:00PM)',
    'AM (6:15AM–2:15PM)',
    'PM (2:00PM–10:00PM)',
    'PM (2:15PM–10:15PM)',
    'Audit (10:00PM–6:00AM)',
    'Audit (10:15PM–6:15AM)',
  ],
  Shuttle: [
    ...neutralShiftOptions,
    'AM (3:30AM–11:30AM)',
    'Crew (10:00AM–6:00PM)',
    'Midday (10:30AM–6:30PM)',
    '10:30am - 6:30pm (c)',
    'PM (5:30PM–1:30AM)',
    'Crew (5:45PM–1:45AM)',
    'Crew (8:00PM–12:00AM)',
    'Crew (9:00PM–1:00AM)',
  ],
  Maintenance: [...neutralShiftOptions, '8AM–4:30PM'],
};

const addDays = (dateText: string, days: number) => {
  const [year, month, day] = dateText.split('-').map(Number);
  const date = new Date(year, month - 1, day + days);
  return [
    date.getFullYear(),
    String(date.getMonth() + 1).padStart(2, '0'),
    String(date.getDate()).padStart(2, '0'),
  ].join('-');
};

const formatDate = (dateText: string) => {
  const [year, month, day] = dateText.split('-').map(Number);
  return new Date(year, month - 1, day).toLocaleDateString(undefined, {
    weekday: 'short',
    month: 'short',
    day: 'numeric',
  });
};

const formatRange = (startDate: string) => {
  const endDate = addDays(startDate, 6);
  return `${formatDate(startDate)} - ${formatDate(endDate)}`;
};

const formatAircrewTime = (value: string) => {
  const [hourValue, minuteValue] = value.split(':').map(Number);
  if (Number.isNaN(hourValue) || Number.isNaN(minuteValue)) return value;
  const suffix = hourValue < 12 ? 'am' : 'pm';
  const displayHour = hourValue % 12 || 12;
  return `${displayHour}:${String(minuteValue).padStart(2, '0')}${suffix}`;
};

const normalizeTimeInput = (value: string) => {
  const [hourValue, minuteValue] = value.split(':').map(Number);
  if (Number.isNaN(hourValue) || Number.isNaN(minuteValue)) return '';
  const hour = Math.min(23, Math.max(0, hourValue));
  const minute = Math.min(59, Math.max(0, minuteValue));
  return `${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}`;
};

const sectionOrder = ['Front Desk', 'Shuttle', 'Breakfast Bar', 'Maintenance'];

const parentheticalTime = (value: string) => {
  const match = value.match(/\(([^)]*\d[^)]*)\)/);
  return match?.[1] ?? null;
};

const normalizeTimeLabel = (value: string) =>
  value
    .replace(/\s*–\s*/g, ' - ')
    .replace(/\s+-\s+/g, ' - ')
    .replace(/\bAM\b|\bPM\b/g, (match) => match.toLowerCase())
    .replace(/\b0(\d:\d{2}(?:am|pm))\b/gi, '$1')
    .replace(/^8:00pm - 12:00am$/i, '8pm - 12am')
    .replace(/^9:00pm - 1:00am$/i, '9pm - 1am');

const formatShiftLabel = (value: string) => {
  if (!value) return '';
  if (value.toLowerCase() === 'set') return '-';
  if (value === 'TIME OFF' || value === 'REQ VAC') return value;
  if (value === '10:30am - 6:30pm (c)') return '10:30am - 6:30pm';

  const withoutPrefix = value.startsWith('Crew Shift ')
    ? value.slice('Crew Shift '.length)
    : value;
  const timeOnly = parentheticalTime(withoutPrefix) ?? withoutPrefix;
  return normalizeTimeLabel(timeOnly);
};

const shiftTone = (value: string) => {
  if (!value || value === 'Set') return 'border-ink/15 bg-[#f6f6f6] text-ink/55';
  if (value === 'OFF' || value === 'N/A') return 'border-[#b45309] bg-[#d97706] text-white';
  if (value === 'TIME OFF' || value === 'REQ VAC') return 'border-[#92400e] bg-[#b45309] text-white';
  if (value === '5AM–12PM' || value === '5AM-12PM') return 'border-transparent bg-[#0ea5e9] text-white';
  if (value === '6AM–12PM') return 'border-[#1e40af] bg-[#1d4ed8] text-white';
  if (value === '7AM–12PM') return 'border-[#6d28d9] bg-[#7c3aed] text-white';
  if (value === '10:30am - 6:30pm (c)') return 'border-[#ea580c] bg-[#f97316] text-white';
  if (value === 'AM (6:00AM–2:00PM)' || value === 'AM (3:30AM–11:30AM)') {
    return 'border-[#166534] bg-[#15803d] text-white';
  }
  if (value === 'AM (6:15AM–2:15PM)') return 'border-[#16a34a] bg-[#22c55e] text-white';
  if (value === 'PM (2:00PM–10:00PM)' || value.startsWith('Midday')) {
    return 'border-[#1e40af] bg-[#1d4ed8] text-white';
  }
  if (value === 'PM (2:15PM–10:15PM)') return 'border-[#2563eb] bg-[#60a5fa] text-white';
  if (value.startsWith('Audit') || value.startsWith('Crew')) {
    return 'border-[#991b1b] bg-[#b91c1c] text-white';
  }
  if (value.startsWith('PM')) return 'border-[#6d28d9] bg-[#7c3aed] text-white';
  if (value.startsWith('AM') || value === '8AM–4:30PM') {
    return 'border-[#166534] bg-[#15803d] text-white';
  }
  if (/\b(?:4|5|6|7|8|9|10|11|12)(?::\d{2})?\s*(?:pm|PM)\b/.test(value)) {
    return 'border-[#991b1b] bg-[#b91c1c] text-white';
  }
  return 'border-ink/15 bg-white text-ink';
};

type ScheduleRow = {
  employeeId: number;
  employeeName: string;
  sectionName: string;
  assignmentsByDate: Record<string, Assignment>;
};

type SectionSchedule = {
  sectionName: string;
  rows: ScheduleRow[];
};

export function SchedulesSection() {
  const [weeks, setWeeks] = useState<WeekSummary[]>([]);
  const [selectedWeekId, setSelectedWeekId] = useState('');
  const [selectedWeek, setSelectedWeek] = useState<Week | null>(null);
  const [error, setError] = useState('');
  const [savingAssignmentId, setSavingAssignmentId] = useState<number | null>(null);
  const [savingAircrewKey, setSavingAircrewKey] = useState('');
  const [aircrewDraftTimes, setAircrewDraftTimes] = useState<Record<string, string>>({});

  useEffect(() => {
    api
      .get<WeekSummary[]>('/api/schedules/weeks')
      .then((data) => {
        setWeeks(data);
        setSelectedWeekId((current) => current || (data[0] ? String(data[0].id) : ''));
      })
      .catch((requestError: Error) => setError(requestError.message));
  }, []);

  const loadSelectedWeek = useCallback(async () => {
    if (!selectedWeekId) {
      setSelectedWeek(null);
      return;
    }

    const week = await api.get<Week>(`/api/schedules/weeks/${selectedWeekId}`);
    setSelectedWeek(week);
  }, [selectedWeekId]);

  useEffect(() => {
    loadSelectedWeek().catch((requestError: Error) => setError(requestError.message));
  }, [loadSelectedWeek]);

  const weekDates = useMemo(() => {
    if (!selectedWeek) return [];
    return Array.from({ length: 7 }, (_value, index) => addDays(selectedWeek.startDate, index));
  }, [selectedWeek]);

  const scheduleRows = useMemo(() => {
    if (!selectedWeek) return [];

    const rowsByEmployee = new Map<number, ScheduleRow>();
    selectedWeek.assignments.forEach((assignment) => {
      const employee = assignment.employee;
      if (!employee) return;

      const current = rowsByEmployee.get(employee.id) ?? {
        employeeId: employee.id,
        employeeName: employee.name,
        sectionName: employee.section?.name ?? 'Unassigned',
        assignmentsByDate: {},
      };
      current.assignmentsByDate[assignment.date] = assignment;
      rowsByEmployee.set(employee.id, current);
    });

    return Array.from(rowsByEmployee.values());
  }, [selectedWeek]);

  const sectionSchedules = useMemo<SectionSchedule[]>(() => {
    const groups = scheduleRows.reduce<Record<string, ScheduleRow[]>>((currentGroups, row) => {
      currentGroups[row.sectionName] = currentGroups[row.sectionName] ?? [];
      currentGroups[row.sectionName].push(row);
      return currentGroups;
    }, {});

    const knownSections = sectionOrder
      .filter((sectionName) => groups[sectionName])
      .map((sectionName) => ({
        sectionName,
        rows: groups[sectionName],
      }));

    const extraSections = Object.entries(groups)
      .filter(([sectionName]) => !sectionOrder.includes(sectionName))
      .sort(([left], [right]) => left.localeCompare(right))
      .map(([sectionName, rows]) => ({ sectionName, rows }));

    return [...knownSections, ...extraSections];
  }, [scheduleRows]);

  const selectedWeekIndex = weeks.findIndex((week) => String(week.id) === selectedWeekId);
  const previousWeek = selectedWeekIndex >= 0 ? weeks[selectedWeekIndex + 1] : undefined;
  const nextWeek = selectedWeekIndex > 0 ? weeks[selectedWeekIndex - 1] : undefined;

  const shiftOptionsBySection = useMemo(() => {
    const optionsBySection: Record<string, string[]> = {};
    sectionSchedules.forEach((section) => {
      const seen = new Set<string>();
      const options: string[] = [];
      const addOption = (value: string) => {
        if (!seen.has(value)) {
          seen.add(value);
          options.push(value);
        }
      };

      (sectionShiftOptions[section.sectionName] ?? neutralShiftOptions).forEach(addOption);
      section.rows.forEach((row) => {
        Object.values(row.assignmentsByDate).forEach((assignment) => addOption(assignment.value));
      });
      optionsBySection[section.sectionName] = options;
    });

    return optionsBySection;
  }, [sectionSchedules]);

  const updateAssignmentValue = async (assignment: Assignment, value: string) => {
    if (!selectedWeek || assignment.value === value) return;

    setError('');
    setSavingAssignmentId(assignment.id);
    try {
      const updatedAssignment = await api.patch<Assignment>(
        `/api/schedules/assignments/${assignment.id}`,
        { value },
      );
      setSelectedWeek((currentWeek) => {
        if (!currentWeek) return currentWeek;
        return {
          ...currentWeek,
          assignments: currentWeek.assignments.map((currentAssignment) =>
            currentAssignment.id === updatedAssignment.id ? updatedAssignment : currentAssignment,
          ),
        };
      });
    } catch (requestError) {
      setError((requestError as Error).message);
    } finally {
      setSavingAssignmentId(null);
    }
  };

  const updateAircrewCell = (
    carrier: string,
    date: string,
    times: string[],
  ) => {
    setSelectedWeek((currentWeek) => {
      if (!currentWeek) return currentWeek;

      const existingRows = currentWeek.aircrewArrivals ?? [];
      const hasCarrier = existingRows.some((row) => row.carrier === carrier);
      const nextRows = (hasCarrier ? existingRows : [...existingRows, { carrier, timesByDate: {} }]).map((row) => {
        if (row.carrier !== carrier) return row;
        return {
          ...row,
          timesByDate: {
            ...row.timesByDate,
            [date]: times,
          },
        };
      });

      return {
        ...currentWeek,
        aircrewArrivals: nextRows,
      };
    });
  };

  const postAircrewTime = async (
    carrier: string,
    date: string,
    action: 'add' | 'remove',
    time: string,
  ) => {
    if (!selectedWeek) return;
    const normalizedTime = normalizeTimeInput(time);
    if (!normalizedTime) {
      setError('Enter a valid aircrew time.');
      return;
    }

    const key = `${carrier}:${date}`;
    setError('');
    setSavingAircrewKey(key);
    try {
      const data = await api.post<{ carrier: string; cells: Record<string, string[]> }>(
        `/api/schedules/weeks/${selectedWeek.id}/aircrew/arrival`,
        {
          carrier,
          date,
          action,
          time: normalizedTime,
        },
      );
      updateAircrewCell(data.carrier, date, data.cells[date] ?? []);
      if (action === 'add') {
        setAircrewDraftTimes((current) => ({ ...current, [key]: normalizedTime }));
      }
    } catch (requestError) {
      setError((requestError as Error).message);
    } finally {
      setSavingAircrewKey('');
    }
  };

  const addAircrewCarrier = async () => {
    const rawName = window.prompt('Enter the new airline crew name:');
    if (rawName === null) return;
    const carrier = rawName.trim();
    if (!carrier) {
      setError('Carrier name is required.');
      return;
    }

    try {
      const data = await api.post<{ carriers: string[] }>('/api/schedules/aircrew/carriers', {
        carrier,
      });
      setSelectedWeek((currentWeek) => {
        if (!currentWeek) return currentWeek;
        const existing = new Set(currentWeek.aircrewArrivals.map((row) => row.carrier));
        const rows = [...currentWeek.aircrewArrivals];
        data.carriers.forEach((carrierName) => {
          if (!existing.has(carrierName)) {
            rows.push({ carrier: carrierName, timesByDate: {} });
          }
        });
        return {
          ...currentWeek,
          aircrewArrivals: rows.sort((left, right) => left.carrier.localeCompare(right.carrier)),
        };
      });
    } catch (requestError) {
      setError((requestError as Error).message);
    }
  };

  const removeAircrewCarrier = async () => {
    const rawName = window.prompt('Enter the carrier name to remove:');
    if (rawName === null) return;
    const carrier = rawName.trim();
    if (!carrier) {
      setError('Carrier name is required.');
      return;
    }
    if (!window.confirm(`Remove ${carrier} and all of its saved arrival times?`)) {
      return;
    }

    try {
      await api.delete(`/api/schedules/aircrew/carriers/${encodeURIComponent(carrier)}`);
      setSelectedWeek((currentWeek) => {
        if (!currentWeek) return currentWeek;
        return {
          ...currentWeek,
          aircrewArrivals: currentWeek.aircrewArrivals.filter((row) => row.carrier !== carrier),
        };
      });
    } catch (requestError) {
      setError((requestError as Error).message);
    }
  };

  const renderAircrewPanel = (rows: AircrewCarrierArrivals[]) => (
    <Card className="overflow-hidden p-0" key="aircrew-arrivals">
      <div className="flex flex-col gap-3 border-b border-ink/10 bg-white/80 px-5 py-4 md:flex-row md:items-center md:justify-between">
        <div>
          <h3 className="text-lg font-bold text-ink">Aircrew Arrivals</h3>
          <p className="mt-1 text-sm text-ink/55">
            Arrival times used by shuttle scheduling for this week.
          </p>
        </div>
        <div className="flex flex-wrap gap-2">
          <Button tone="secondary" type="button" onClick={addAircrewCarrier}>
            Add Carrier
          </Button>
          <Button tone="secondary" type="button" onClick={removeAircrewCarrier}>
            Remove Carrier
          </Button>
        </div>
      </div>

      <div className="overflow-x-auto bg-white">
        <div className="min-w-[1120px]">
          <div
            className="grid border-b border-ink/10 bg-mist text-xs font-bold uppercase tracking-[0.14em] text-ink/50"
            style={{ gridTemplateColumns: '180px repeat(7, minmax(132px, 1fr))' }}
          >
            <div className="sticky left-0 z-20 bg-mist px-4 py-3">Carrier</div>
            {weekDates.map((dateText) => (
              <div className="border-l border-ink/10 px-3 py-3 text-center" key={dateText}>
                {formatDate(dateText)}
              </div>
            ))}
          </div>

          {rows.map((row) => (
            <div
              className="grid border-b border-ink/10 last:border-b-0"
              key={row.carrier}
              style={{ gridTemplateColumns: '180px repeat(7, minmax(132px, 1fr))' }}
            >
              <div className="sticky left-0 z-10 flex min-h-20 items-center bg-white px-4 py-2 text-sm font-semibold text-ink">
                {row.carrier}
              </div>
              {weekDates.map((dateText) => {
                const key = `${row.carrier}:${dateText}`;
                const times = row.timesByDate[dateText] ?? [];
                const draftTime = aircrewDraftTimes[key] ?? '18:00';
                const saving = savingAircrewKey === key;

                return (
                  <div
                    className="flex min-h-20 flex-col gap-2 border-l border-ink/10 px-2 py-2"
                    key={dateText}
                  >
                    <div className="flex min-h-7 flex-wrap gap-1">
                      {times.map((time) => (
                        <span
                          className="inline-flex items-center gap-1 rounded-full bg-ink px-2 py-1 text-xs font-semibold text-white"
                          key={time}
                        >
                          {formatAircrewTime(time)}
                          <button
                            aria-label={`Remove ${formatAircrewTime(time)}`}
                            className="ml-1 rounded-full px-1 text-white/75 transition hover:bg-white/15 hover:text-white"
                            disabled={saving}
                            onClick={() => postAircrewTime(row.carrier, dateText, 'remove', time)}
                            type="button"
                          >
                            x
                          </button>
                        </span>
                      ))}
                    </div>
                    <div className="flex gap-2">
                      <Input
                        aria-label={`${row.carrier} ${formatDate(dateText)} arrival time`}
                        className="h-9 rounded-xl px-2 py-1 text-xs"
                        disabled={saving}
                        onChange={(event) =>
                          setAircrewDraftTimes((current) => ({
                            ...current,
                            [key]: event.target.value,
                          }))
                        }
                        type="time"
                        value={draftTime}
                      />
                      <Button
                        className="h-9 px-3 py-1 text-xs"
                        disabled={saving}
                        onClick={() => postAircrewTime(row.carrier, dateText, 'add', draftTime)}
                        type="button"
                      >
                        +
                      </Button>
                    </div>
                  </div>
                );
              })}
            </div>
          ))}
        </div>
      </div>
    </Card>
  );

  return (
    <div className="space-y-6">
      <Card className="border-ink/10 bg-white text-ink">
        <div className="flex flex-col gap-4 md:flex-row md:items-end md:justify-between">
          <div>
            <p className="text-sm uppercase tracking-[0.24em] text-ink/45">Planning</p>
            <h2 className="mt-3 font-display text-3xl">
              {selectedWeek ? `Schedule: ${formatRange(selectedWeek.startDate)}` : 'Schedules'}
            </h2>
            <p className="mt-2 max-w-xl text-sm text-ink/65">
              Weekly roster tables arranged like the original scheduler, wrapped in the new panel style.
            </p>
          </div>
          <div className="flex flex-wrap gap-2">
            <Badge tone="warning">{weeks.length} weeks</Badge>
            {selectedWeek ? <Badge tone="success">{selectedWeek.assignments.length} assignments</Badge> : null}
          </div>
        </div>
      </Card>

      <Card>
        <div className="flex flex-col gap-4 lg:flex-row lg:items-end lg:justify-between">
          <div>
            <h3 className="text-lg font-bold text-ink">Schedule controls</h3>
            <p className="mt-1 text-sm text-ink/60">
              Move between generated weeks and use the same primary actions as the old schedule page.
            </p>
          </div>
          <div className="grid gap-3 md:grid-cols-[auto_minmax(220px,320px)_auto] md:items-end">
            <Button
              disabled={!previousWeek}
              tone="secondary"
              type="button"
              onClick={() => previousWeek && setSelectedWeekId(String(previousWeek.id))}
            >
              &lt; Previous
            </Button>
            <label className="space-y-2">
              <span className="block text-sm font-semibold text-ink/80">Week</span>
              <Select
                value={selectedWeekId}
                onChange={(event) => setSelectedWeekId(event.target.value)}
              >
                {weeks.map((week) => (
                  <option key={week.id} value={week.id}>
                    {week.startDate} ({week.assignmentCount})
                  </option>
                ))}
              </Select>
            </label>
            <Button
              disabled={!nextWeek}
              tone="secondary"
              type="button"
              onClick={() => nextWeek && setSelectedWeekId(String(nextWeek.id))}
            >
              Next &gt;
            </Button>
          </div>
        </div>
        {error ? <p className="mt-4 text-sm font-semibold text-clay">{error}</p> : null}
        <div className="mt-5 flex flex-wrap gap-2">
          <Button type="button">Generate New Schedule</Button>
          <Button tone="secondary" type="button">Export to Excel</Button>
          <Button tone="secondary" type="button">Export Shuttle/Air Crew</Button>
          <Button tone="secondary" type="button">Manager Meal List</Button>
          <Button
            tone="secondary"
            type="button"
            onClick={() =>
              loadSelectedWeek().catch((requestError: Error) => setError(requestError.message))
            }
          >
            Refresh
          </Button>
          <Button tone="danger" type="button">Delete This Schedule</Button>
        </div>
      </Card>

      <div className="space-y-6">
        {sectionSchedules.map((section) => (
          <div className="space-y-6" key={section.sectionName}>
            <Card className="overflow-hidden p-0">
              <div className="flex flex-col gap-2 border-b border-ink/10 bg-white/80 px-5 py-4 md:flex-row md:items-center md:justify-between">
                <div>
                  <h3 className="text-lg font-bold text-ink">{section.sectionName}</h3>
                  <p className="mt-1 text-sm text-ink/55">{section.rows.length} employees</p>
                </div>
                <div className="flex flex-wrap gap-2">
                  <Badge>{section.rows.reduce((count, row) => count + Object.keys(row.assignmentsByDate).length, 0)} shifts</Badge>
                </div>
              </div>

              <div className="overflow-x-auto bg-white">
                <div className="min-w-[980px]">
                  <div
                    className="grid border-b border-ink/10 bg-mist text-xs font-bold uppercase tracking-[0.14em] text-ink/50"
                    style={{ gridTemplateColumns: '180px repeat(7, minmax(112px, 1fr))' }}
                  >
                    <div className="sticky left-0 z-20 bg-mist px-4 py-3">Employees</div>
                    {weekDates.map((dateText) => (
                      <div className="border-l border-ink/10 px-3 py-3 text-center" key={dateText}>
                        {formatDate(dateText)}
                      </div>
                    ))}
                  </div>

                  {section.rows.map((row) => (
                    <div
                      className="grid border-b border-ink/10 last:border-b-0"
                      key={row.employeeId}
                      style={{ gridTemplateColumns: '180px repeat(7, minmax(112px, 1fr))' }}
                    >
                      <div className="sticky left-0 z-10 flex min-h-12 items-center bg-white px-4 py-2 text-sm font-semibold text-ink">
                        {row.employeeName}
                      </div>
                      {weekDates.map((dateText) => {
                        const assignment = row.assignmentsByDate[dateText];
                        return (
                          <div
                            className="flex min-h-12 items-center justify-center border-l border-ink/10 px-2 py-2"
                            key={dateText}
                          >
                            {assignment ? (
                              <select
                                aria-label={`${row.employeeName} ${formatDate(dateText)}`}
                                className={`h-8 w-full rounded-md border px-2 text-center text-xs font-semibold outline-none transition focus:ring-2 focus:ring-gold/50 disabled:opacity-70 ${shiftTone(
                                  assignment.value,
                                )}`}
                                disabled={savingAssignmentId === assignment.id}
                                onChange={(event) => updateAssignmentValue(assignment, event.target.value)}
                                value={assignment.value}
                              >
                                {(shiftOptionsBySection[section.sectionName] ?? [assignment.value]).map((option) => (
                                  <option key={option} value={option}>
                                    {formatShiftLabel(option)}
                                  </option>
                                ))}
                              </select>
                            ) : (
                              <span className="text-sm text-ink/25">-</span>
                            )}
                          </div>
                        );
                      })}
                    </div>
                  ))}

                  {!section.rows.length ? (
                    <div className="px-4 py-8 text-center text-sm text-ink/60">
                      No employees in this section.
                    </div>
                  ) : null}
                </div>
              </div>
            </Card>
            {section.sectionName === 'Shuttle' && selectedWeek
              ? renderAircrewPanel(selectedWeek.aircrewArrivals ?? [])
              : null}
          </div>
        ))}

        {selectedWeek && !sectionSchedules.length ? (
          <Card>
            <p className="py-8 text-center text-sm text-ink/60">No assignments found for this week.</p>
          </Card>
        ) : null}
      </div>
    </div>
  );
}
