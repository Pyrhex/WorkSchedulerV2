import { type FormEvent, useCallback, useEffect, useMemo, useState } from 'react';

import { api } from '../../api/client';
import { Badge, Button, Card, Input, Select } from '../../components/ui';
import type { Employee, Section } from '../../types/api';

const emptyForm = {
  name: '',
  sectionId: '',
  firstName: '',
  lastName: '',
  preferredShift: '',
  maxShiftsPerWeek: '',
  temporary: 'false',
};

export function EmployeesSection() {
  const [employees, setEmployees] = useState<Employee[]>([]);
  const [sections, setSections] = useState<Section[]>([]);
  const [editingId, setEditingId] = useState<number | null>(null);
  const [form, setForm] = useState(emptyForm);
  const [query, setQuery] = useState('');
  const [error, setError] = useState('');

  const loadData = useCallback(async () => {
    const [employeeData, sectionData] = await Promise.all([
      api.get<Employee[]>('/api/employees'),
      api.get<Section[]>('/api/sections'),
    ]);
    setEmployees(employeeData);
    setSections(sectionData);
  }, []);

  useEffect(() => {
    loadData().catch((requestError: Error) => setError(requestError.message));
  }, [loadData]);

  const groupedEmployees = useMemo(() => {
    const normalizedQuery = query.trim().toLowerCase();
    const visible = normalizedQuery
      ? employees.filter((employee) => employee.name.toLowerCase().includes(normalizedQuery))
      : employees;

    return visible.reduce<Record<string, Employee[]>>((groups, employee) => {
      const key = employee.section?.name ?? 'Unassigned';
      groups[key] = groups[key] ?? [];
      groups[key].push(employee);
      return groups;
    }, {});
  }, [employees, query]);

  const resetForm = () => {
    setEditingId(null);
    setForm(emptyForm);
  };

  const handleSubmit = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    setError('');

    const payload = {
      name: form.name,
      sectionId: Number(form.sectionId),
      firstName: form.firstName || null,
      lastName: form.lastName || null,
      preferredShift: form.preferredShift || null,
      maxShiftsPerWeek: form.maxShiftsPerWeek ? Number(form.maxShiftsPerWeek) : null,
      temporary: form.temporary === 'true',
    };

    try {
      if (editingId) {
        await api.put(`/api/employees/${editingId}`, payload);
      } else {
        await api.post('/api/employees', payload);
      }
      resetForm();
      await loadData();
    } catch (requestError) {
      setError((requestError as Error).message);
    }
  };

  const startEdit = (employee: Employee) => {
    setEditingId(employee.id);
    setForm({
      name: employee.name,
      sectionId: String(employee.sectionId),
      firstName: employee.firstName ?? '',
      lastName: employee.lastName ?? '',
      preferredShift: employee.preferredShift ?? '',
      maxShiftsPerWeek: employee.maxShiftsPerWeek ? String(employee.maxShiftsPerWeek) : '',
      temporary: String(employee.temporary),
    });
  };

  const handleDelete = async (id: number) => {
    if (!window.confirm('Delete this employee?')) return;

    try {
      await api.delete(`/api/employees/${id}`);
      await loadData();
      if (editingId === id) resetForm();
    } catch (requestError) {
      setError((requestError as Error).message);
    }
  };

  return (
    <div className="space-y-6">
      <Card className="border-ink/10 bg-white text-ink">
        <div className="flex flex-col gap-4 md:flex-row md:items-end md:justify-between">
          <div>
            <p className="text-sm uppercase tracking-[0.24em] text-ink/45">Roster</p>
            <h2 className="mt-3 font-display text-3xl">Employees</h2>
            <p className="mt-2 max-w-xl text-sm text-ink/65">
              Manage the same employee and role data currently maintained in Flask.
            </p>
          </div>
          <Badge tone="warning">{employees.length} employees</Badge>
        </div>
      </Card>

      <Card>
        <form className="grid gap-3 md:grid-cols-3" onSubmit={handleSubmit}>
          <label className="space-y-2 md:col-span-2">
            <span className="block text-sm font-semibold text-ink/80">Name</span>
            <Input
              required
              value={form.name}
              onChange={(event) => setForm((current) => ({ ...current, name: event.target.value }))}
            />
          </label>
          <label className="space-y-2">
            <span className="block text-sm font-semibold text-ink/80">Primary section</span>
            <Select
              required
              value={form.sectionId}
              onChange={(event) => setForm((current) => ({ ...current, sectionId: event.target.value }))}
            >
              <option value="">Choose section</option>
              {sections.map((section) => (
                <option key={section.id} value={section.id}>
                  {section.name}
                </option>
              ))}
            </Select>
          </label>
          <label className="space-y-2">
            <span className="block text-sm font-semibold text-ink/80">First name</span>
            <Input
              value={form.firstName}
              onChange={(event) => setForm((current) => ({ ...current, firstName: event.target.value }))}
            />
          </label>
          <label className="space-y-2">
            <span className="block text-sm font-semibold text-ink/80">Last name</span>
            <Input
              value={form.lastName}
              onChange={(event) => setForm((current) => ({ ...current, lastName: event.target.value }))}
            />
          </label>
          <label className="space-y-2">
            <span className="block text-sm font-semibold text-ink/80">Preferred shift</span>
            <Input
              value={form.preferredShift}
              onChange={(event) => setForm((current) => ({ ...current, preferredShift: event.target.value }))}
            />
          </label>
          <label className="space-y-2">
            <span className="block text-sm font-semibold text-ink/80">Max shifts/week</span>
            <Input
              min="0"
              type="number"
              value={form.maxShiftsPerWeek}
              onChange={(event) =>
                setForm((current) => ({ ...current, maxShiftsPerWeek: event.target.value }))
              }
            />
          </label>
          <label className="space-y-2">
            <span className="block text-sm font-semibold text-ink/80">Temporary</span>
            <Select
              value={form.temporary}
              onChange={(event) => setForm((current) => ({ ...current, temporary: event.target.value }))}
            >
              <option value="false">No</option>
              <option value="true">Yes</option>
            </Select>
          </label>
          <div className="flex flex-wrap items-end gap-3 md:col-span-3">
            <Button type="submit">{editingId ? 'Update employee' : 'Add employee'}</Button>
            {editingId ? (
              <Button tone="secondary" type="button" onClick={resetForm}>
                Cancel
              </Button>
            ) : null}
          </div>
        </form>
        {error ? <p className="mt-4 text-sm font-semibold text-clay">{error}</p> : null}
      </Card>

      <Card>
        <div className="mb-5 flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
          <h3 className="text-lg font-bold text-ink">Current roster</h3>
          <Input
            className="md:max-w-xs"
            placeholder="Search employees"
            value={query}
            onChange={(event) => setQuery(event.target.value)}
          />
        </div>
        <div className="space-y-5">
          {Object.entries(groupedEmployees).map(([sectionName, sectionEmployees]) => (
            <div key={sectionName}>
              <h4 className="mb-3 text-sm font-bold uppercase tracking-[0.18em] text-ink/45">
                {sectionName}
              </h4>
              <div className="overflow-hidden rounded-2xl border border-ink/10 bg-white">
                <table className="w-full border-collapse text-left text-sm">
                  <thead className="bg-mist text-xs uppercase tracking-[0.14em] text-ink/50">
                    <tr>
                      <th className="px-4 py-3">Employee</th>
                      <th className="px-4 py-3">Preferred</th>
                      <th className="px-4 py-3">Max</th>
                      <th className="px-4 py-3">Actions</th>
                    </tr>
                  </thead>
                  <tbody>
                    {sectionEmployees.map((employee) => (
                      <tr className="border-t border-ink/10" key={employee.id}>
                        <td className="px-4 py-3 font-semibold">{employee.name}</td>
                        <td className="px-4 py-3 text-ink/70">{employee.preferredShift ?? '-'}</td>
                        <td className="px-4 py-3 text-ink/70">{employee.maxShiftsPerWeek ?? '-'}</td>
                        <td className="flex flex-wrap gap-2 px-4 py-3">
                          <Button tone="secondary" type="button" onClick={() => startEdit(employee)}>
                            Edit
                          </Button>
                          <Button tone="danger" type="button" onClick={() => handleDelete(employee.id)}>
                            Delete
                          </Button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          ))}
        </div>
      </Card>
    </div>
  );
}
