import { type FormEvent, useCallback, useEffect, useState } from 'react';

import { api } from '../../api/client';
import { Badge, Button, Card, Input, Select } from '../../components/ui';
import type { TimeOffRequest } from '../../types/api';

const emptyForm = {
  employeeName: '',
  role: '',
  startDate: '',
  endDate: '',
  approved: 'false',
  vacation: 'false',
};

export function TimeOffSection() {
  const [requests, setRequests] = useState<TimeOffRequest[]>([]);
  const [form, setForm] = useState(emptyForm);
  const [error, setError] = useState('');

  const loadRequests = useCallback(async () => {
    const data = await api.get<TimeOffRequest[]>('/api/time-off');
    setRequests(data);
  }, []);

  useEffect(() => {
    loadRequests().catch((requestError: Error) => setError(requestError.message));
  }, [loadRequests]);

  const handleSubmit = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    setError('');

    try {
      await api.post('/api/time-off', {
        employeeName: form.employeeName,
        role: form.role,
        startDate: form.startDate,
        endDate: form.endDate,
        approved: form.approved === 'true',
        vacation: form.vacation === 'true',
      });
      setForm(emptyForm);
      await loadRequests();
    } catch (requestError) {
      setError((requestError as Error).message);
    }
  };

  const handleDelete = async (id: number) => {
    if (!window.confirm('Delete this time off request?')) return;

    try {
      await api.delete(`/api/time-off/${id}`);
      await loadRequests();
    } catch (requestError) {
      setError((requestError as Error).message);
    }
  };

  return (
    <div className="space-y-6">
      <Card className="border-ink/10 bg-white text-ink">
        <div className="flex flex-col gap-4 md:flex-row md:items-end md:justify-between">
          <div>
            <p className="text-sm uppercase tracking-[0.24em] text-ink/45">Requests</p>
            <h2 className="mt-3 font-display text-3xl">Time off</h2>
            <p className="mt-2 max-w-xl text-sm text-ink/65">
              A first pass at the React version of the time-off workflow.
            </p>
          </div>
          <Badge tone="warning">{requests.length} requests</Badge>
        </div>
      </Card>

      <Card>
        <form className="grid gap-3 md:grid-cols-2" onSubmit={handleSubmit}>
          <label className="space-y-2">
            <span className="block text-sm font-semibold text-ink/80">Employee</span>
            <Input
              required
              value={form.employeeName}
              onChange={(event) =>
                setForm((current) => ({ ...current, employeeName: event.target.value }))
              }
            />
          </label>
          <label className="space-y-2">
            <span className="block text-sm font-semibold text-ink/80">Role</span>
            <Input
              required
              value={form.role}
              onChange={(event) => setForm((current) => ({ ...current, role: event.target.value }))}
            />
          </label>
          <label className="space-y-2">
            <span className="block text-sm font-semibold text-ink/80">From</span>
            <Input
              required
              type="date"
              value={form.startDate}
              onChange={(event) =>
                setForm((current) => ({ ...current, startDate: event.target.value }))
              }
            />
          </label>
          <label className="space-y-2">
            <span className="block text-sm font-semibold text-ink/80">To</span>
            <Input
              required
              type="date"
              value={form.endDate}
              onChange={(event) => setForm((current) => ({ ...current, endDate: event.target.value }))}
            />
          </label>
          <label className="space-y-2">
            <span className="block text-sm font-semibold text-ink/80">Approved</span>
            <Select
              value={form.approved}
              onChange={(event) => setForm((current) => ({ ...current, approved: event.target.value }))}
            >
              <option value="false">No</option>
              <option value="true">Yes</option>
            </Select>
          </label>
          <label className="space-y-2">
            <span className="block text-sm font-semibold text-ink/80">Vacation</span>
            <Select
              value={form.vacation}
              onChange={(event) => setForm((current) => ({ ...current, vacation: event.target.value }))}
            >
              <option value="false">No</option>
              <option value="true">Yes</option>
            </Select>
          </label>
          <div className="md:col-span-2">
            <Button type="submit">Add request</Button>
          </div>
        </form>
        {error ? <p className="mt-4 text-sm font-semibold text-clay">{error}</p> : null}
      </Card>

      <Card>
        <h3 className="text-lg font-bold text-ink">Current requests</h3>
        <div className="mt-5 overflow-hidden rounded-2xl border border-ink/10 bg-white">
          <table className="w-full border-collapse text-left text-sm">
            <thead className="bg-mist text-xs uppercase tracking-[0.14em] text-ink/50">
              <tr>
                <th className="px-4 py-3">Employee</th>
                <th className="px-4 py-3">Role</th>
                <th className="px-4 py-3">Dates</th>
                <th className="px-4 py-3">Status</th>
                <th className="px-4 py-3">Actions</th>
              </tr>
            </thead>
            <tbody>
              {requests.map((request) => (
                <tr className="border-t border-ink/10" key={request.id}>
                  <td className="px-4 py-3 font-semibold">{request.employeeName}</td>
                  <td className="px-4 py-3 text-ink/70">{request.role}</td>
                  <td className="px-4 py-3 text-ink/70">
                    {new Date(request.startDate).toLocaleDateString()} to{' '}
                    {new Date(request.endDate).toLocaleDateString()}
                  </td>
                  <td className="px-4 py-3">
                    <Badge tone={request.approved ? 'success' : 'default'}>
                      {request.approved ? 'Approved' : 'Pending'}
                    </Badge>
                  </td>
                  <td className="px-4 py-3">
                    <Button tone="danger" type="button" onClick={() => handleDelete(request.id)}>
                      Delete
                    </Button>
                  </td>
                </tr>
              ))}
              {!requests.length ? (
                <tr>
                  <td className="px-4 py-8 text-center text-ink/60" colSpan={5}>
                    No time off requests found yet.
                  </td>
                </tr>
              ) : null}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}
