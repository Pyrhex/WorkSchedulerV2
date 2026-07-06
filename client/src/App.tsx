import { Navigate, Route, Routes } from 'react-router-dom';

import { EmployeesSection } from './features/employees/EmployeesSection';
import { SchedulesSection } from './features/schedules/SchedulesSection';
import { TimeOffSection } from './features/timeoff/TimeOffSection';
import { AppShell } from './layouts/AppShell';

function SchedulesPage() {
  return (
    <AppShell>
      <SchedulesSection />
    </AppShell>
  );
}

function EmployeesPage() {
  return (
    <AppShell>
      <EmployeesSection />
    </AppShell>
  );
}

function TimeOffPage() {
  return (
    <AppShell>
      <TimeOffSection />
    </AppShell>
  );
}

export default function App() {
  return (
    <Routes>
      <Route path="/" element={<Navigate replace to="/admin/schedules" />} />
      <Route path="/admin" element={<Navigate replace to="/admin/schedules" />} />
      <Route path="/admin/schedules" element={<SchedulesPage />} />
      <Route path="/admin/employees" element={<EmployeesPage />} />
      <Route path="/admin/time-off" element={<TimeOffPage />} />
      <Route path="*" element={<Navigate replace to="/admin/schedules" />} />
    </Routes>
  );
}
