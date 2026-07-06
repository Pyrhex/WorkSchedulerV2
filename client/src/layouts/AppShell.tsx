import type { PropsWithChildren } from 'react';

import { NavLink } from 'react-router-dom';

import schedulerLogo from '../assets/scheduler-logo.svg';

const navItems = [
  { to: '/admin/schedules', label: 'Schedules', description: 'Weekly roster' },
  { to: '/admin/employees', label: 'Employees', description: 'Roster and roles' },
  { to: '/admin/time-off', label: 'Time off', description: 'Requests' },
];

const pendingNavItems = [
  'Meals',
  'Templates',
  'Aircrew',
  'Occupancy',
];

export function AppShell({ children }: PropsWithChildren) {
  return (
    <div className="min-h-screen px-3 pb-8 pt-4 text-ink md:px-5 lg:px-6">
      <nav className="sticky top-3 z-40 mx-auto mb-6 max-w-[1800px] rounded-[28px] border border-white/70 bg-white/90 px-4 py-3 shadow-panel backdrop-blur">
        <div className="flex flex-col gap-3 lg:flex-row lg:items-center lg:justify-between">
          <div className="flex items-center gap-3">
            <div className="flex h-12 w-12 shrink-0 items-center justify-center rounded-2xl bg-ink p-2">
              <img alt="Work Scheduler logo" className="h-full w-full object-contain" src={schedulerLogo} />
            </div>
            <div>
              <div className="font-display text-2xl leading-none text-ink">Work Scheduler</div>
              <div className="mt-1 text-xs font-semibold uppercase tracking-[0.18em] text-ink/45">
                Operations Planning
              </div>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            {navItems.map((item) => (
              <NavLink
                className={({ isActive }) =>
                  `group rounded-2xl border px-4 py-2 transition ${
                    isActive
                      ? 'border-ink bg-ink text-white'
                      : 'border-ink/10 bg-mist text-ink hover:border-fairway hover:text-fairway'
                  }`
                }
                key={item.to}
                to={item.to}
              >
                <span className="block text-sm font-bold leading-tight">{item.label}</span>
                <span className="block text-xs leading-tight opacity-65">{item.description}</span>
              </NavLink>
            ))}
          </div>
        </div>

        <div className="mt-3 flex flex-wrap items-center gap-2 border-t border-ink/10 pt-3">
          <span className="text-xs font-bold uppercase tracking-[0.18em] text-ink/40">Next</span>
          {pendingNavItems.map((item) => (
            <button
              className="rounded-full border border-dashed border-ink/15 px-3 py-1 text-xs font-semibold text-ink/45"
              disabled
              key={item}
              type="button"
            >
              {item}
            </button>
          ))}
        </div>
      </nav>

      <main className="mx-auto max-w-[1800px]">{children}</main>
    </div>
  );
}
