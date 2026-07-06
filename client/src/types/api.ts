export type Section = {
  id: number;
  name: string;
  requiredPerDay: number | null;
};

export type EmployeeRole = {
  id: number;
  employeeId: number;
  sectionId: number;
  section: Section;
};

export type Employee = {
  id: number;
  name: string;
  sectionId: number;
  section: Section;
  availability: string | null;
  preferredShift: string | null;
  seniority: number | null;
  preferredShiftsPerWeek: number | null;
  maxShiftsPerWeek: number | null;
  sortOrder: number | null;
  firstName: string | null;
  lastName: string | null;
  temporary: boolean;
  roles: EmployeeRole[];
};

export type Week = {
  id: number;
  startDate: string;
  assignments: Assignment[];
  aircrewArrivals: AircrewCarrierArrivals[];
};

export type AircrewCarrierArrivals = {
  carrier: string;
  timesByDate: Record<string, string[]>;
};

export type WeekSummary = {
  id: number;
  startDate: string;
  assignmentCount: number;
};

export type Assignment = {
  id: number;
  weekId: number;
  employeeId: number;
  date: string;
  value: string;
  dismissedTimeOff: boolean;
  employee?: Employee;
};

export type TimeOffRequest = {
  id: number;
  employeeName: string;
  role: string;
  startDate: string;
  endDate: string;
  approved: boolean;
  vacation: boolean;
};
