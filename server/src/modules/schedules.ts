import { Router } from 'express';
import { z } from 'zod';

import { HttpError, asyncHandler } from '../lib/http';
import { prisma } from '../lib/prisma';

export const schedulesRouter = Router();

const defaultAircrewCarriers = ['Aeromexico', 'Skywest'];

const assignmentUpdateSchema = z.object({
  value: z.string().trim().min(1).max(120),
});

const aircrewArrivalSchema = z.object({
  carrier: z.string().trim().min(1).max(80),
  date: z.string().regex(/^\d{4}-\d{2}-\d{2}$/),
  time: z.string().trim().min(1).max(20),
  action: z.enum(['add', 'remove']).default('add'),
});

const aircrewCarrierSchema = z.object({
  carrier: z.string().trim().min(1).max(80),
});

const normalizeCarrierLabel = (value: string) =>
  value
    .trim()
    .replace(/\s+/g, ' ')
    .replace(/\b\w/g, (char) => char.toUpperCase());

const normalizeAircrewTime = (value: string) => {
  const token = value.trim().toUpperCase().replace(/\s+/g, '').replace(/\./g, '').replace(/–/g, '-');
  const match = token.match(/^(\d{1,2}):?(\d{2})(AM|PM)?$/);
  if (!match) {
    throw new HttpError(400, 'Invalid time format.');
  }

  let hour = Number(match[1]);
  const minute = Number(match[2]);
  const suffix = match[3];
  if (hour > 23 || minute > 59) {
    throw new HttpError(400, 'Time is out of range.');
  }

  if (suffix) {
    if (hour > 12 || hour === 0) {
      throw new HttpError(400, 'Invalid 12-hour time.');
    }
    if (hour === 12) hour = 0;
    if (suffix === 'PM') hour += 12;
  }

  return `${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}`;
};

const aircrewSortKey = (value: string) => {
  const [hour, minute] = value.split(':').map(Number);
  let totalMinutes = hour * 60 + minute;
  if (totalMinutes >= 0 && totalMinutes < 60) totalMinutes += 24 * 60;
  return totalMinutes;
};

const sortAircrewTimes = (values: string[]) =>
  Array.from(new Set(values.filter(Boolean))).sort((left, right) => {
    const diff = aircrewSortKey(left) - aircrewSortKey(right);
    return diff || left.localeCompare(right);
  });

const deserializeAircrewTimes = (payload: string | null) => {
  if (!payload) return [];
  const cleaned: string[] = [];

  try {
    const data = JSON.parse(payload);
    if (Array.isArray(data)) {
      data.forEach((entry) => {
        try {
          cleaned.push(normalizeAircrewTime(String(entry)));
        } catch {
          // Ignore legacy invalid entries.
        }
      });
      return sortAircrewTimes(cleaned);
    }
  } catch {
    // Fall through to legacy text parsing.
  }

  payload.split(/[,\n/]+/).forEach((entry) => {
    try {
      cleaned.push(normalizeAircrewTime(entry));
    } catch {
      // Ignore legacy invalid entries.
    }
  });
  return sortAircrewTimes(cleaned);
};

const serializeAircrewTimes = (times: string[]) => JSON.stringify(sortAircrewTimes(times));

const getAircrewCarriers = async () => {
  const rows = await prisma.aircrewCarrier.findMany({ orderBy: { name: 'asc' } });
  return Array.from(new Set([...defaultAircrewCarriers, ...rows.map((row) => row.name)])).sort((left, right) =>
    left.localeCompare(right),
  );
};

const assertDateInsideWeek = (week: { startDate: string }, date: string) => {
  const start = new Date(`${week.startDate}T00:00:00`);
  const target = new Date(`${date}T00:00:00`);
  const end = new Date(start);
  end.setDate(start.getDate() + 6);
  if (target < start || target > end) {
    throw new HttpError(400, 'Date outside of selected week.');
  }
};

schedulesRouter.get(
  '/weeks',
  asyncHandler(async (_req, res) => {
    const weeks = await prisma.week.findMany({
      include: {
        _count: {
          select: {
            assignments: true,
          },
        },
      },
      orderBy: { startDate: 'desc' },
    });

    res.json(
      weeks.map((week) => ({
        id: week.id,
        startDate: week.startDate,
        assignmentCount: week._count.assignments,
      })),
    );
  }),
);

schedulesRouter.get(
  '/weeks/:id',
  asyncHandler(async (req, res) => {
    const id = Number(req.params.id);
    if (Number.isNaN(id)) {
      throw new HttpError(400, 'Invalid week id.');
    }

    const week = await prisma.week.findUnique({
      where: { id },
      include: {
        assignments: {
          include: {
            employee: {
              include: {
                section: true,
              },
            },
          },
          orderBy: [
            { employee: { section: { name: 'asc' } } },
            { employee: { sortOrder: 'asc' } },
            { employee: { name: 'asc' } },
            { date: 'asc' },
          ],
        },
        aircrewArrivals: {
          orderBy: [{ carrier: 'asc' }, { date: 'asc' }],
        },
      },
    });

    if (!week) {
      throw new HttpError(404, 'Week not found.');
    }

    const carriers = await getAircrewCarriers();
    const arrivals = carriers.map((carrier) => ({
      carrier,
      timesByDate: Object.fromEntries(
        week.aircrewArrivals
          .filter((arrival) => arrival.carrier === carrier)
          .map((arrival) => [arrival.date, deserializeAircrewTimes(arrival.times)]),
      ),
    }));

    res.json({
      ...week,
      aircrewArrivals: arrivals,
    });
  }),
);

schedulesRouter.patch(
  '/assignments/:id',
  asyncHandler(async (req, res) => {
    const id = Number(req.params.id);
    if (Number.isNaN(id)) {
      throw new HttpError(400, 'Invalid assignment id.');
    }

    const input = assignmentUpdateSchema.parse(req.body);
    const assignment = await prisma.assignment.update({
      where: { id },
      data: {
        value: input.value,
      },
      include: {
        employee: {
          include: {
            section: true,
          },
        },
      },
    });

    res.json(assignment);
  }),
);

schedulesRouter.post(
  '/weeks/:id/aircrew/arrival',
  asyncHandler(async (req, res) => {
    const id = Number(req.params.id);
    if (Number.isNaN(id)) {
      throw new HttpError(400, 'Invalid week id.');
    }

    const week = await prisma.week.findUnique({ where: { id } });
    if (!week) {
      throw new HttpError(404, 'Week not found.');
    }

    const input = aircrewArrivalSchema.parse(req.body);
    const carrier = normalizeCarrierLabel(input.carrier);
    assertDateInsideWeek(week, input.date);

    const normalizedTime = normalizeAircrewTime(input.time);
    const existing = await prisma.aircrewArrival.findFirst({
      where: {
        weekId: id,
        carrier,
        date: input.date,
      },
    });

    const currentTimes = deserializeAircrewTimes(existing?.times ?? null);
    const nextTimes =
      input.action === 'remove'
        ? currentTimes.filter((time) => time !== normalizedTime)
        : sortAircrewTimes([...currentTimes, normalizedTime]);

    if (existing && nextTimes.length === 0) {
      await prisma.aircrewArrival.delete({ where: { id: existing.id } });
    } else if (existing) {
      await prisma.aircrewArrival.update({
        where: { id: existing.id },
        data: { times: serializeAircrewTimes(nextTimes) },
      });
    } else if (nextTimes.length > 0) {
      await prisma.aircrewArrival.create({
        data: {
          weekId: id,
          carrier,
          date: input.date,
          times: serializeAircrewTimes(nextTimes),
        },
      });
    }

    res.json({
      carrier,
      weekId: id,
      cells: {
        [input.date]: nextTimes,
      },
    });
  }),
);

schedulesRouter.post(
  '/aircrew/carriers',
  asyncHandler(async (req, res) => {
    const input = aircrewCarrierSchema.parse(req.body);
    const carrier = normalizeCarrierLabel(input.carrier);
    await prisma.aircrewCarrier.upsert({
      where: { name: carrier },
      create: { name: carrier },
      update: {},
    });

    res.status(201).json({
      carrier,
      carriers: await getAircrewCarriers(),
    });
  }),
);

schedulesRouter.delete(
  '/aircrew/carriers/:carrier',
  asyncHandler(async (req, res) => {
    const carrierParam = Array.isArray(req.params.carrier)
      ? req.params.carrier[0]
      : req.params.carrier;
    const carrier = normalizeCarrierLabel(decodeURIComponent(carrierParam));
    if (defaultAircrewCarriers.includes(carrier)) {
      throw new HttpError(400, 'Default carriers cannot be removed.');
    }

    const existing = await prisma.aircrewCarrier.findUnique({ where: { name: carrier } });
    if (!existing) {
      throw new HttpError(404, 'Carrier not found.');
    }

    await prisma.$transaction([
      prisma.aircrewArrival.deleteMany({ where: { carrier } }),
      prisma.aircrewCarrier.delete({ where: { name: carrier } }),
    ]);

    res.json({
      carrier,
      carriers: await getAircrewCarriers(),
    });
  }),
);
