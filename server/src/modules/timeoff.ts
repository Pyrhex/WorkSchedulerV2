import { Router } from 'express';
import { z } from 'zod';

import { HttpError, asyncHandler } from '../lib/http';
import { prisma } from '../lib/prisma';

export const timeOffRouter = Router();

const timeOffSchema = z.object({
  employeeName: z.string().trim().min(1).max(120),
  role: z.string().trim().min(1).max(80),
  startDate: z.string().regex(/^\d{4}-\d{2}-\d{2}$/),
  endDate: z.string().regex(/^\d{4}-\d{2}-\d{2}$/),
  approved: z.boolean().optional(),
  vacation: z.boolean().optional(),
});

timeOffRouter.get(
  '/',
  asyncHandler(async (_req, res) => {
    const requests = await prisma.timeOff.findMany({
      orderBy: [{ startDate: 'desc' }, { employeeName: 'asc' }],
    });

    res.json(requests);
  }),
);

timeOffRouter.post(
  '/',
  asyncHandler(async (req, res) => {
    const input = timeOffSchema.parse(req.body);
    const request = await prisma.timeOff.create({
      data: {
        employeeName: input.employeeName,
        role: input.role,
        startDate: input.startDate,
        endDate: input.endDate,
        approved: input.approved ?? false,
        vacation: input.vacation ?? false,
      },
    });

    res.status(201).json(request);
  }),
);

timeOffRouter.delete(
  '/:id',
  asyncHandler(async (req, res) => {
    const id = Number(req.params.id);
    if (Number.isNaN(id)) {
      throw new HttpError(400, 'Invalid time off request id.');
    }

    await prisma.timeOff.delete({ where: { id } });
    res.status(204).send();
  }),
);
