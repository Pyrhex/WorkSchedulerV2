import { Router } from 'express';
import { z } from 'zod';

import { HttpError, asyncHandler } from '../lib/http';
import { prisma } from '../lib/prisma';

export const sectionsRouter = Router();

const sectionSchema = z.object({
  name: z.string().trim().min(1).max(80),
  requiredPerDay: z.number().int().min(0).optional().nullable(),
});

sectionsRouter.get(
  '/',
  asyncHandler(async (_req, res) => {
    const sections = await prisma.section.findMany({
      orderBy: { name: 'asc' },
    });

    res.json(sections);
  }),
);

sectionsRouter.post(
  '/',
  asyncHandler(async (req, res) => {
    const input = sectionSchema.parse(req.body);
    const section = await prisma.section.create({
      data: {
        name: input.name,
        requiredPerDay: input.requiredPerDay,
      },
    });

    res.status(201).json(section);
  }),
);

sectionsRouter.put(
  '/:id',
  asyncHandler(async (req, res) => {
    const id = Number(req.params.id);
    if (Number.isNaN(id)) {
      throw new HttpError(400, 'Invalid section id.');
    }

    const input = sectionSchema.parse(req.body);
    const section = await prisma.section.update({
      where: { id },
      data: {
        name: input.name,
        requiredPerDay: input.requiredPerDay,
      },
    });

    res.json(section);
  }),
);
