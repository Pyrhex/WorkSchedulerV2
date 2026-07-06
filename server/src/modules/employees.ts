import { Router } from 'express';
import { z } from 'zod';

import { HttpError, asyncHandler } from '../lib/http';
import { prisma } from '../lib/prisma';

export const employeesRouter = Router();

const employeeSchema = z.object({
  name: z.string().trim().min(1).max(120),
  sectionId: z.number().int().positive(),
  availability: z.string().trim().max(500).optional().nullable(),
  preferredShift: z.string().trim().max(120).optional().nullable(),
  seniority: z.number().int().min(0).optional().nullable(),
  preferredShiftsPerWeek: z.number().int().min(0).optional().nullable(),
  maxShiftsPerWeek: z.number().int().min(0).optional().nullable(),
  sortOrder: z.number().int().min(0).optional().nullable(),
  firstName: z.string().trim().max(80).optional().nullable(),
  lastName: z.string().trim().max(80).optional().nullable(),
  temporary: z.boolean().optional(),
  roleIds: z.array(z.number().int().positive()).optional(),
});

employeesRouter.get(
  '/',
  asyncHandler(async (_req, res) => {
    const employees = await prisma.employee.findMany({
      include: {
        section: true,
        roles: {
          include: {
            section: true,
          },
          orderBy: {
            section: {
              name: 'asc',
            },
          },
        },
      },
      orderBy: [{ section: { name: 'asc' } }, { sortOrder: 'asc' }, { name: 'asc' }],
    });

    res.json(employees);
  }),
);

employeesRouter.post(
  '/',
  asyncHandler(async (req, res) => {
    const input = employeeSchema.parse(req.body);

    const employee = await prisma.employee.create({
      data: {
        name: input.name,
        sectionId: input.sectionId,
        availability: input.availability,
        preferredShift: input.preferredShift,
        seniority: input.seniority,
        preferredShiftsPerWeek: input.preferredShiftsPerWeek,
        maxShiftsPerWeek: input.maxShiftsPerWeek,
        sortOrder: input.sortOrder,
        firstName: input.firstName,
        lastName: input.lastName,
        temporary: input.temporary ?? false,
        roles: input.roleIds?.length
          ? {
              create: input.roleIds.map((sectionId) => ({
                section: {
                  connect: { id: sectionId },
                },
              })),
            }
          : undefined,
      },
      include: {
        section: true,
        roles: {
          include: {
            section: true,
          },
        },
      },
    });

    res.status(201).json(employee);
  }),
);

employeesRouter.put(
  '/:id',
  asyncHandler(async (req, res) => {
    const id = Number(req.params.id);
    if (Number.isNaN(id)) {
      throw new HttpError(400, 'Invalid employee id.');
    }

    const input = employeeSchema.parse(req.body);

    const employee = await prisma.employee.update({
      where: { id },
      data: {
        name: input.name,
        sectionId: input.sectionId,
        availability: input.availability,
        preferredShift: input.preferredShift,
        seniority: input.seniority,
        preferredShiftsPerWeek: input.preferredShiftsPerWeek,
        maxShiftsPerWeek: input.maxShiftsPerWeek,
        sortOrder: input.sortOrder,
        firstName: input.firstName,
        lastName: input.lastName,
        temporary: input.temporary ?? false,
      },
      include: {
        section: true,
        roles: {
          include: {
            section: true,
          },
        },
      },
    });

    res.json(employee);
  }),
);

employeesRouter.delete(
  '/:id',
  asyncHandler(async (req, res) => {
    const id = Number(req.params.id);
    if (Number.isNaN(id)) {
      throw new HttpError(400, 'Invalid employee id.');
    }

    await prisma.employee.delete({ where: { id } });
    res.status(204).send();
  }),
);
