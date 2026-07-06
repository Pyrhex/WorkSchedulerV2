import cors from 'cors';
import express from 'express';

import { errorHandler, notFound } from './lib/http';
import { employeesRouter } from './modules/employees';
import { schedulesRouter } from './modules/schedules';
import { sectionsRouter } from './modules/sections';
import { timeOffRouter } from './modules/timeoff';

export const app = express();

app.use(
  cors({
    origin: ['http://localhost:5173', 'http://localhost:5174'],
  }),
);
app.use(express.json({ limit: '10mb' }));

app.use('/api', (_req, res, next) => {
  res.setHeader('Cache-Control', 'no-store');
  next();
});

app.get('/api/health', (_req, res) => {
  res.json({ status: 'ok' });
});

app.use('/api/employees', employeesRouter);
app.use('/api/sections', sectionsRouter);
app.use('/api/schedules', schedulesRouter);
app.use('/api/time-off', timeOffRouter);

app.use(notFound);
app.use(errorHandler);
