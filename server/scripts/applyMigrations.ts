import { spawnSync } from 'node:child_process';
import { existsSync } from 'node:fs';
import { resolve } from 'node:path';

const databaseUrl = process.env.DATABASE_URL ?? 'file:./dev.db';
const schemaPath = resolve(__dirname, '../prisma/schema.prisma');

if (!existsSync(schemaPath)) {
  throw new Error(`Prisma schema not found at ${schemaPath}`);
}

const result = spawnSync(
  'npx',
  ['prisma', 'migrate', 'dev', '--schema', schemaPath],
  {
    env: {
      ...process.env,
      DATABASE_URL: databaseUrl,
    },
    stdio: 'inherit',
  },
);

process.exit(result.status ?? 1);
