import cors from 'cors';
import express from 'express';
import { artifactsRouter } from './routes/artifacts';
import { healthRouter } from './routes/health';
import { manifestRouter } from './routes/manifest';
import { logger } from './services/logger';

export function createApp() {
  const app = express();

  app.use(cors({ origin: process.env.CORS_ORIGIN || true, credentials: true }));
  app.use(express.json({ limit: process.env.JSON_BODY_LIMIT || '5mb' }));
  app.use(healthRouter);
  app.use(manifestRouter);
  app.use(artifactsRouter);

  app.use((err: unknown, _req: express.Request, res: express.Response, _next: express.NextFunction) => {
    logger.error('request failed', { error: err instanceof Error ? err.message : String(err) });
    res.status(500).json({ error: 'Document artifact request failed' });
  });

  return app;
}
