import { Router } from 'express';

export const healthRouter = Router();

healthRouter.get('/health', (_req, res) => {
  res.json({
    status: 'ok',
    service: 'bblaw-doc-artifacts',
    timestamp: new Date().toISOString(),
  });
});
