import { Router } from 'express';
import { capabilityManifest } from '../config/capability';

export const manifestRouter = Router();

manifestRouter.get('/manifest', (_req, res) => {
  res.json(capabilityManifest);
});
