import { createApp } from './app';
import { logger } from './services/logger';

const port = Number(process.env.PORT || 3010);
const app = createApp();

app.listen(port, () => {
  logger.info('document artifact service listening', { port });
});
