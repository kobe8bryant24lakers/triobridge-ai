import { dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createSharedModelStore } from './shared/shared-model-store.mjs';
import { modelAction } from './models.mjs';
import { servePane } from './pane-session.mjs';

const store = createSharedModelStore();
await servePane({
  host: process.argv[2],
  root: dirname(fileURLToPath(import.meta.url)),
  input: process.stdin,
  output: process.stdout,
  model: (action, payload, signal) => modelAction(store, action, payload, signal),
});
