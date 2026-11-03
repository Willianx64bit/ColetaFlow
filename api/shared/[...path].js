import * as handlers from '../../lib/shared/index.js';
import { configureRuntime } from '../../lib/shared/_db.js';

const routes = {
  'bootstrap-owner': handlers.bootstrapOwner,
  close: handlers.closeCollection,
  collections: handlers.collections,
  create: handlers.createCollection,
  export: handlers.exportCollection,
  'export-unmatched': handlers.exportUnmatched,
  finish: handlers.finishParticipant,
  join: handlers.joinCollection,
  login: handlers.login,
  logout: handlers.logout,
  mapping: handlers.mapping,
  me: handlers.me,
  progress: handlers.progress,
  register: handlers.register,
  sync: handlers.sync,
  template: handlers.template,
  users: handlers.users
};

export default async function handler(req, res) {
  try {
    configureRuntime(process.env);
    const query = req.query || {};
    const url = new URL(req.url || '/', 'https://coletaflow.invalid');
    const pathValue = query.path ?? url.pathname.slice('/api/shared/'.length);
    const parts = Array.isArray(pathValue) ? pathValue : [pathValue].filter(Boolean);
    const joinedPath = parts.join('/');
    const routeName = joinedPath.endsWith('/') ? joinedPath.slice(0, -1) : joinedPath;
    const action = routes[routeName];
    res.setHeader('Cache-Control', 'no-store');
    if (!action) return res.status(404).json({ error: 'Rota não encontrada', route: routeName });

    req.headers = req.headers || {};
    req.query = { ...query };
    delete req.query.path;
    await action(req, res);
  } catch (error) {
    console.error('ColetaFlow API invocation', error);
    if (!res.headersSent) return res.status(500).json({ error: 'Falha interna na API ColetaFlow' });
  }
}
