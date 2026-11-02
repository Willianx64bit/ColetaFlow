import * as handlers from '../../lib/shared/index.js';
import { configureRuntime } from '../../lib/shared/_db.js';

const routes = {
  'bootstrap-owner': handlers.bootstrapOwner,
  close: handlers.closeCollection,
  collections: handlers.collections,
  create: handlers.createCollection,
  export: handlers.exportCollection,
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
  configureRuntime(process.env);
  const parts = Array.isArray(req.query.path) ? req.query.path : [req.query.path].filter(Boolean);
  const routeName = parts.join('/').replace(/\\/$/, '');
  const action = routes[routeName];
  if (!action) return res.status(404).setHeader('Cache-Control', 'no-store').json({ error: 'Rota não encontrada' });

  req.headers = req.headers || {};
  req.query = { ...req.query };
  delete req.query.path;
  res.setHeader('Cache-Control', 'no-store');
  try {
    await action(req, res);
  } catch (error) {
    console.error('ColetaFlow API', routeName, error);
    if (!res.headersSent) return res.status(500).json({ error: 'Erro interno do servidor' });
  }
}
