import {
  backup,
  resume,
  bootstrapOwner,
  closeCollection,
  collections,
  createCollection,
  deleteCollection,
  exportCollection,
  exportUnmatched,
  finishParticipant,
  joinCollection,
  login,
  logout,
  mapping,
  me,
  progress,
  register,
  sync,
  template,
  users
} from '../../../lib/shared/index.js';

const routes = {
  backup,
  resume,
  'bootstrap-owner': bootstrapOwner,
  close: closeCollection,
  collections,
  create: createCollection,
  delete: deleteCollection,
  export: exportCollection,
  'export-unmatched': exportUnmatched,
  finish: finishParticipant,
  join: joinCollection,
  login,
  logout,
  mapping,
  me,
  progress,
  register,
  sync,
  template,
  users
};

export async function onRequest({ request, env }) {
  const { configureRuntime } = await import('../../../lib/shared/_db.js');
  configureRuntime(env);

  const url = new URL(request.url);
  const routeName = url.pathname.slice('/api/shared/'.length).replace(/\/$/, '');
  const handler = routes[routeName];
  if (!handler) return Response.json({ error: 'Rota não encontrada' }, { status: 404 });

  let body = {};
  if (request.method !== 'GET' && request.method !== 'HEAD') {
    const contentType = request.headers.get('content-type') || '';
    if (contentType.includes('application/json')) {
      try { body = await request.json(); }
      catch { return Response.json({ error: 'JSON inválido' }, { status: 400 }); }
    }
  }

  let status = 200;
  let responseBody = null;
  const responseHeaders = new Headers({ 'Cache-Control': 'no-store' });
  const req = {
    method: request.method,
    headers: Object.fromEntries(request.headers.entries()),
    query: Object.fromEntries(url.searchParams.entries()),
    body
  };
  const res = {
    status(code) { status = code; return this; },
    setHeader(name, value) { responseHeaders.set(name, value); return this; },
    json(value) { responseBody = value; return this; }
  };

  try {
    await handler(req, res);
    return Response.json(responseBody ?? {}, { status, headers: responseHeaders });
  } catch (error) {
    console.error('ColetaFlow Pages Function', routeName, error);
    return Response.json({ error: 'Erro interno do servidor' }, { status: 500, headers: responseHeaders });
  }
}
