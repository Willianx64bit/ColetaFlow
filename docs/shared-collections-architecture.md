# ColetaFlow — shared collection architecture

This branch adds the backend foundation for collaborative inventory collection.
It is intentionally separate from MobFlow and does not change the existing local/offline collector.

## Target flow
- A manager creates a collection and receives a short access code.
- Each collector joins with the code and a display name; the server issues a random participant token.
- The browser persists each scan locally first (IndexedDB), then syncs it using a stable UUID event_id.
- The server inserts events with `ON CONFLICT (event_id) DO NOTHING`, making retries idempotent.
- Manager progress is computed from server-side grouped totals, never from last-write-wins client state.
- Offline queues remain on each device and are retried when connectivity returns.

## Database
Apply `migrations/001_shared_collections.sql` only to the dedicated ColetaFlow Neon database.
Tables: collections, participants, scan events. `cf_collection_totals` aggregates quantities by barcode.

## Security rules
- Never expose DATABASE_URL or token hashes in browser code.
- API endpoints must validate participant tokens and collection membership on every request.
- Store only hashes of participant tokens.
- Rate-limit access-code join attempts.
- Do not trust quantities, participant IDs, or collection IDs supplied by the client without server validation.

## Rollout
1. Provision the dedicated Neon database and set DATABASE_URL in Vercel.
2. Apply the migration.
3. Add server API endpoints and the shared-collection UI.
4. Keep shared mode opt-in; local-only/offline mode remains unchanged.
5. Load-test concurrent event submissions and verify event counts and grouped totals before production enablement.
