-- Normalize the platform account model to exactly two account types:
-- one platform owner and invited users. Apply only to the dedicated ColetaFlow database.
BEGIN;

ALTER TABLE cf_users DROP CONSTRAINT IF EXISTS cf_users_role_check;
UPDATE cf_users u
SET role = CASE
  WHEN EXISTS (SELECT 1 FROM cf_platform_owner o WHERE o.user_id = u.id) THEN 'owner'
  ELSE 'user'
END;
ALTER TABLE cf_users ALTER COLUMN role SET DEFAULT 'user';
ALTER TABLE cf_users ADD CONSTRAINT cf_users_role_check CHECK (role IN ('owner','user'));

ALTER TABLE cf_participants DROP CONSTRAINT IF EXISTS cf_participants_role_check;
UPDATE cf_participants p
SET role = CASE
  WHEN EXISTS (SELECT 1 FROM cf_platform_owner o WHERE o.user_id = p.user_id) THEN 'owner'
  ELSE 'user'
END;
ALTER TABLE cf_participants ALTER COLUMN role SET DEFAULT 'user';
ALTER TABLE cf_participants ADD CONSTRAINT cf_participants_role_check CHECK (role IN ('owner','user'));

COMMIT;
