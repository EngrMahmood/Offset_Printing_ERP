# Deploying Offset ERP to Production

Two Oracle Cloud VMs run identical Docker Compose stacks — a **primary** and
a **standby**. Both need to be updated on every deploy; they run separate
databases (not a live replica pair), so deploying to one does not touch the
other.

The old Windows LAN server (192.168.88.30) this doc used to describe is
**retired**. These two Oracle VMs are the only production targets.

## Servers

| Role | Host | SSH user | SSH key (on this dev machine) | Arch / hostname |
|---|---|---|---|---|
| Primary | `offseterp.duckdns.org` | `ubuntu` | `C:\Users\Universal Engr\.ssh\offset-erp-oracle-a1` (ED25519) | ARM64, `offset-erp-a1` |
| Standby | `offseterpbackup.duckdns.org` | `ubuntu` | `C:\Users\Universal Engr\.ssh\offset-erp-oracle.key` (RSA) | x86_64, `offset-erp-server` |

The two servers use **different keys** — don't assume one key works for both.
Both hosts are already in `known_hosts` on this machine.

Both run from `/home/ubuntu/offset-erp` (a clone of this repo, `main`
branch) via `docker-compose.yml` in that directory: three services —
`redis`, `web` (the Django app via Daphne), `nginx` (TLS termination,
reverse proxy). SQLite is the database engine on both (not Postgres),
matching what `Offset_ERP/settings.py` defaults to.

## Deploying a change

Once your change is committed and pushed to `origin/main`, run this against
**each** server in turn (primary first, then standby):

```bash
ssh -i <key for that host> ubuntu@<host> "cd /home/ubuntu/offset-erp && git pull origin main"
ssh -i <key for that host> ubuntu@<host> "cd /home/ubuntu/offset-erp && sudo docker compose build web"
ssh -i <key for that host> ubuntu@<host> "cd /home/ubuntu/offset-erp && sudo docker compose up -d web"
```

**The `build` step is not optional.** The image has no source bind-mount —
`Dockerfile` `COPY`s the working tree in at build time — so `git pull`
alone changes nothing running until the image is rebuilt and the container
recreated. Skipping `build` and running only `up -d` silently redeploys the
*old* code.

Then confirm a clean startup:

```bash
ssh -i <key for that host> ubuntu@<host> "cd /home/ubuntu/offset-erp && sudo docker compose logs web --tail 15"
```

Look for `Listening on TCP address 0.0.0.0:8000` with no tracebacks above
it. `sudo docker compose ps` should show `web`, `redis`, and `nginx` all
`Up`.

## What happens automatically vs. what doesn't

`docker-entrypoint.sh` runs on every container start (so on every deploy),
in order:

1. `manage.py migrate --noinput`
2. `seed_access_control`, `seed_chat_permissions`, `seed_viewer_role`,
   `seed_bots`, `seed_bom_masters`, `seed_bom_permissions` (each `|| true`
   — safe to re-run, won't fail the boot if something's already seeded)
3. `backfill_raw_items_from_rm_skus` (`|| true`)
4. `collectstatic --noinput` — needed on **every** start, not just the
   first: `staticfiles` is a named Docker volume, so the image's build-time
   `collectstatic` output never reaches the live volume after the first
   deploy.
5. `daphne` starts and listens on `:8000`.

So **migrations and permission/master-data seeding never need a manual
step** — they're already covered by the steps above.

**One-off management commands** (a backfill script written for a specific
bug, `--apply` flags, etc.) are *not* run automatically and need an explicit
call after the container is up:

```bash
ssh -i <key for that host> ubuntu@<host> "cd /home/ubuntu/offset-erp && sudo docker compose exec web python manage.py <command> [--apply]"
```

Always dry-run first (omit `--apply`) if the command supports it, read the
output, then re-run with `--apply`.

## Notes

- `docker compose` (not `docker-compose`) — the host uses the Compose v2
  plugin. Commands need `sudo`.
- The `docker-compose.yml: the attribute 'version' is obsolete` warning on
  every command is cosmetic — ignore it.
- The standby has its in-process bot/backup schedulers disabled via
  settings (visible in its boot log as `... disabled via settings.
  Skipping.`) so it doesn't double-send the scheduled emails/backups the
  primary already sends. This is expected, not a bug.
- Redis and Nginx are almost never touched by a normal deploy — `up -d web`
  only recreates the `web` service, leaving `redis`/`nginx` running
  undisturbed (they've been up for weeks in practice).
- Both servers pull from the same `origin/main` — there is no separate
  release branch. A push to `main` is a push toward production; treat it
  accordingly.
