# wednesday910

Serves https://rajichalla11.github.io/wednesday-hof/ at a cleaner address via a Vercel rewrite.
`/api/results` is an edge function that caches the public ratings/games JSON for 60 s (stale-while-revalidate 10 min) so first visits load fast.
`/api/sheet` is the attendance-sheet CSV, cached the same way (Google takes ~2 s to export it).
`/` is served by `/api/page`, an edge function that caches the page from GitHub Pages for 60 s (stale-while-revalidate a day), ~0.1 s instead of ~0.9 s.
`/api/push` sends web push notifications; only the Apps Script backend calls it (shared secret). Keys are in `api/_keys.js`, which is not in git (or env vars VAPID_PUBLIC / VAPID_PRIVATE / PUSH_SECRET).
Source and deploys live in https://github.com/rajichalla11/wednesday-hof (this folder is deployed to the Vercel project by hand).
