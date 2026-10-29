// The site's page, served from Vercel's edge cache instead of going to GitHub Pages on every visit (~0.9 s → ~0.1 s).
// Fresh for 60 s (so a push to GitHub shows within a minute), then served stale for up to a day while it refreshes.
export const config = { runtime: 'edge' };
const ORIGIN = 'https://rajichalla11.github.io/wednesday-hof/';
// ?fresh=1 (the app's background check for a new version) skips both caches, so one open is enough to pick up the latest.
export default async function handler(req) {
  const fresh = new URL(req.url).searchParams.has('fresh');
  try {
    const r = await fetch(fresh ? ORIGIN + '?t=' + Date.now() : ORIGIN, { redirect: 'follow' });
    const body = await r.text();
    if (!r.ok || !body.includes('<html')) throw new Error('bad upstream');
    return new Response(body, { headers: { 'content-type': 'text/html; charset=utf-8',
      'cache-control': 'public, max-age=0, must-revalidate',   // browsers always check for a newer page
      'cdn-cache-control': fresh ? 'no-store' : 'max-age=60, stale-while-revalidate=86400',   // Vercel's edge keeps a copy
      'x-page-version': r.headers.get('etag') || '' } });
  } catch (e) {
    return Response.redirect(ORIGIN, 302);
  }
}
