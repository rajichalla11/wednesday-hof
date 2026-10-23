// The site's page, served from Vercel's edge cache instead of going to GitHub Pages on every visit (~0.9 s → ~0.1 s).
// Fresh for 60 s (so a push to GitHub shows within a minute), then served stale for up to a day while it refreshes.
export const config = { runtime: 'edge' };
const ORIGIN = 'https://rajichalla11.github.io/wednesday-hof/';
export default async function handler() {
  try {
    const r = await fetch(ORIGIN, { redirect: 'follow' });
    const body = await r.text();
    if (!r.ok || !body.includes('<html')) throw new Error('bad upstream');
    return new Response(body, { headers: { 'content-type': 'text/html; charset=utf-8', 'cache-control': 'public, max-age=0, s-maxage=60, stale-while-revalidate=86400', 'etag': r.headers.get('etag') || '' } });
  } catch (e) {
    return Response.redirect(ORIGIN, 302);
  }
}
