// Public results (cards, games, man of the match) cached at Vercel's edge, so a first visit doesn't wait on Apps Script.
// The CDN keeps a copy for 60 s and serves the last good copy for up to 10 min while it refreshes in the background.
// Nothing private goes through here: it is the same JSON the site already fetches from the public web app.
export const config = { runtime: 'edge' };
const API = 'https://script.google.com/macros/s/AKfycby3NDACkLNP9_MKwWQYqPTFPROcDiAoXbpOnbtMxYGOvSAVvfeAHIeKLb4krs-mcHGf/exec';
export default async function handler() {
  try {
    const r = await fetch(API + '?action=results', { redirect: 'follow' });
    const body = await r.text();
    if (!r.ok || !body.trim().startsWith('{')) throw new Error('bad upstream');
    return new Response(body, { headers: { 'content-type': 'application/json; charset=utf-8', 'cache-control': 'public, s-maxage=60, stale-while-revalidate=600' } });
  } catch (e) {
    return new Response(JSON.stringify({ ok: false, error: 'upstream' }), { status: 502, headers: { 'content-type': 'application/json', 'cache-control': 'no-store' } });
  }
}
