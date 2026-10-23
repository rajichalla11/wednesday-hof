// The public attendance sheet (CSV) cached at Vercel's edge: Google takes 1.5–2 s to export it, the edge copy ~0.1 s.
// Fresh for 60 s, then served stale for up to 10 min while it refreshes in the background. Same public data as before.
export const config = { runtime: 'edge' };
const CSV = 'https://docs.google.com/spreadsheets/d/1UUcM5qxOEA-Ykr-gIG4WSHDlzpoi0TJAaRh-8HgYRaI/export?format=csv&gid=0';
export default async function handler() {
  try {
    const r = await fetch(CSV, { redirect: 'follow' });
    const body = await r.text();
    if (!r.ok || !/games played/i.test(body.slice(0, 4000))) throw new Error('bad upstream');
    return new Response(body, { headers: { 'content-type': 'text/csv; charset=utf-8', 'cache-control': 'public, s-maxage=60, stale-while-revalidate=600' } });
  } catch (e) {
    return new Response('upstream', { status: 502, headers: { 'cache-control': 'no-store' } });
  }
}
