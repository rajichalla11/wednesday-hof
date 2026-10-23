// Sends web push notifications for Wednesday 9–10. Only the Apps Script backend calls this (it knows the shared secret);
// it decides who gets what and passes the subscriptions. Returns the endpoints that are gone so the backend can forget them.
const webpush = require('web-push');
const keys = (() => { try { return require('./_keys.js'); } catch (_) { return {}; } })();   // kept out of git; also read from env if set
const PUB = process.env.VAPID_PUBLIC || keys.pub, PRIV = process.env.VAPID_PRIVATE || keys.priv, SECRET = process.env.PUSH_SECRET || keys.secret;
webpush.setVapidDetails('mailto:raji.doubleone@gmail.com', PUB, PRIV);
module.exports = async (req, res) => {
  if (req.method !== 'POST') return res.status(405).json({ ok: false });
  let body = req.body; if (typeof body === 'string') { try { body = JSON.parse(body); } catch (_) { body = {}; } }
  if (!body || body.secret !== SECRET) return res.status(403).json({ ok: false, error: 'forbidden' });
  const msgs = Array.isArray(body.messages) ? body.messages.slice(0, 500) : [];
  const gone = [], failed = [];
  await Promise.all(msgs.map(async m => {
    try {
      await webpush.sendNotification(m.sub, JSON.stringify({ title: m.title, body: m.body, url: m.url || '/', tag: m.tag || '' }), { TTL: m.ttl || 6 * 3600, urgency: 'high' });
    } catch (e) { if (e.statusCode === 404 || e.statusCode === 410) gone.push(m.sub && m.sub.endpoint); else failed.push(String(e.statusCode || e.message)); }
  }));
  res.status(200).json({ ok: true, sent: msgs.length - gone.length - failed.length, gone, failed });
};
