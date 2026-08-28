import { Redis } from '@upstash/redis';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const INDICATOR_DATA = JSON.parse(readFileSync(path.join(__dirname, '..', 'src', 'data', 'indicators.json'), 'utf8'));

function getRedis() {
  const url = process.env.KV_REST_API_URL || process.env.UPSTASH_REDIS_REST_URL;
  const token = process.env.KV_REST_API_TOKEN || process.env.UPSTASH_REDIS_REST_TOKEN;
  if (!url || !token) {
    const err = new Error('Database Redis (Upstash) belum terkonfigurasi di server ini. Pasang integrasi "Upstash for Redis" lewat tab Storage di dashboard Vercel — lihat README bagian Troubleshooting.');
    err.code = 'REDIS_NOT_CONFIGURED';
    throw err;
  }
  return new Redis({ url, token });
}

const ADMIN_CODE = process.env.ADMIN_CODE || 'akreditasi2026';
function normalize(v) {
  return (v || '').toString().trim();
}
function checkAdminCode(input) {
  return normalize(input) === normalize(ADMIN_CODE);
}

function emptyEvidence() {
  const ev = {};
  for (const k of ['kp', 'kk', 'il']) {
    INDICATOR_DATA[k].butir.forEach((b) =>
      b.items.forEach((it) => {
        ev[it.code] = { skor: 0, buktiTersedia: '', buktiPerlu: '', lokasi: '', catatan: '', updatedAt: null };
      })
    );
  }
  return ev;
}

function parseBody(req) {
  if (!req.body) return {};
  if (typeof req.body === 'string') {
    try { return JSON.parse(req.body); } catch (e) { return {}; }
  }
  return req.body;
}

export default async function handler(req, res) {
  const schoolId = req.query && req.query.schoolId;

  let redis;
  try {
    redis = getRedis();
  } catch (e) {
    return res.status(500).json({ error: e.message, code: e.code || 'REDIS_NOT_CONFIGURED' });
  }

  if (req.method === 'GET') {
    if (!schoolId) return res.status(400).json({ error: 'schoolId wajib diisi' });
    let data;
    try {
      data = await redis.get('evidence:' + schoolId);
    } catch (e) {
      return res.status(500).json({ error: 'Gagal membaca database Redis.', code: 'REDIS_NOT_CONFIGURED' });
    }
    return res.status(200).json(data || emptyEvidence());
  }

  if (req.method === 'PUT') {
    const body = parseBody(req);
    const { schoolId: sid, evidence } = body;
    if (!sid || !evidence) return res.status(400).json({ error: 'schoolId dan evidence wajib diisi' });
    await redis.set('evidence:' + sid, evidence);
    return res.status(200).json({ ok: true });
  }

  if (req.method === 'DELETE') {
    const body = parseBody(req);
    const sid = schoolId || body.schoolId;
    const adminCode = (req.headers && (req.headers['x-admin-code'] || req.headers['X-Admin-Code'])) || body.adminCode;
    if (!checkAdminCode(adminCode)) return res.status(403).json({ error: 'Kode admin salah.' });
    if (!sid) return res.status(400).json({ error: 'schoolId wajib diisi' });
    const empty = emptyEvidence();
    await redis.set('evidence:' + sid, empty);
    return res.status(200).json(empty);
  }

  return res.status(405).json({ error: 'Metode tidak didukung' });
}
