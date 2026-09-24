// 실행 카운터 — Cloudflare KV(HITS)에 누적 실행 횟수 + 일별 실행 횟수를 저장
// GET  /api/hits             : 누적값 + 오늘(KST) 실행 횟수 반환
// GET  /api/hits?days=N      : 누적값 + 최근 N일(KST) 일별 실행 횟수 반환 (N<=31)
// POST /api/hits              : 누적/오늘 카운트를 1 증가시키고 반환 (성공 시에만 호출됨)
const KEY_TOTAL = "total";

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type, Authorization",
};

function json(body) {
  return new Response(JSON.stringify(body), {
    headers: {
      "content-type": "application/json",
      "cache-control": "no-store",
      ...CORS_HEADERS,
    },
  });
}

function kstDateKey(date = new Date()) {
  const kst = new Date(date.getTime() + 9 * 60 * 60 * 1000);
  return `d:${kst.toISOString().slice(0, 10)}`;
}

async function readInt(kv, key) {
  const v = parseInt((await kv.get(key)) || "0", 10);
  return Number.isFinite(v) ? v : 0;
}

export async function onRequest({ request, env }) {
  if (request.method === "OPTIONS") {
    return new Response(null, { status: 204, headers: CORS_HEADERS });
  }

  const kv = env.HITS;
  if (!kv) return json({ count: null }); // 바인딩 없으면 조용히 무시

  if (request.method === "POST") {
    const n = (await readInt(kv, KEY_TOTAL)) + 1;
    await kv.put(KEY_TOTAL, String(n));

    const dKey = kstDateKey();
    const dn = (await readInt(kv, dKey)) + 1;
    await kv.put(dKey, String(dn));

    return json({ count: n, today: dn });
  }

  const n = await readInt(kv, KEY_TOTAL);
  const url = new URL(request.url);
  const days = Math.min(parseInt(url.searchParams.get("days") || "0", 10) || 0, 31);

  if (days > 0) {
    const daily = [];
    const now = new Date();
    for (let i = days - 1; i >= 0; i--) {
      const d = new Date(now.getTime() - i * 86400000);
      const key = kstDateKey(d);
      daily.push({ date: key.slice(2), count: await readInt(kv, key) });
    }
    return json({ count: n, daily });
  }

  return json({ count: n, today: await readInt(kv, kstDateKey()) });
}
