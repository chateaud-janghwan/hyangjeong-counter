// 방문자 카운터 — Cloudflare KV(HITS)에 총 방문수를 저장
// GET  /api/hits  : 현재 값 반환(증가 없음)
// POST /api/hits  : 1 증가 후 값 반환
const KEY = "total";

function json(body) {
  return new Response(JSON.stringify(body), {
    headers: { "content-type": "application/json", "cache-control": "no-store" },
  });
}

export async function onRequest({ request, env }) {
  const kv = env.HITS;
  if (!kv) return json({ count: null }); // 바인딩 없으면 조용히 무시

  let n = parseInt((await kv.get(KEY)) || "0", 10);
  if (!Number.isFinite(n)) n = 0;

  if (request.method === "POST") {
    n += 1;
    await kv.put(KEY, String(n));
  }
  return json({ count: n });
}
