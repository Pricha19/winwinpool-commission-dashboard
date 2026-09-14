// GET /api/months?from=YYYY-MM&to=YYYY-MM   -> only months in [from, to] (inclusive, string compare)
// GET /api/months?all=1                     -> every month that has data
// Returns { "2026-09": {...monthDoc}, "2026-10": {...monthDoc} } — only months that both exist
// in the index AND are within the requested range are included; a missing/empty index yields {}.
import type { Context } from "@netlify/functions";
import { getStore } from "@netlify/blobs";

export default async (req: Request, context: Context) => {
  if (req.method !== "GET") {
    return new Response(JSON.stringify({ error: "Method not allowed" }), {
      status: 405,
      headers: { "content-type": "application/json" },
    });
  }

  try {
    const store = getStore("dashboard");
    const url = new URL(req.url);
    const all = url.searchParams.get("all") === "1";
    const from = url.searchParams.get("from") || "";
    const to = url.searchParams.get("to") || "";

    const idx = (await store.get("meta/index", { type: "json" })) as
      | { months?: { ym: string }[] }
      | null;
    const allMonths: string[] = (idx?.months || []).map((m) => m.ym);
    const wanted = all ? allMonths : allMonths.filter((ym) => ym >= from && ym <= to);

    const result: Record<string, unknown> = {};
    await Promise.all(
      wanted.map(async (ym) => {
        const doc = await store.get(`months/${ym}`, { type: "json" });
        if (doc) result[ym] = doc;
      })
    );

    return new Response(JSON.stringify(result), {
      headers: { "content-type": "application/json" },
    });
  } catch (err: any) {
    console.error("GET /api/months failed:", err);
    return new Response(
      JSON.stringify({ error: "อ่านข้อมูลรายเดือนไม่สำเร็จ: " + (err?.message || String(err)) }),
      { status: 500, headers: { "content-type": "application/json" } }
    );
  }
};
