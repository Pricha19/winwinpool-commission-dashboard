// GET /api/index
// Returns the small "meta/index" document: the list of months that have data,
// the shared salesperson-name directory, and info about the most recent upload.
// If nothing has ever been uploaded, returns an empty-but-valid shape instead of 404,
// so the frontend can render its "no data yet" state without special-casing errors.
import type { Context } from "@netlify/functions";
import { getStore } from "@netlify/blobs";

const EMPTY_INDEX = { months: [], salesNames: {}, latestUpload: null };

export default async (req: Request, context: Context) => {
  if (req.method !== "GET") {
    return new Response(JSON.stringify({ error: "Method not allowed" }), {
      status: 405,
      headers: { "content-type": "application/json" },
    });
  }

  try {
    const store = getStore("dashboard");
    const idx = await store.get("meta/index", { type: "json" });
    return new Response(JSON.stringify(idx || EMPTY_INDEX), {
      headers: { "content-type": "application/json" },
    });
  } catch (err: any) {
    console.error("GET /api/index failed:", err);
    return new Response(
      JSON.stringify({ error: "อ่านข้อมูลดัชนีไม่สำเร็จ: " + (err?.message || String(err)) }),
      { status: 500, headers: { "content-type": "application/json" } }
    );
  }
};
