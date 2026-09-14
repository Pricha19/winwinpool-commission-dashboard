// POST /api/upload
// Body: {
//   monthDocs: { "YYYY-MM": { ym, label, totals, rows, statusCounts, totalRows, confirmedRows,
//                              sourceConfirmed, sourceNet, uploadedAt } , ... },
//   salesNames: { salesCode: name, ... },
//   uploaderName, invFileName, taxFileName, depFileName,
//   invRowCount, taxRowCount, depRowCount, unknownDateRows
// }
// All the heavy computation (parsing, classification, the DEP partial-deposit allocation, month
// bucketing) already happened client-side — this function only persists the already-computed
// result: one small JSON blob per month plus a merged "meta/index" summary document. It never
// receives or stores the raw uploaded Excel files.
import type { Context } from "@netlify/functions";
import { getStore } from "@netlify/blobs";

type MonthDoc = {
  ym: string;
  label?: string;
  totals?: Record<string, number>;
  rows?: unknown[];
  statusCounts?: Record<string, number>;
  totalRows?: number;
  confirmedRows?: number;
  sourceConfirmed?: Record<string, number>;
  sourceNet?: Record<string, number>;
  uploadedAt?: string;
};

type IndexDoc = {
  months: { ym: string; label: string; totalRows: number; confirmedRows: number; uploadedAt: string }[];
  salesNames: Record<string, string>;
  latestUpload: Record<string, unknown> | null;
};

const YM_RE = /^(\d{4}-\d{2}|unknown)$/;

export default async (req: Request, context: Context) => {
  if (req.method !== "POST") {
    return new Response(JSON.stringify({ error: "Method not allowed" }), {
      status: 405,
      headers: { "content-type": "application/json" },
    });
  }

  let body: any;
  try {
    body = await req.json();
  } catch {
    return new Response(JSON.stringify({ error: "รูปแบบข้อมูลที่ส่งมาไม่ใช่ JSON ที่ถูกต้อง" }), {
      status: 400,
      headers: { "content-type": "application/json" },
    });
  }

  const monthDocs = body?.monthDocs as Record<string, MonthDoc> | undefined;
  if (!monthDocs || typeof monthDocs !== "object" || Object.keys(monthDocs).length === 0) {
    return new Response(JSON.stringify({ error: "ไม่พบข้อมูลรายเดือนที่จะบันทึก (monthDocs ว่างเปล่า)" }), {
      status: 400,
      headers: { "content-type": "application/json" },
    });
  }

  const monthKeys = Object.keys(monthDocs);
  for (const ym of monthKeys) {
    if (!YM_RE.test(ym)) {
      return new Response(JSON.stringify({ error: "รูปแบบเดือนไม่ถูกต้อง: " + ym }), {
        status: 400,
        headers: { "content-type": "application/json" },
      });
    }
  }

  const incomingSalesNames = (body?.salesNames && typeof body.salesNames === "object" ? body.salesNames : {}) as Record<string, string>;

  try {
    const store = getStore("dashboard");
    const nowIso = new Date().toISOString();

    // 1) Write each month document. Each upload only ever touches the months present in the
    //    uploaded files — months not included here are left completely untouched.
    await Promise.all(
      monthKeys.map((ym) => store.setJSON(`months/${ym}`, { ...monthDocs[ym], ym, uploadedAt: nowIso }))
    );

    // 2) Read the existing index (if any) so we can upsert rather than clobber other months'
    //    listing entries and previously-known salesperson names.
    const existing = (await store.get("meta/index", { type: "json" })) as IndexDoc | null;
    const existingMonths = existing?.months || [];
    const existingSalesNames = existing?.salesNames || {};

    const monthsByYm: Record<string, IndexDoc["months"][number]> = {};
    existingMonths.forEach((m) => {
      monthsByYm[m.ym] = m;
    });
    monthKeys.forEach((ym) => {
      const doc = monthDocs[ym];
      monthsByYm[ym] = {
        ym,
        label: doc.label || ym,
        totalRows: doc.totalRows || 0,
        confirmedRows: doc.confirmedRows || 0,
        uploadedAt: nowIso,
      };
    });
    const mergedMonths = Object.values(monthsByYm).sort((a, b) => (a.ym < b.ym ? -1 : a.ym > b.ym ? 1 : 0));

    // New non-empty names overwrite old ones for the same code; a code missing from this
    // upload's payload keeps whatever name was already on file.
    const mergedSalesNames: Record<string, string> = { ...existingSalesNames };
    Object.keys(incomingSalesNames).forEach((code) => {
      const name = incomingSalesNames[code];
      if (name) mergedSalesNames[code] = name;
    });

    const latestUpload = {
      uploadedAt: nowIso,
      uploaderName: body?.uploaderName || "",
      invFileName: body?.invFileName || "",
      taxFileName: body?.taxFileName || "",
      depFileName: body?.depFileName || "",
      invRowCount: body?.invRowCount || 0,
      taxRowCount: body?.taxRowCount || 0,
      depRowCount: body?.depRowCount || 0,
      unknownDateRows: body?.unknownDateRows || 0,
      monthsUpdated: monthKeys,
    };

    const newIndex: IndexDoc = {
      months: mergedMonths,
      salesNames: mergedSalesNames,
      latestUpload,
    };

    // 3) Write the merged index back last, once both the month docs and the merge are ready.
    await store.setJSON("meta/index", newIndex);

    return new Response(JSON.stringify({ ok: true, monthsUpdated: monthKeys }), {
      headers: { "content-type": "application/json" },
    });
  } catch (err: any) {
    console.error("POST /api/upload failed:", err);
    return new Response(
      JSON.stringify({ error: "บันทึกข้อมูลไม่สำเร็จ: " + (err?.message || String(err)) }),
      { status: 500, headers: { "content-type": "application/json" } }
    );
  }
};
