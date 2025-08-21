(async () => {
  // === 查詢條件 ===
  const ACC_ID = "CBA_MS_ID";
  const CODE1  = "05"; // 資訊安全管理系統
  const CODE2  = "01"; // 資訊安全管理系統驗證方案
  const LANG   = "zh_TW";
  const PAGE_SIZE = 50;        // 每頁 50 筆
  const PAGE_DELAY_MS = 250;   // 每頁間隔（避免過於頻繁）
  const DETAIL_CONCURRENCY = 8; // 詳情並發抓取數
  const DETAIL_DELAY_MS = 100;  // 並發中的最小間隔

  const BASE = location.origin;
  const LIST_API   = `${BASE}/system/modules/com.thesys.project.taf/pages/ajax/cbaclicent-list-process.jsp`;
  const DETAIL_API = `${BASE}/system/modules/com.thesys.project.taf/pages/ajax/cbaclicent-detail-process.jsp`;

  const sleep = (ms) => new Promise(r => setTimeout(r, ms));

  async function getJSON(url, params) {
    const qs = new URLSearchParams(params);
    const res = await fetch(`${url}?${qs.toString()}`, {
      method: "GET",
      headers: { "X-Requested-With": "XMLHttpRequest", "Accept": "application/json" },
      credentials: "same-origin",
    });
    if (!res.ok) {
      const t = await res.text().catch(() => "");
      throw new Error(`HTTP ${res.status} @ ${url} :: ${t.slice(0,200)}`);
    }
    return res.json();
  }

  // 扁平化多語欄位：{ zh_TW, en } -> 兩欄；純字串則保留原欄
  function flattenLangKV(key, val) {
    const out = {};
    if (val && typeof val === "object" && ("zh_TW" in val || "en" in val)) {
      out[`${key}_zh_TW`] = val.zh_TW ?? "";
      out[`${key}_en`]    = val.en ?? "";
    } else {
      out[key] = val ?? "";
    }
    return out;
  }

  function mergeFlatten(listItem = {}, detail = {}) {
    const row = {};
    const langKeys = [
      "custName", "institutionName", "companyAddr", "note",
      "otherPlaceName", "otherPlaceAddr", "otherPlaceRange"
    ];
    for (const k of langKeys) Object.assign(row, flattenLangKV(k, listItem[k]));
    row.verifyAccredit      = listItem.verifyAccredit ?? "";
    row.initialAccreditDate = (listItem.initialAccreditDate ?? "").slice(0,10);
    row.accreditValidDate   = (listItem.accreditValidDate ?? "").slice(0,10);
    row.tel                 = listItem.tel ?? "";
    row.certificateNo       = listItem.certificateNo ?? "";
    row.uuid                = listItem.uuid ?? "";

    // 詳情補齊（若為多語物件也拆欄；日期欄位只取 yyyy-mm-dd）
    for (const [k, v] of Object.entries(detail || {})) {
      if (v && typeof v === "object" && ("zh_TW" in v || "en" in v)) {
        Object.assign(row, flattenLangKV(k, v));
      } else if (typeof v === "string") {
        row[k] = /date/i.test(k) ? v.slice(0,10) : v;
      } else if (typeof v !== "object") {
        row[k] = v ?? "";
      }
    }
    return row;
  }

  function toCSV(rows) {
    const header = new Set();
    rows.forEach(r => Object.keys(r).forEach(k => header.add(k)));
    const headers = Array.from(header);
    const esc = (s) => {
      const str = (s ?? "").toString();
      return /[",\n]/.test(str) ? `"${str.replace(/"/g,'""')}"` : str;
    };
    const lines = [headers.map(esc).join(",")];
    for (const r of rows) lines.push(headers.map(h => esc(r[h])).join(","));
    return lines.join("\n");
  }

  async function mapLimit(arr, limit, mapper) {
    const ret = new Array(arr.length);
    let i = 0, active = 0, rejectOnce;
    return new Promise((resolve, reject) => {
      rejectOnce = reject;
      const next = () => {
        if (i >= arr.length && active === 0) return resolve(ret);
        while (active < limit && i < arr.length) {
          const idx = i++;
          active++;
          Promise.resolve()
            .then(() => mapper(arr[idx], idx))
            .then(v => { ret[idx] = v; })
            .catch(rejectOnce)
            .finally(async () => { active--; if (DETAIL_DELAY_MS) await sleep(DETAIL_DELAY_MS); next(); });
        }
      };
      next();
    });
  }

  console.time("[TAF] 抓取");
  console.log("[TAF] 讀取第 1 頁（每頁 50 筆）…");

  // 1) 先拿第 1 頁，確認 pageCount / queryCount
  const baseParams = {
    accId: ACC_ID, code1: CODE1, code2: CODE2, custCname: "",
    lang: LANG, pageSize: PAGE_SIZE, pageIndex: 1
  };
  const first = await getJSON(LIST_API, baseParams);
  if (!first || !Array.isArray(first.itemList)) throw new Error("首頁回傳格式異常或無資料。");
  const pageCount = Number(first.pageCount || 1);
  const queryCount = Number(first.queryCount || 0);
  console.log(`[TAF] 總筆數約 ${queryCount}，總頁數 ${pageCount}。`);

  // 2) 逐頁抓清單（從第 2 頁開始；第 1 頁已在 first）
  const items = [...first.itemList];
  for (let p = 2; p <= pageCount; p++) {
    const data = await getJSON(LIST_API, { ...baseParams, pageIndex: p });
    const list = Array.isArray(data.itemList) ? data.itemList : [];
    items.push(...list);
    console.log(`[TAF] 第 ${p}/${pageCount} 頁 -> 收到 ${list.length} 筆，累計 ${items.length}`);
    if (p < pageCount) await sleep(PAGE_DELAY_MS);
  }

  // 去重（以 uuid）
  const seen = new Set();
  const uniq = [];
  for (const it of items) {
    const id = it?.uuid;
    if (!id || seen.has(id)) continue;
    seen.add(id);
    uniq.push(it);
  }
  console.log(`[TAF] 清單總筆數：${uniq.length}（已去重）`);

  // 3) 詳情抓取並合併
  console.log("[TAF] 下載詳情並合併欄位…");
  const merged = await mapLimit(uniq, DETAIL_CONCURRENCY, async (it) => {
    let detail = {};
    if (it.uuid) {
      try {
        detail = await getJSON(DETAIL_API, { uuid: it.uuid, lang: LANG });
      } catch (e) {
        console.warn("[TAF] 詳情失敗：", it.uuid, e?.message || e);
      }
    }
    return mergeFlatten(it, detail);
  });

  // 4) 生成並下載 CSV
  const csv = toCSV(merged);
  const filename = `TAF_${ACC_ID}_C${CODE1}_${CODE2}_full_${new Date().toISOString().slice(0,10)}.csv`;
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = Object.assign(document.createElement("a"), { href: url, download: filename });
  document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);

  console.timeEnd("[TAF] 抓取");
  console.log(`[TAF] 完成，輸出 ${merged.length} 筆 -> ${filename}`);
})();
