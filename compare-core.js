/*
 * 表格比对核心逻辑（浏览器 / Node.js 通用）
 * 规则：
 *   按用户指定的比对列，比较两个表格中的字符串（数字/字母，可能含 *）：
 *     7 位字符：前 2 后 2 相同即配对
 *     9 位字符：前 2 后 4 相同即配对
 *     10 位字符：前 2 后 5 相同即配对
 *     11 位字符：前 3 后 4 相同即配对
 *   配对成功后，把第一个表格中用户选定的列，粘贴到第二个表格对应行的右侧空列中；
 *   并将配对成功的“比对号码”单元格以及粘贴后不为空的单元格背景标为绿色。
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) {
    module.exports = factory(require("xlsx"));
  } else {
    root.CompareCore = factory(root.XLSX);
  }
})(typeof self !== "undefined" ? self : this, function (XLSX) {
  "use strict";

  function cellText(v) {
    if (v === null || v === undefined) return "";
    return String(v);
  }

  function isEmpty(v) {
    return v === null || v === undefined || cellText(v).trim() === "";
  }

  /**
   * 按字符数量提取前缀/后缀匹配键；不支持的长度返回 null。
   */
  function matchKey(value) {
    const s = cellText(value).trim();
    const len = s.length;
    if (len === 7) return { pre: s.slice(0, 2), suf: s.slice(-2) };
    if (len === 9) return { pre: s.slice(0, 2), suf: s.slice(-4) };
    if (len === 10) return { pre: s.slice(0, 2), suf: s.slice(-5) };
    if (len === 11) return { pre: s.slice(0, 3), suf: s.slice(-4) };
    return null;
  }

  function keysMatch(a, b) {
    const ka = matchKey(a);
    const kb = matchKey(b);
    return !!(ka && kb && ka.pre === kb.pre && ka.suf === kb.suf);
  }

  /**
   * 查找列：先精确匹配，再忽略大小写，再包含匹配。
   */
  function findColumn(header, name) {
    const target = String(name || "").trim();
    if (!target) return -1;
    for (let c = 0; c < header.length; c++) {
      if (cellText(header[c]).trim() === target) return c;
    }
    for (let c = 0; c < header.length; c++) {
      if (cellText(header[c]).trim().toLowerCase() === target.toLowerCase()) return c;
    }
    for (let c = 0; c < header.length; c++) {
      if (cellText(header[c]).trim().indexOf(target) !== -1) return c;
    }
    return -1;
  }

  /**
   * 处理比对：把 srcWb 指定子表中选定的列，按匹配规则粘贴到 tgtWb 指定子表的右侧空列。
   * 直接修改 tgtWb，并返回处理摘要。
   */
  function processCompare(srcWb, srcSheetName, compareCol1, tgtWb, tgtSheetName, compareCol2, pasteCols) {
    const src = srcWb.Sheets[srcSheetName];
    const tgt = tgtWb.Sheets[tgtSheetName];
    if (!src || !src["!ref"]) return { ok: false, error: "第一个表格的子表为空，请重新选择" };
    if (!tgt || !tgt["!ref"]) return { ok: false, error: "第二个表格的子表为空，请重新选择" };

    const srcAoa = XLSX.utils.sheet_to_json(src, { header: 1, defval: null, raw: true });
    const tgtAoa = XLSX.utils.sheet_to_json(tgt, { header: 1, defval: null, raw: true });
    const srcHeader = srcAoa[0] || [];
    const tgtHeader = tgtAoa[0] || [];

    const c1 = findColumn(srcHeader, compareCol1);
    const c2 = findColumn(tgtHeader, compareCol2);
    if (c1 < 0) return { ok: false, error: "第一个表格中未找到比对列“" + compareCol1 + "”" };
    if (c2 < 0) return { ok: false, error: "第二个表格中未找到比对列“" + compareCol2 + "”" };

    // 需要粘贴的列（来自第一个表格）
    const wanted = (Array.isArray(pasteCols) ? pasteCols : [])
      .map((s) => String(s || "").trim())
      .filter(Boolean);
    const pasteIdx = [];
    const skipped = [];
    for (const w of wanted) {
      const idx = findColumn(srcHeader, w);
      if (idx >= 0) pasteIdx.push({ name: cellText(srcHeader[idx]).trim(), idx });
      else skipped.push(w);
    }
    if (pasteIdx.length === 0) return { ok: false, error: "第一个表格中未找到任何需要粘贴的列名" };

    // 建立源表匹配键 -> 行（同键多行取第一行）
    const keyMap = new Map();
    let srcRowsWithKey = 0;
    for (let r = 1; r < srcAoa.length; r++) {
      const row = srcAoa[r] || [];
      if (!row.some((v) => !isEmpty(v))) continue;
      const key = matchKey(row[c1]);
      if (!key) continue;
      srcRowsWithKey++;
      const k = key.pre + "\u0000" + key.suf;
      if (!keyMap.has(k)) keyMap.set(k, row);
    }

    // 目标表最右侧数据列
    let maxCol = 0;
    for (let r = 0; r < tgtAoa.length; r++) {
      const row = tgtAoa[r] || [];
      for (let c = 0; c < row.length; c++) {
        if (!isEmpty(row[c])) maxCol = Math.max(maxCol, c);
      }
    }
    const newColIndex = maxCol + 1;

    const addRows = [pasteIdx.map((p) => p.name)];
    const matchedRowIdx = [];
    let matched = 0;
    let unmatched = 0;
    const range = XLSX.utils.decode_range(tgt["!ref"]);
    const maxRow = Math.max(range.e.r, tgtAoa.length - 1);
    for (let r = 1; r <= maxRow; r++) {
      const row = tgtAoa[r] || [];
      if (!row.some((v) => !isEmpty(v))) {
        addRows.push(new Array(pasteIdx.length).fill(null));
        continue;
      }
      const key = matchKey(row[c2]);
      if (!key) {
        unmatched++;
        addRows.push(new Array(pasteIdx.length).fill(null));
        continue;
      }
      const srcRow = keyMap.get(key.pre + "\u0000" + key.suf);
      if (srcRow) {
        matched++;
        matchedRowIdx.push(r);
        addRows.push(pasteIdx.map((p) => (srcRow[p.idx] === undefined ? null : srcRow[p.idx])));
      } else {
        unmatched++;
        addRows.push(new Array(pasteIdx.length).fill(null));
      }
    }

    XLSX.utils.sheet_add_aoa(tgt, addRows, {
      origin: XLSX.utils.encode_cell({ r: 0, c: newColIndex }),
    });

    // 配对成功的行：比对列单元格与粘贴后不为空的单元格背景标为绿色
    const GREEN = "C6EFCE";
    const greenStyle = { fill: { patternType: "solid", fgColor: { rgb: GREEN } } };
    const greenCells = [];
    for (const r of matchedRowIdx) {
      const cmpAddr = XLSX.utils.encode_cell({ r, c: c2 });
      greenCells.push(cmpAddr);
      const cmpCell = tgt[cmpAddr];
      if (cmpCell) cmpCell.s = greenStyle;
      for (let i = 0; i < pasteIdx.length; i++) {
        const addr = XLSX.utils.encode_cell({ r, c: newColIndex + i });
        const cell = tgt[addr];
        if (cell && !isEmpty(cell.v)) {
          greenCells.push(addr);
          cell.s = greenStyle;
        }
      }
    }

    return {
      ok: true,
      matchedRows: matched,
      highlightedRows: matchedRowIdx.length,
      greenCells,
      unmatchedRows: unmatched,
      targetDataRows: matched + unmatched,
      pastedColumns: pasteIdx.map((p) => p.name),
      skippedColumns: skipped,
      srcRowsWithKey,
      newColIndex,
    };
  }

  return {
    matchKey,
    keysMatch,
    findColumn,
    processCompare,
  };
});
