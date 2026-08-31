/*
 * 表格比对核心逻辑（浏览器 / Node.js 通用）
 * 规则：
 *   按用户指定的比对列，比较两个表格中的字符串（数字/字母，可能含 *）：
 *     7 位字符：前 2 后 2 相同即配对
 *     9 位字符：前 2 后 4 相同即配对
 *     10 位字符：前 2 后 5 相同即配对
 *     11 位字符：前 3 后 4 相同即配对
 *   配对成功后，把第一个表格中用户选定的列，粘贴到第二个表格对应行的右侧空列中；
 *   输出两个表：表1（来源）、表2（目标）。选择粘贴时两个表中配对成功的号码与粘贴信息标绿；
 *   未选择粘贴时仅两个表中配对成功的号码标绿。
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
   * 解析“x升级y”，返回预期含税价 y - x；无“升级”字样返回 null。
   */
  function upgradeDiff(cell) {
    const m = cellText(cell).match(/(\d+)\s*升级\s*(\d+)/);
    if (!m) return null;
    return Number(m[2]) - Number(m[1]);
  }

  function numbersEqual(a, b) {
    const na = Number(a);
    const nb = Number(b);
    if (!isFinite(na) || !isFinite(nb)) return false;
    return Math.abs(na - nb) < 1e-6;
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

    // 确认金额相关列（可有可无）
    const srcBizCol = findColumn(srcHeader, "业务");
    const tgtTaxCol = findColumn(tgtHeader, "含税价");
    const tgtBizCol = findColumn(tgtHeader, "业务");
    const srcTaxCol = findColumn(srcHeader, "含税价");

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
      if (!keyMap.has(k)) keyMap.set(k, { row, idx: r });
    }

    // 绿色高亮样式
    const GREEN = "A9D08E";
    const greenStyle = { fill: { patternType: "solid", fgColor: { rgb: GREEN } } };
    const confirmTgtCells = [];
    const confirmSrcCells = [];
    let amountConfirmedRows = 0;

    // 目标表最右侧数据列
    let maxCol = 0;
    for (let r = 0; r < tgtAoa.length; r++) {
      const row = tgtAoa[r] || [];
      for (let c = 0; c < row.length; c++) {
        if (!isEmpty(row[c])) maxCol = Math.max(maxCol, c);
      }
    }
    const newColIndex = maxCol + 1;

    const addRows = pasteIdx.length > 0 ? [pasteIdx.map((p) => p.name)] : null;
    const matchedRowIdx = [];
    const matchedKeys = new Set();
    let matched = 0;
    let unmatched = 0;
    const range = XLSX.utils.decode_range(tgt["!ref"]);
    const maxRow = Math.max(range.e.r, tgtAoa.length - 1);
    for (let r = 1; r <= maxRow; r++) {
      const row = tgtAoa[r] || [];
      if (!row.some((v) => !isEmpty(v))) {
        if (addRows) addRows.push(new Array(pasteIdx.length).fill(null));
        continue;
      }
      const key = matchKey(row[c2]);
      if (!key) {
        unmatched++;
        if (addRows) addRows.push(new Array(pasteIdx.length).fill(null));
        continue;
      }
      const srcEntry = keyMap.get(key.pre + "\u0000" + key.suf);
      if (srcEntry) {
        const srcRow = srcEntry.row;
        matched++;
        matchedRowIdx.push(r);
        matchedKeys.add(key.pre + "\u0000" + key.suf);
        if (addRows) addRows.push(pasteIdx.map((p) => (srcRow[p.idx] === undefined ? null : srcRow[p.idx])));
        let rowConfirmed = false;
        // 确认金额：表1 业务“x升级y”的预期含税价 与 表2 含税价 比较
        if (srcBizCol >= 0 && tgtTaxCol >= 0) {
          const expected = upgradeDiff(srcRow[srcBizCol]);
          if (expected !== null && numbersEqual(expected, row[tgtTaxCol])) {
            const addr = XLSX.utils.encode_cell({ r, c: tgtTaxCol });
            confirmTgtCells.push(addr);
            if (tgt[addr]) tgt[addr].s = greenStyle;
            rowConfirmed = true;
          }
        }
        // 反向：表2 业务 与 表1 含税价
        if (tgtBizCol >= 0 && srcTaxCol >= 0) {
          const expected = upgradeDiff(row[tgtBizCol]);
          if (expected !== null && numbersEqual(expected, srcRow[srcTaxCol])) {
            const addr = XLSX.utils.encode_cell({ r: srcEntry.idx, c: srcTaxCol });
            confirmSrcCells.push(addr);
            if (src[addr]) src[addr].s = greenStyle;
            rowConfirmed = true;
          }
        }
        if (rowConfirmed) amountConfirmedRows++;
      } else {
        unmatched++;
        if (addRows) addRows.push(new Array(pasteIdx.length).fill(null));
      }
    }

    if (addRows) {
      XLSX.utils.sheet_add_aoa(tgt, addRows, {
        origin: XLSX.utils.encode_cell({ r: 0, c: newColIndex }),
      });
    }

    // 表2（目标）：配对成功的行，比对列单元格与粘贴后不为空的单元格背景标为绿色
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
    for (const addr of confirmTgtCells) greenCells.push(addr);

    // 表1（来源）：被配对到的行，比对列单元格与选定的粘贴信息单元格标为绿色
    const srcGreenCells = [];
    const srcMatchedRowIdx = [];
    let srcMatchedRows = 0;
    for (let r = 1; r < srcAoa.length; r++) {
      const row = srcAoa[r] || [];
      if (!row.some((v) => !isEmpty(v))) continue;
      const key = matchKey(row[c1]);
      if (!key || !matchedKeys.has(key.pre + "\u0000" + key.suf)) continue;
      srcMatchedRows++;
      srcMatchedRowIdx.push(r);
      const cmpAddr = XLSX.utils.encode_cell({ r, c: c1 });
      srcGreenCells.push(cmpAddr);
      const cmpCell = src[cmpAddr];
      if (cmpCell) cmpCell.s = greenStyle;
      for (const p of pasteIdx) {
        const addr = XLSX.utils.encode_cell({ r, c: p.idx });
        const cell = src[addr];
        if (cell && !isEmpty(cell.v)) {
          srcGreenCells.push(addr);
          cell.s = greenStyle;
        }
      }
    }
    for (const addr of confirmSrcCells) srcGreenCells.push(addr);

    return {
      ok: true,
      matchedRows: matched,
      highlightedRows: matchedRowIdx.length,
      greenCells,
      srcGreenCells,
      srcMatchedRows,
      srcMatchedRowIdx,
      srcCompareColIndex: c1,
      srcPasteColIndexes: pasteIdx.map((p) => p.idx),
      amountConfirmedRows,
      confirmTgtCells,
      confirmSrcCells,
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
    upgradeDiff,
    numbersEqual,
    processCompare,
  };
});
