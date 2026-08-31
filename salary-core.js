/*
 * Salary 计算核心逻辑（浏览器 / Node.js 通用）
 * 规则：
 *   “业务”列：计算前先按税率扣除税收：基数 × (1 - 税率)
 *   1. 单元格中恰好 1 个数字且文本包含“新装”：Salary 1 = 数字 × 比例；Salary 2 = 数字 × (1 - 比例)
 *   2. 恰好 2 个数字且文本包含“新装”：取“新装”后面的数字 n，Salary 1 = n × 比例；Salary 2 = n × (1 - 比例)
 *   3. 恰好 2 个数字且文本包含“升级”：Salary 1 = (大数 - 小数) × 比例；Salary 2 = (大数 - 小数) × (1 - 比例)
 *   4. 其他情况（3 个及以上数字、2 个数字但无“新装”和“升级”、1 个数字但无“新装”、没有数字）：两个单元格均输出“不符合格式，未计算”
 *   “利润额”列：单个数字（正数不带正号、负数带负号），Salary 1 = 数字 × 比例；Salary 2 = 数字 × (1 - 比例)；否则输出“格式不符，未计算”
 *   第二阶段（可选）：按“受理人”列每行“+”的个数 n（0、1、2、3）平分 Salary 1：Salary 1 ÷ (n+1)，追加在 Salary 2 右侧，列名“Salary 1平分”
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) {
    module.exports = factory(require("xlsx"));
  } else {
    root.SalaryCore = factory(root.XLSX);
  }
})(typeof self !== "undefined" ? self : this, function (XLSX) {
  "use strict";

  const BAD = "不符合格式，未计算";
  const BAD_PROFIT = "格式不符，未计算";
  const HEADER1 = "Salary 1";
  const HEADER2 = "Salary 2";
  const STAGE_HEADER = "Salary 1平分";

  function cellText(v) {
    if (v === null || v === undefined) return "";
    return String(v);
  }

  function extractNumbers(text) {
    const s = cellText(text);
    const m = s.match(/\d+/g);
    return m ? m.map(Number) : [];
  }

  function computeRow(cell, ratio, taxRate) {
    const text = cellText(cell).trim();
    const nums = extractNumbers(text);
    const taxFactor = 1 - (Number.isFinite(Number(taxRate)) ? Math.min(Math.max(Number(taxRate), 0), 1) : 0);
    if (nums.length === 0) return { ok: false };
    if (nums.length === 1) {
      if (text.indexOf("新装") !== -1) {
        return { ok: true, s1: round2(nums[0] * taxFactor * ratio), s2: round2(nums[0] * taxFactor * (1 - ratio)) };
      }
      return { ok: false };
    }
    if (nums.length === 2 && text.indexOf("新装") !== -1) {
      const m = text.match(/新装\s*(\d+)/);
      if (m) {
        const n = Number(m[1]);
        return { ok: true, s1: round2(n * taxFactor * ratio), s2: round2(n * taxFactor * (1 - ratio)) };
      }
      return { ok: false };
    }
    if (nums.length === 2 && text.indexOf("升级") !== -1) {
      const diff = Math.abs(nums[0] - nums[1]);
      return { ok: true, s1: round2(diff * taxFactor * ratio), s2: round2(diff * taxFactor * (1 - ratio)) };
    }
    return { ok: false };
  }

  const SIGNED_NUM_RE = /^[+-]?\d+(?:\.\d+)?$/;

  /**
   * “利润额”列规则：单元格必须是单个数字（正数不带正号如 120，负数带负号如 -35.5）。
   * 符合格式：Salary 1 = 数字 × 比例；Salary 2 = 数字 × (1 - 比例)。
   */
  function computeProfitRow(cell, ratio) {
    const text = cellText(cell).trim();
    if (!SIGNED_NUM_RE.test(text)) return { ok: false };
    const n = parseFloat(text);
    return { ok: true, s1: round2(n * ratio), s2: round2(n * (1 - ratio)) };
  }

  function isEmptyValue(v) {
    return v === null || v === undefined || cellText(v).trim() === "";
  }

  function round2(x) {
    return Math.round((x + Number.EPSILON) * 100) / 100;
  }

  /**
   * 处理工作簿：对指定子表查找目标列名，若找到则在最右侧空列追加 Salary 1 / Salary 2。
   * @param {object} wb 工作簿
   * @param {string} columnName 需要处理的列名
   * @param {number} ratio 比例（0 ~ 1）
   * @param {string[]} [sheetNames] 需要处理的子表名列表；不传或空数组则处理全部子表
   * @param {string} [stage] 第二阶段分配方式（“双人5-5”/“三人3-3-3”/“四人平分”），不传则不进行第二阶段
   * @returns 每个子表的处理摘要
   */
  function processWorkbook(wb, columnName, ratio, sheetNames, stage, taxRate) {
    const target = String(columnName || "").trim();
    const stageEnabled = typeof stage === "string" && stage.trim() !== "" && stage !== "none";
    let mode;
    if (target === "利润额") mode = "profit";
    else if (target === "业务") mode = "business";
    else throw new Error("仅支持处理“业务”或“利润额”列名");
    const badText = mode === "profit" ? BAD_PROFIT : BAD;
    const summaries = [];
    const requested = Array.isArray(sheetNames)
      ? sheetNames.map((s) => String(s || "").trim()).filter((s) => s !== "")
      : [];

    // 用户指定的子表名若不存在，先给出提示条目
    if (requested.length > 0) {
      for (const name of requested) {
        const exists = wb.SheetNames.some((n) => n.toLowerCase() === name.toLowerCase());
        if (!exists) {
          summaries.push({
            sheet: name,
            missing: true,
            found: false,
            columnIndex: -1,
            newColIndex: -1,
            rowsProcessed: 0,
            rowsBad: 0,
          });
        }
      }
    }

    for (const sheetName of wb.SheetNames) {
      if (requested.length > 0) {
        const selected = requested.some((r) => r.toLowerCase() === sheetName.toLowerCase());
        if (!selected) {
          summaries.push({
            sheet: sheetName,
            skipped: true,
            found: false,
            columnIndex: -1,
            newColIndex: -1,
            rowsProcessed: 0,
            rowsBad: 0,
          });
          continue;
        }
      }

      const sheet = wb.Sheets[sheetName];
      const summary = {
        sheet: sheetName,
        found: false,
        columnIndex: -1,
        newColIndex: -1,
        rowsProcessed: 0,
        rowsBad: 0,
      };

      if (!sheet || !sheet["!ref"]) {
        summaries.push(summary);
        continue;
      }

      const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: true });
      const header = aoa[0] || [];

      // 查找列名（先精确匹配，再忽略大小写匹配）
      let colIdx = -1;
      for (let c = 0; c < header.length; c++) {
        const h = cellText(header[c]).trim();
        if (h === target || h.toLowerCase() === target.toLowerCase()) {
          colIdx = c;
          break;
        }
      }
      summary.found = colIdx >= 0;
      summary.columnIndex = colIdx;
      if (!summary.found) {
        summaries.push(summary);
        continue;
      }

      // 最右侧数据列
      let maxCol = 0;
      for (let r = 0; r < aoa.length; r++) {
        const row = aoa[r] || [];
        for (let c = 0; c < row.length; c++) {
          if (!isEmptyValue(row[c])) maxCol = Math.max(maxCol, c);
        }
      }
      const newColIndex = maxCol + 1;
      summary.newColIndex = newColIndex;

      const range = XLSX.utils.decode_range(sheet["!ref"]);
      const maxRow = Math.max(range.e.r, aoa.length - 1);
      // 第二阶段需要“受理人”列
      let agentColIdx = -1;
      if (stageEnabled) {
        for (let c = 0; c < header.length; c++) {
          const h = cellText(header[c]).trim();
          if (h === "受理人" || h.toLowerCase() === "受理人") {
            agentColIdx = c;
            break;
          }
        }
      }
      const addRows = [[HEADER1, HEADER2, ...(stageEnabled ? [STAGE_HEADER] : [])]];

      for (let r = 1; r <= maxRow; r++) {
        const row = aoa[r] || [];
        const hasAny = row.some((v) => !isEmptyValue(v));
        if (!hasAny) {
          addRows.push(stageEnabled ? [null, null, null] : [null, null]);
          continue;
        }
        summary.rowsProcessed++;
        const res = mode === "profit" ? computeProfitRow(row[colIdx], ratio) : computeRow(row[colIdx], ratio, taxRate);
        if (!res.ok) {
          summary.rowsBad++;
          addRows.push(stageEnabled ? [badText, badText, badText] : [badText, badText]);
        } else if (stageEnabled) {
          let split;
          if (agentColIdx < 0) {
            split = "未找到受理人列，未计算";
          } else {
            const agentText = cellText(row[agentColIdx]);
            const plusCount = (agentText.match(/\+/g) || []).length;
            split = round2(res.s1 / (plusCount + 1));
          }
          addRows.push([res.s1, res.s2, split]);
        } else {
          addRows.push([res.s1, res.s2]);
        }
      }

      XLSX.utils.sheet_add_aoa(sheet, addRows, {
        origin: XLSX.utils.encode_cell({ r: 0, c: newColIndex }),
      });
      // 新增的数值单元格显示两位小数
      for (let r = 1; r < addRows.length; r++) {
        const lastNewCol = newColIndex + (stageEnabled ? 2 : 1);
        for (let c = newColIndex; c <= lastNewCol; c++) {
          const addr = XLSX.utils.encode_cell({ r, c });
          const cell = sheet[addr];
          if (cell && typeof cell.v === "number") cell.z = "0.00";
        }
      }
      summary.stageLabel = stageEnabled ? STAGE_HEADER : null;
      summary.stageColIndex = stageEnabled ? newColIndex + 2 : -1;
      summary.stageMissingAgent = stageEnabled && agentColIdx < 0;
      summaries.push(summary);
    }
    return summaries;
  }

  return {
    BAD,
    BAD_PROFIT,
    HEADER1,
    HEADER2,
    STAGE_HEADER,
    extractNumbers,
    computeRow,
    computeProfitRow,
    processWorkbook,
  };
});
