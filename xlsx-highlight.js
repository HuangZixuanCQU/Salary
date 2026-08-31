/*
 * xlsx 绿色填充注入（浏览器 / Node.js 通用）
 * SheetJS 社区版不支持写出单元格填充样式，这里在导出后的 xlsx（zip）上
 * 直接向 xl/styles.xml 追加绿色 solid fill，并在对应工作表的单元格上写入样式引用。
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) {
    module.exports = factory(require("jszip"));
  } else {
    root.HighlightXlsx = factory(root.JSZip);
  }
})(typeof self !== "undefined" ? self : this, function (JSZip) {
  "use strict";

  const GREEN_ARGB = "FFA9D08E";

  /**
   * @param {ArrayBuffer|Uint8Array} bytes 由 SheetJS 生成的 xlsx 字节
   * @param {Object<string,string[]>} sheetRefsMap { 子表名: ["B2","D2",...] }
   * @returns {Promise<Uint8Array>} 注入绿色填充后的 xlsx 字节
   */
  async function highlightXlsx(bytes, sheetRefsMap) {
    const zip = await JSZip.loadAsync(bytes);

    // 子表名 -> 工作表 xml 路径（xl/worksheets/sheetN.xml）
    const wbXml = await zip.file("xl/workbook.xml").async("string");
    const relsXml = await zip.file("xl/_rels/workbook.xml.rels").async("string");
    const rels = {};
    const relRe = /<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g;
    let rm;
    while ((rm = relRe.exec(relsXml))) rels[rm[1]] = rm[2];
    const sheetEntries = [];
    const sheetRe = /<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"/g;
    let sm;
    while ((sm = sheetRe.exec(wbXml))) sheetEntries.push({ name: sm[1], rid: sm[2] });

    // 需要高亮的单元格集合
    const targets = [];
    for (const [sheetName, refs] of Object.entries(sheetRefsMap || {})) {
      if (!Array.isArray(refs) || refs.length === 0) continue;
      const entry = sheetEntries.find((e) => e.name === sheetName);
      if (!entry) continue;
      let target = rels[entry.rid] || "";
      if (target.startsWith("/")) target = target.slice(1);
      if (!target.startsWith("worksheets/")) target = "worksheets/" + target;
      targets.push({ path: "xl/" + target, refs });
    }
    if (targets.length === 0) return new Uint8Array(bytes);

    // 追加绿色 solid fill 和对应 cellXfs
    const stylesPath = "xl/styles.xml";
    let stylesXml = await zip.file(stylesPath).async("string");
    const fillsMatch = stylesXml.match(/<fills count="(\d+)">/);
    const xfsMatch = stylesXml.match(/<cellXfs count="(\d+)">/);
    if (!fillsMatch || !xfsMatch) return new Uint8Array(bytes);

    const fillCount = parseInt(fillsMatch[1], 10);
    const xfCount = parseInt(xfsMatch[1], 10);
    const fillId = fillCount;
    const xfId = xfCount;
    const fillXml = '<fill><patternFill patternType="solid"><fgColor rgb="' + GREEN_ARGB + '"/><bgColor indexed="64"/></patternFill></fill>';
    const xfXml = '<xf numFmtId="0" fontId="0" fillId="' + fillId + '" borderId="0" xfId="0" applyFill="1"/>';
    stylesXml = stylesXml.replace(/<fills count="\d+">/, '<fills count="' + (fillCount + 1) + '">');
    stylesXml = stylesXml.replace(/<\/fills>/, fillXml + "</fills>");
    stylesXml = stylesXml.replace(/<cellXfs count="\d+">/, '<cellXfs count="' + (xfCount + 1) + '">');
    stylesXml = stylesXml.replace(/<\/cellXfs>/, xfXml + "</cellXfs>");
    zip.file(stylesPath, stylesXml);

    // 给目标单元格写入样式引用 s="xfId"
    for (const t of targets) {
      let sheetXml = await zip.file(t.path).async("string");
      for (const ref of t.refs) {
        const re = new RegExp('<c r="' + ref + '"([^>]*)>', "g");
        sheetXml = sheetXml.replace(re, (all, rest) => {
          if (rest.indexOf(' s="') !== -1) return all;
          return '<c r="' + ref + '" s="' + xfId + '"' + rest + ">";
        });
      }
      zip.file(t.path, sheetXml);
    }

    return await zip.generateAsync({ type: "uint8array" });
  }

  return { highlightXlsx, GREEN_ARGB };
});
