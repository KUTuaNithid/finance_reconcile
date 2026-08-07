// ======================================================
// GL & TB Reconciliation Tool — Full Feature JavaScript
// Ported from Python Streamlit app.py (~2050 lines)
// ======================================================

// ======================================================
// 0. DEFAULT CONFIGURATION (from fs_config.json)
// ======================================================
const DEFAULT_CONFIG = {
  fs_line_items: {
    "เงินสดและรายการเทียบเท่าเงินสด": { prefixes: ["111"], keywords: ["เงินสด", "เงินฝาก"], side: "bs_debit" },
    "ลูกหนี้การค้าและลูกหนี้หมุนเวียนอื่น": { prefixes: ["113", "1153-01"], keywords: ["ลูกหนี้"], side: "bs_debit" },
    "เงินให้กู้ยืมแก่บุคคลที่เกี่ยวข้องกัน": { prefixes: ["121"], keywords: ["เงินให้กู้ยืม"], side: "bs_debit" },
    "สินทรัพย์หมุนเวียนอื่น": { prefixes: ["115", "119", "150"], keywords: ["ภาษีถูกหัก", "จ่ายล่วงหน้า", "ดอกเบี้ยค้างรับ"], side: "bs_debit" },
    "อุปกรณ์-สุทธิ": { prefixes: ["141"], keywords: ["เครื่องมือ", "เครื่องจักร", "อุปกรณ์สำนักงาน"], side: "bs_debit" },
    "ค่าเสื่อมราคาสะสม": { prefixes: ["142"], keywords: ["ค่าเสื่อมราคาสะสม"], side: "bs_credit" },
    "เจ้าหนี้การค้าและเจ้าหนี้หมุนเวียนอื่น": { prefixes: ["211", "212", "2131"], keywords: ["เจ้าหนี้การค้า", "เจ้าหนี้", "ค้างจ่าย", "สอบบัญชี", "ทำบัญชี"], side: "bs_credit" },
    "หนี้สินหมุนเวียนอื่น": { prefixes: ["2132", "2137", "2131-04"], keywords: ["ภาษีหัก", "ภงด", "กรมสรรพากร", "ประกันสังคม", "รอนำส่ง"], side: "bs_credit" },
    "เงินกู้ยืมจากบุคคลที่เกี่ยวข้องกัน": { prefixes: ["2138"], keywords: ["เงินกู้ยืม"], side: "bs_credit" },
    "ทุนเรือนหุ้น": { prefixes: ["31"], keywords: ["ทุน"], side: "bs_credit" },
    "กำไร(ขาดทุน)สะสม": { prefixes: ["32"], keywords: ["กำไร"], side: "bs_debit" },
    "รายได้จากการให้บริการ": { prefixes: ["41"], keywords: ["รายได้จากการ"], side: "pl_credit" },
    "รายได้อื่น": { prefixes: ["42"], keywords: ["รายได้อื่น", "ดอกเบี้ยรับ"], side: "pl_credit" },
    "ต้นทุนการให้บริการ": { prefixes: ["51", "5200-11"], keywords: ["ต้นทุน", "ซื้อ", "ค่าจ้าง"], side: "pl_debit" },
    "ค่าใช้จ่ายในการขายและบริหาร": { prefixes: ["52", "53"], keywords: ["ค่าใช้จ่าย", "เงินเดือน", "ค่าธรรมเนียม", "ค่าเสื่อม", "ค่าเช่า", "ประกัน", "สอบบัญชี", "บริการ"], side: "pl_debit" },
    "ต้นทุนทางการเงิน": { prefixes: ["54", "55"], keywords: ["ดอกเบี้ยจ่าย", "ต้นทุนทางการเงิน"], side: "pl_debit" },
    "ภาษีเงินได้": { prefixes: [], keywords: ["ภาษีเงินได้นิติบุคคล"], side: "pl_debit" },
  },
  bs_structure: [
    ["header", "สินทรัพย์", null, null],
    ["header", "สินทรัพย์หมุนเวียน", null, null],
    ["item", "เงินสดและรายการเทียบเท่าเงินสด", 4, "เงินสดและรายการเทียบเท่าเงินสด"],
    ["item", "ลูกหนี้การค้าและลูกหนี้หมุนเวียนอื่น", 5, "ลูกหนี้การค้าและลูกหนี้หมุนเวียนอื่น"],
    ["item", "เงินให้กู้ยืมแก่บุคคลที่เกี่ยวข้องกัน", null, "เงินให้กู้ยืมแก่บุคคลที่เกี่ยวข้องกัน"],
    ["item", "สินทรัพย์หมุนเวียนอื่น", 6, "สินทรัพย์หมุนเวียนอื่น"],
    ["subtotal", "รวมสินทรัพย์หมุนเวียน", null, ["เงินสดและรายการเทียบเท่าเงินสด", "ลูกหนี้การค้าและลูกหนี้หมุนเวียนอื่น", "เงินให้กู้ยืมแก่บุคคลที่เกี่ยวข้องกัน", "สินทรัพย์หมุนเวียนอื่น"]],
    ["spacer", null, null, null],
    ["header", "สินทรัพย์ไม่หมุนเวียน", null, null],
    ["item", "อุปกรณ์-สุทธิ", 7, ["อุปกรณ์-สุทธิ", "-ค่าเสื่อมราคาสะสม"]],
    ["subtotal", "รวมสินทรัพย์ไม่หมุนเวียน", null, ["อุปกรณ์-สุทธิ"]],
    ["subtotal", "รวมสินทรัพย์", null, ["รวมสินทรัพย์หมุนเวียน", "รวมสินทรัพย์ไม่หมุนเวียน"]],
    ["spacer", null, null, null],
    ["header", "หนี้สินและส่วนของเจ้าของ", null, null],
    ["header", "หนี้สินหมุนเวียน", null, null],
    ["item", "เจ้าหนี้การค้าและเจ้าหนี้หมุนเวียนอื่น", 8, "เจ้าหนี้การค้าและเจ้าหนี้หมุนเวียนอื่น"],
    ["item", "เงินกู้ยืมจากบุคคลที่เกี่ยวข้องกัน", 9, "เงินกู้ยืมจากบุคคลที่เกี่ยวข้องกัน"],
    ["item", "หนี้สินหมุนเวียนอื่น", 10, "หนี้สินหมุนเวียนอื่น"],
    ["subtotal", "รวมหนี้สินหมุนเวียน", null, ["เจ้าหนี้การค้าและเจ้าหนี้หมุนเวียนอื่น", "เงินกู้ยืมจากบุคคลที่เกี่ยวข้องกัน", "หนี้สินหมุนเวียนอื่น"]],
    ["subtotal", "รวมหนี้สิน", null, ["รวมหนี้สินหมุนเวียน"]],
    ["spacer", null, null, null],
    ["header", "ส่วนของเจ้าของ", null, null],
    ["item", "ทุนเรือนหุ้น", null, "ทุนเรือนหุ้น"],
    ["item", "กำไร(ขาดทุน)สะสม", null, "กำไร(ขาดทุน)สะสม"],
    ["subtotal", "รวมส่วนของเจ้าของ", null, ["ทุนเรือนหุ้น", "กำไร(ขาดทุน)สะสม"]],
    ["subtotal", "รวมหนี้สินและส่วนของเจ้าของ", null, ["รวมหนี้สิน", "รวมส่วนของเจ้าของ"]],
  ],
  pl_structure: [
    ["header", "รายได้", null, null],
    ["item", "รายได้จากการให้บริการ", null, "รายได้จากการให้บริการ"],
    ["item", "รายได้อื่น", null, "รายได้อื่น"],
    ["subtotal", "รวมรายได้", null, ["รายได้จากการให้บริการ", "รายได้อื่น"]],
    ["spacer", null, null, null],
    ["header", "ค่าใช้จ่าย", null, null],
    ["item", "ต้นทุนการให้บริการ", null, "ต้นทุนการให้บริการ"],
    ["item", "ค่าใช้จ่ายในการขายและบริหาร", null, "ค่าใช้จ่ายในการขายและบริหาร"],
    ["subtotal", "รวมค่าใช้จ่าย", null, ["ต้นทุนการให้บริการ", "ค่าใช้จ่ายในการขายและบริหาร"]],
    ["spacer", null, null, null],
    ["subtotal", "กำไร(ขาดทุน)ก่อนต้นทุนทางการเงินและภาษีเงินได้", null, ["รวมรายได้", "-รวมค่าใช้จ่าย"]],
    ["item", "ต้นทุนทางการเงิน", null, "ต้นทุนทางการเงิน"],
    ["subtotal", "กำไร(ขาดทุน)ก่อนภาษีเงินได้", null, ["กำไร(ขาดทุน)ก่อนต้นทุนทางการเงินและภาษีเงินได้", "-ต้นทุนทางการเงิน"]],
    ["item", "ภาษีเงินได้", null, "ภาษีเงินได้"],
    ["subtotal", "กำไร(ขาดทุน)สุทธิ", null, ["กำไร(ขาดทุน)ก่อนภาษีเงินได้", "-ภาษีเงินได้"]],
  ],
  eq_structure: {
    columns: ["ทุนที่ออกและเรียกชำระแล้ว", "กำไร(ขาดทุน)สะสม", "รวมส่วนของเจ้าของ"],
    rows: [
      { type: "opening", label_template: "ยอด ณ วันต้นปี {year}", values: { capital: "open_capital", retained: "open_retained", total: "auto_sum" } },
      { type: "movement", label: "ออกหุ้นและชำระเงิน", values: { capital: "issued_capital", retained: 0, total: "auto_sum" } },
      { type: "movement", label: "กำไร(ขาดทุน)สุทธิสำหรับปี", values: { capital: 0, retained: "net_profit", total: "auto_sum" } },
      { type: "closing", label_template: "ยอดคงเหลือ ณ วันที่ 31 ธันวาคม {year}", values: { capital: "close_capital", retained: "close_retained", total: "auto_sum" } },
    ],
    sheet_title: "งบการเปลี่ยนแปลงส่วนของเจ้าของ",
  },
};

// ======================================================
// 1. CONFIG MANAGER
// ======================================================
const CONFIG_STORAGE_KEY = "gl_tb_reconciler_config";

function getActiveConfig() {
  try {
    const saved = localStorage.getItem(CONFIG_STORAGE_KEY);
    if (saved) return JSON.parse(saved);
  } catch (e) { /* ignore */ }
  return JSON.parse(JSON.stringify(DEFAULT_CONFIG));
}

function saveConfig(config) {
  localStorage.setItem(CONFIG_STORAGE_KEY, JSON.stringify(config));
}

function resetConfig() {
  localStorage.removeItem(CONFIG_STORAGE_KEY);
}

function getFsLineItems() {
  const cfg = getActiveConfig();
  const items = cfg.fs_line_items || DEFAULT_CONFIG.fs_line_items;
  const cleaned = {};
  for (const [k, v] of Object.entries(items)) cleaned[k.trim()] = v;
  return cleaned;
}

function getBsStructure() {
  const cfg = getActiveConfig();
  return cfg.bs_structure || DEFAULT_CONFIG.bs_structure;
}

function getPlStructure() {
  const cfg = getActiveConfig();
  return cfg.pl_structure || DEFAULT_CONFIG.pl_structure;
}

function getEqStructure() {
  const cfg = getActiveConfig();
  return cfg.eq_structure || DEFAULT_CONFIG.eq_structure;
}

function validateConfig(config) {
  const issues = [];
  if (typeof config !== "object" || config === null) {
    return [{ severity: "error", message: "Config root must be a JSON object" }];
  }
  const fsItems = config.fs_line_items || {};
  if (typeof fsItems !== "object") {
    issues.push({ severity: "error", message: "fs_line_items must be a JSON object" });
    return issues;
  }
  const validSides = new Set(["bs_debit", "bs_credit", "pl_debit", "pl_credit"]);
  for (const [name, rules] of Object.entries(fsItems)) {
    if (typeof rules !== "object") {
      issues.push({ severity: "error", message: `'${name}': value must be an object with prefixes/keywords/side` });
      continue;
    }
    if (!rules.side) issues.push({ severity: "warning", message: `'${name}': missing 'side' field` });
    else if (!validSides.has(rules.side)) issues.push({ severity: "warning", message: `'${name}': side='${rules.side}' is not valid` });
    if (!rules.prefixes) issues.push({ severity: "warning", message: `'${name}': missing 'prefixes' field` });
    if (!rules.keywords) issues.push({ severity: "warning", message: `'${name}': missing 'keywords' field` });
  }
  // Validate BS/PL structure references
  for (const [structName, struct] of [["bs_structure", config.bs_structure], ["pl_structure", config.pl_structure]]) {
    if (!Array.isArray(struct)) continue;
    for (const row of struct) {
      if (!Array.isArray(row) || row.length < 4) { issues.push({ severity: "error", message: `${structName} row too short: ${JSON.stringify(row)}` }); continue; }
      const [rType, rLabel, rNote, rKey] = row;
      if (rType === "item" && rKey) {
        const refs = Array.isArray(rKey) ? rKey : [rKey];
        for (const ref of refs) {
          const cleanRef = ref.replace(/^-/, "");
          if (!fsItems[cleanRef]) issues.push({ severity: "warning", message: `${structName} item '${rLabel}' references '${cleanRef}' not in fs_line_items` });
        }
      }
    }
  }
  if (issues.length === 0) issues.push({ severity: "success", message: "All validations passed ✅" });
  return issues;
}

// ======================================================
// 2. THAI PDF PUA CHARACTER CLEANING
// ======================================================
const THAI_FONT_MAP = {
  "\uf700": "\u0e40", "\uf701": "\u0e41", "\uf702": "\u0e42", "\uf703": "\u0e43",
  "\uf704": "\u0e44", "\uf705": "\u0e48", "\uf706": "\u0e49", "\uf707": "\u0e4a",
  "\uf708": "\u0e4b", "\uf709": "\u0e4c", "\uf70a": "\u0e48", "\uf70b": "\u0e49",
  "\uf70c": "\u0e4a", "\uf70d": "\u0e4b", "\uf70e": "\u0e4c", "\uf70f": "\u0e4d",
  "\uf710": "\u0e31", "\uf711": "\u0e34", "\uf712": "\u0e35", "\uf713": "\u0e36",
  "\uf714": "\u0e37", "\uf715": "\u0e38", "\uf716": "\u0e39", "\uf717": "\u0e47",
  "\uf718": "\u0e48", "\uf719": "\u0e49", "\uf71a": "\u0e4a", "\uf71b": "\u0e4b",
  "\uf71c": "\u0e4c",
};

function cleanThaiPdfText(text) {
  for (const [char, replacement] of Object.entries(THAI_FONT_MAP)) {
    text = text.split(char).join(replacement);
  }
  text = text.replace(/[\uf700-\uf74f]/g, "");
  return text;
}

// ======================================================
// 3. PDF TEXT EXTRACTION
// ======================================================
async function extractTextFromPdf(file) {
  const arrayBuffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument(arrayBuffer).promise;
  const fullLines = [];

  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const textContent = await page.getTextContent();
    const items = textContent.items.map(item => ({
      str: item.str,
      x: item.transform[4],
      y: item.transform[5],
    }));
    items.sort((a, b) => {
      if (Math.abs(a.y - b.y) > 2) return b.y - a.y;
      return a.x - b.x;
    });

    let currentY = null;
    let line = [];
    for (const item of items) {
      if (currentY === null || Math.abs(currentY - item.y) > 2) {
        if (line.length > 0) fullLines.push(line.join(" "));
        line = [item.str.trim()];
        currentY = item.y;
      } else {
        if (item.str.trim() !== "") line.push(item.str.trim());
      }
    }
    if (line.length > 0) fullLines.push(line.join(" "));
  }
  return fullLines;
}

// ======================================================
// 4. NUMBER PARSING UTILITIES
// ======================================================
function parseNumberStr(str) {
  if (!str) return 0;
  str = str.replace(/,/g, "");
  const isNegative = str.includes("(") && str.includes(")");
  str = str.replace(/\(/g, "").replace(/\)/g, "");
  const val = parseFloat(str);
  if (isNaN(val)) return 0;
  return isNegative ? -val : val;
}

function formatMoney(val) {
  if (val === null || val === undefined || isNaN(val)) return "";
  return val.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function moneyFormatter(cell) {
  return formatMoney(cell.getValue());
}

// ======================================================
// 5. PARSERS
// ======================================================

// 5.1 Trial Balance Excel Parser (with extended PL/BS columns)
async function parseTbExcel(file) {
  const arrayBuffer = await file.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

  const tbData = [];
  let startReading = false;

  for (const row of rows) {
    const accId = String(row[0] || "").trim();
    if (!startReading) {
      if (/^\d{4}-\d{2}$/.test(accId)) startReading = true;
      else continue;
    }
    if (!accId || !/^\d{4}-\d{2}$/.test(accId)) continue;

    const accName = String(row[1] || "").trim();
    const debit = parseFloat(row[2]) || 0.0;
    const credit = parseFloat(row[3]) || 0.0;
    const plDebit = row.length > 4 ? parseFloat(row[4]) || 0.0 : 0.0;
    const plCredit = row.length > 5 ? parseFloat(row[5]) || 0.0 : 0.0;
    const bsDebit = row.length > 6 ? parseFloat(row[6]) || 0.0 : 0.0;
    const bsCredit = row.length > 7 ? parseFloat(row[7]) || 0.0 : 0.0;

    tbData.push({
      "Account ID": accId,
      "Account Name": accName,
      "TB Debit": debit,
      "TB Credit": credit,
      "TB Net Balance": Math.abs(debit - credit),
      "PL Debit": plDebit,
      "PL Credit": plCredit,
      "BS Debit": bsDebit,
      "BS Credit": bsCredit,
    });
  }
  return tbData;
}

// 5.2 Working Paper PDF Parser
async function parseTbWorkingPaperPdf(file) {
  const lines = await extractTextFromPdf(file);
  const tbData = [];

  for (const rawLine of lines) {
    const line = rawLine.trim();
    const match = line.match(/^\s*(\d{4}-\d{2})\s+(.*)/);
    if (!match) continue;

    const accId = match[1].trim();
    const rest = match[2];
    const allNums = rest.match(/[\d,]+\.\d{2}/g);
    if (!allNums || allNums.length === 0) continue;

    const nameMatch = rest.match(/^(.+?)(?=\s{3,}[\d,]+\.\d{2})/);
    let accName = nameMatch ? cleanThaiPdfText(nameMatch[1].trim()) : "Unknown";
    if (/^[\d,]+\.\d{2}$/.test(accName)) accName = "Unknown";

    const isBsAccount = ["1", "2", "3"].includes(accId[0]);
    const isDebitNormal = ["1", "5"].includes(accId[0]);

    let tbDebit = 0, tbCredit = 0, plDebit = 0, plCredit = 0, bsDebit = 0, bsCredit = 0;

    if (allNums.length >= 1) {
      const val = parseFloat(allNums[0].replace(/,/g, ""));
      if (isDebitNormal) tbDebit = val;
      else tbCredit = val;
    }

    if (allNums.length >= 2) {
      const lastVal = parseFloat(allNums[allNums.length - 1].replace(/,/g, ""));
      if (isBsAccount) {
        if (isDebitNormal) bsDebit = lastVal;
        else bsCredit = lastVal;
      } else {
        if (accId[0] === "4") plCredit = lastVal;
        else plDebit = lastVal;
      }
    } else {
      const val = parseFloat(allNums[0].replace(/,/g, ""));
      if (isBsAccount) {
        if (isDebitNormal) bsDebit = val;
        else bsCredit = val;
      } else {
        if (accId[0] === "4") plCredit = val;
        else plDebit = val;
      }
    }

    tbData.push({
      "Account ID": accId,
      "Account Name": accName,
      "TB Debit": tbDebit,
      "TB Credit": tbCredit,
      "TB Net Balance": Math.abs(tbDebit - tbCredit),
      "PL Debit": plDebit,
      "PL Credit": plCredit,
      "BS Debit": bsDebit,
      "BS Credit": bsCredit,
    });
  }
  return tbData;
}

// 5.3 General Ledger PDF Parser
async function parseGlPdf(file) {
  const lines = await extractTextFromPdf(file);
  const cleanedLines = lines.map(l => cleanThaiPdfText(l));
  const glData = [];
  let currentAccId = null;
  let extractedCompany = "";
  let extractedYear = "";

  // Extract company info from header
  for (let i = 0; i < Math.min(10, cleanedLines.length); i++) {
    const line = cleanedLines[i].trim();
    if (line.includes("บริษัท") && !extractedCompany) {
      const parts = line.split(/\s{3,}|\t|หน้า|หนา/);
      extractedCompany = parts[0].trim();
    }
    if (line.includes("วันที่จาก") && !extractedYear) {
      const yrMatch = line.match(/(25\d{2})/);
      if (yrMatch) extractedYear = "พ.ศ. " + yrMatch[1];
    }
  }

  for (const line of cleanedLines) {
    const trimmed = line.trim();
    const matchId = trimmed.match(/^(\d{4}-\d{2})\s+(.*)/);
    if (matchId) {
      currentAccId = matchId[1].trim();
      const restOfLine = matchId[2];
      const existing = glData.find(item => item["Account ID"] === currentAccId);

      if (!existing) {
        const accNameMatch = restOfLine.match(/^([^\(\s]+(?:\s[^\(\s]+)*)/);
        const glAccName = accNameMatch ? cleanThaiPdfText(accNameMatch[1].trim()) : "Unknown";

        let bfVal = 0.0;
        const bfMatch = restOfLine.match(/\s+((?:\()?\s*[\d,]+\.\d{2}(?:\))?)$/);
        if (bfMatch) {
          const bfStr = bfMatch[1];
          bfVal = parseFloat(bfStr.replace("(", "").replace(")", "").replace(/,/g, ""));
          if (bfStr.includes("(")) bfVal = -bfVal;
        }

        const initialNet = ["1", "2", "3"].includes(currentAccId[0]) ? Math.abs(bfVal) : 0.0;
        glData.push({
          "Account ID": currentAccId,
          "GL Account Name": glAccName,
          "GL Brought Forward": bfVal,
          "GL Debit": 0.0,
          "GL Credit": 0.0,
          "GL Net Balance": initialNet,
        });
      }
    }

    const totalMatch = trimmed.match(/รวม\s+([\d,.]+)\s+([\d,.]+)/);
    if (totalMatch && currentAccId !== null) {
      const debitVal = parseFloat(totalMatch[1].replace(/,/g, "")) || 0.0;
      const creditVal = parseFloat(totalMatch[2].replace(/,/g, "")) || 0.0;
      const existing = glData.find(item => item["Account ID"] === currentAccId);
      if (existing) {
        existing["GL Debit"] = debitVal;
        existing["GL Credit"] = creditVal;
        if (["1", "2", "3"].includes(currentAccId[0])) {
          existing["GL Net Balance"] = Math.abs(existing["GL Brought Forward"] + debitVal - creditVal);
        } else {
          existing["GL Net Balance"] = Math.abs(debitVal - creditVal);
        }
      }
    }
  }
  return { glData, extractedCompany, extractedYear };
}

// 5.4 Trial Balance PDF (BF Check)
async function parseTbPdf(file) {
  const lines = await extractTextFromPdf(file);
  const cleanedLines = lines.map(l => cleanThaiPdfText(l));
  const tbBfData = [];

  for (const line of cleanedLines) {
    const matchId = line.match(/^(\d{4}-\d{2})\s+(.*)/);
    if (matchId) {
      const accId = matchId[1];
      const restOfLine = matchId[2];
      const numbers = restOfLine.match(/((?:\()?[\d,]+\.\d{2}(?:\))?)/g);
      if (numbers && numbers.length >= 6) {
        const bfDr = parseNumberStr(numbers[numbers.length - 6]);
        const bfCr = parseNumberStr(numbers[numbers.length - 5]);
        tbBfData.push({
          "Account ID": accId,
          "TB Brought Forward Net": Math.abs(bfDr - bfCr),
        });
      }
    }
  }
  return tbBfData;
}

// 5.5 Tax Extraction from Working Paper Excel
async function extractTaxFromWorkingPaper(file) {
  const arrayBuffer = await file.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const result = {
    tax_amount: null,
    taxable_profit: null,
    accounting_profit: null,
    non_deductible: null,
    prepaid_tax: null,
    withholding_tax: null,
  };
  let foundSection = false;

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

    for (const row of rows) {
      const cellA = String(row[0] || "").replace(/\xa0/g, " ").trim();
      const cellB = row.length > 1 ? String(row[1] || "").replace(/\xa0/g, " ").trim() : "";
      const combined = (cellA + " " + cellB).trim();

      if (combined.includes("การคำนวณภาษีเงินได้นิติบุคคล")) {
        foundSection = true;
        continue;
      }
      if (!foundSection) continue;

      const colC = row.length > 2 && typeof row[2] === "number" ? row[2] : null;
      const colD = row.length > 3 && typeof row[3] === "number" ? row[3] : null;

      if (combined.includes("กำไรสุทธิทางบัญชี") && colD !== null) result.accounting_profit = colD;
      if (combined.includes("ค่าใช้จ่ายต้องห้าม")) {
        if (colD !== null) result.non_deductible = colD;
        else if (colC !== null) result.non_deductible = colC;
      }
      if ((combined.includes("กำไร(ขาดทุน)สุทธิทางภาษี") || combined.includes("กำไรสุทธิทางภาษี")) && colD !== null) result.taxable_profit = colD;
      if (combined.includes("อัตราภาษี") && !combined.includes("300,000") && !combined.includes("ได้รับยกเว้น")) {
        if (colD !== null) result.tax_amount = colD;
        else if (colC !== null) result.tax_amount = colC;
      }
      if (combined.includes("ภาษีเงินได้นิติบุคคลจ่ายล่วงหน้า") && colD !== null) result.prepaid_tax = colD;
      if (combined.includes("ภาษีถูกหัก") && combined.includes("ม.3") && !combined.includes("ค้างจ่าย")) {
        if (colC !== null && result.withholding_tax === null) result.withholding_tax = colC;
      }
    }
  }
  return foundSection ? result : null;
}

// ======================================================
// 6. AUTO-MAPPING ENGINE
// ======================================================
function autoMap(row) {
  const activeItems = getFsLineItems();
  const accId = String(row["Account ID"] || "").trim();
  const accName = String(row["Account Name"] || "").trim();

  // Priority 1: Prefix matching (longest match wins)
  const matches = [];
  for (const [fsLine, rules] of Object.entries(activeItems)) {
    for (const prefix of rules.prefixes || []) {
      if (accId.startsWith(prefix)) {
        matches.push({ len: prefix.length, line: fsLine });
      }
    }
  }
  if (matches.length > 0) {
    matches.sort((a, b) => b.len - a.len);
    return matches[0].line;
  }

  // Priority 2: Keyword matching
  for (const [fsLine, rules] of Object.entries(activeItems)) {
    for (const keyword of rules.keywords || []) {
      if (accName.includes(keyword)) return fsLine;
    }
  }

  return "ไม่จัดประเภท (Unmapped)";
}

// ======================================================
// 7. FS VALUE CALCULATION
// ======================================================
function getFsValue(row) {
  const activeItems = getFsLineItems();
  const fsLine = row["FS Line Item"] || "";
  if (fsLine === "ไม่จัดประเภท (Unmapped)") return row["TB Net Balance"] || 0;

  const rules = activeItems[fsLine] || {};
  const side = rules.side || "bs_debit";
  const bsDr = row["BS Debit"] || 0;
  const bsCr = row["BS Credit"] || 0;
  const plDr = row["PL Debit"] || 0;
  const plCr = row["PL Credit"] || 0;

  // If no BS/PL breakdown, fallback to TB Net Balance
  if (bsDr === 0 && bsCr === 0 && plDr === 0 && plCr === 0) return row["TB Net Balance"] || 0;

  switch (side) {
    case "bs_debit": return bsDr - bsCr;
    case "bs_credit": return bsCr - bsDr;
    case "pl_debit": return plDr - plCr;
    case "pl_credit": return plCr - plDr;
    default: return row["TB Net Balance"] || 0;
  }
}

// ======================================================
// 8. FS STRUCTURE COMPUTATION ENGINE
// ======================================================
function buildFsFromMapping(fsMapping, structure) {
  const computed = {};
  for (const row of structure) {
    const [rowType, label, note, key] = row;
    if (rowType === "item") {
      if (Array.isArray(key)) {
        let total = 0;
        for (const k of key) {
          if (k.startsWith("-")) {
            total -= computed[k.substring(1)] ?? fsMapping[k.substring(1)] ?? 0;
          } else {
            total += computed[k] ?? fsMapping[k] ?? 0;
          }
        }
        computed[label] = total;
      } else if (key) {
        computed[label] = fsMapping[key] ?? 0;
      } else {
        computed[label] = 0;
      }
    } else if (rowType === "subtotal" && Array.isArray(key)) {
      let total = 0;
      for (const k of key) {
        if (k.startsWith("-")) {
          total -= computed[k.substring(1)] ?? 0;
        } else {
          total += computed[k] ?? 0;
        }
      }
      computed[label] = total;
    } else {
      computed[label] = null;
    }
  }
  return computed;
}

// ======================================================
// 9. RECONCILIATION ENGINE
// ======================================================
function processReconciliation(tbData, glData, tbBfData) {
  const tbMap = Object.fromEntries(tbData.map(x => [x["Account ID"], x]));
  const glMap = Object.fromEntries(glData.map(x => [x["Account ID"], x]));
  const bfMap = Object.fromEntries(tbBfData.map(x => [x["Account ID"], x]));

  const allIds = new Set([...Object.keys(tbMap), ...Object.keys(glMap)]);
  const merged = [];

  for (const id of allIds) {
    const tbItem = tbMap[id] || {};
    const glItem = glMap[id] || {};
    const bfItem = bfMap[id] || {};

    const row = {
      "Account ID": id,
      "Account Name": tbItem["Account Name"] || glItem["GL Account Name"] || "Unknown",
      "TB Net Balance": tbItem["TB Net Balance"] ?? 0,
      "GL Net Balance": glItem["GL Net Balance"] ?? 0,
      "TB Debit": tbItem["TB Debit"] || 0,
      "TB Credit": tbItem["TB Credit"] || 0,
      "GL Debit": glItem["GL Debit"] || 0,
      "GL Credit": glItem["GL Credit"] || 0,
      "GL Brought Forward": glItem["GL Brought Forward"] || 0,
      "TB Brought Forward Net": bfItem["TB Brought Forward Net"] || 0,
      "PL Debit": tbItem["PL Debit"] || 0,
      "PL Credit": tbItem["PL Credit"] || 0,
      "BS Debit": tbItem["BS Debit"] || 0,
      "BS Credit": tbItem["BS Credit"] || 0,
      "Missing in TB original": tbItem["TB Net Balance"] === undefined,
      "Missing in GL original": glItem["GL Net Balance"] === undefined,
    };

    row["Difference"] = Math.abs(row["GL Net Balance"] - row["TB Net Balance"]);
    row["Is Match"] = row["Difference"] <= 0.02;
    row["BF Difference"] = Math.abs(Math.abs(row["GL Brought Forward"]) - row["TB Brought Forward Net"]);

    // Status determination
    if (row["Missing in TB original"]) {
      row["Status"] = row["GL Net Balance"] === 0 ? "Match (Zero Balance/Closed)" : "Missing in TB";
    } else if (row["Missing in GL original"]) {
      row["Status"] = row["TB Net Balance"] === 0 ? "Match (Zero Balance/Closed)" : "Missing in GL";
    } else if (!row["Is Match"]) {
      row["Status"] = "Amount Mismatch";
    } else {
      row["Status"] = "Match";
    }

    row["Max Absolute Balance"] = Math.max(row["TB Net Balance"], row["GL Net Balance"]);
    row["FS Line Item"] = autoMap(row);
    row["FS Value"] = getFsValue(row);

    merged.push(row);
  }

  return merged;
}

// ======================================================
// 10. CORPORATE TAX CALCULATION
// ======================================================
function calculateCorporateTax(mergedData, wpTaxInfo, taxMethod) {
  let corporateTax = 0;
  let taxSource = "";
  const activeItems = getFsLineItems();

  // Step 1: From working paper
  if (wpTaxInfo && wpTaxInfo.tax_amount !== null && wpTaxInfo.tax_amount > 0) {
    corporateTax = Math.round(wpTaxInfo.tax_amount * 100) / 100;
    taxSource = "working_paper";
    return { corporateTax, taxSource, wpTaxInfo };
  }

  // Step 2: Check if tax already in TB
  const hasTaxInTb = mergedData.some(row =>
    (row["Account Name"] || "").includes("ภาษีเงินได้นิติบุคคล") ||
    (row["Account Name"] || "").includes("ค่าใช้จ่ายภาษีเงินได้")
  );
  if (hasTaxInTb) {
    taxSource = "from_tb";
    return { corporateTax: 0, taxSource, wpTaxInfo: null };
  }

  // Step 3: Auto-calculate
  if (taxMethod === "manual") {
    taxSource = "manual_skip";
    return { corporateTax: 0, taxSource, wpTaxInfo: null };
  }

  let revTotal = 0, expTotal = 0;
  for (const row of mergedData) {
    const accId = row["Account ID"] || "";
    if (accId.startsWith("4")) {
      revTotal += (row["PL Credit"] || 0) - (row["PL Debit"] || 0);
    }
    if (accId.startsWith("5")) {
      expTotal += (row["PL Debit"] || 0) - (row["PL Credit"] || 0);
    }
  }
  if (revTotal === 0 && expTotal === 0) {
    for (const row of mergedData) {
      const accId = row["Account ID"] || "";
      if (accId.startsWith("4")) revTotal += row["TB Net Balance"] || 0;
      if (accId.startsWith("5")) expTotal += row["TB Net Balance"] || 0;
    }
  }

  const preTaxProfit = revTotal - expTotal;
  let appliedRule = taxMethod;

  if (taxMethod === "auto") {
    let totalRevenue = 0, totalCapital = 0;
    for (const row of mergedData) {
      if ((row["Account ID"] || "").startsWith("4")) totalRevenue += row["TB Net Balance"] || 0;
      if ((row["Account ID"] || "").startsWith("3") && (row["Account Name"] || "").includes("ทุน")) {
        totalCapital += row["TB Net Balance"] || 0;
      }
    }
    appliedRule = (totalCapital <= 5000000 && totalRevenue <= 30000000) ? "sme" : "standard";
  }

  if (preTaxProfit > 0) {
    taxSource = "auto_calc";
    let nonDeductible = 0;
    for (const row of mergedData) {
      if ((row["Account Name"] || "").includes("ค่าใช้จ่ายต้องห้าม") || (row["Account Name"] || "").includes("เบี้ยปรับเงินเพิ่ม")) {
        nonDeductible += row["TB Net Balance"] || 0;
      }
    }
    const taxableProfit = preTaxProfit + nonDeductible;

    if (appliedRule === "standard") {
      corporateTax = Math.round(taxableProfit * 0.20 * 100) / 100;
    } else if (appliedRule === "sme") {
      if (taxableProfit <= 300000) corporateTax = 0;
      else if (taxableProfit <= 3000000) corporateTax = Math.round((taxableProfit - 300000) * 0.15 * 100) / 100;
      else corporateTax = Math.round((405000 + (taxableProfit - 3000000) * 0.20) * 100) / 100;
    }

    return { corporateTax, taxSource, wpTaxInfo: null, preTaxProfit, taxableProfit, nonDeductible, appliedRule };
  }

  return { corporateTax: 0, taxSource: "", wpTaxInfo: null };
}

// ======================================================
// 10.1 PRIOR YEAR FS PARSER (Excel)
// ======================================================
async function parsePriorFsExcel(file) {
  const arrayBuffer = await file.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const activeItems = getFsLineItems();
  const allConfigKeys = new Set(Object.keys(activeItems));

  const bsStructure = getBsStructure();
  const plStructure = getPlStructure();
  const structureItemLabels = new Set();
  [...bsStructure, ...plStructure].forEach(row => {
    if (row[0] === "item" && row[1]) structureItemLabels.add(String(row[1]).trim());
  });

  const allKnownLabels = new Set([...allConfigKeys, ...structureItemLabels]);
  const extracted = {};

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

    for (const row of rows) {
      if (!Array.isArray(row) || row.length === 0) continue;

      let cellLabel = "";
      let labelCol = -1;
      for (let c = 0; c < Math.min(4, row.length); c++) {
        const val = String(row[c] || "").replace(/\xa0/g, " ").trim();
        if (val && isNaN(val) && !/^\d+$/.test(val)) {
          cellLabel = val;
          labelCol = c;
          break;
        }
      }
      if (!cellLabel) continue;

      let matchedKey = null;
      if (allKnownLabels.has(cellLabel)) {
        matchedKey = cellLabel;
      } else {
        const skipKeywords = ["รวม", "กำไร(ขาดทุน)ก่อน", "กำไร(ขาดทุน)สุทธิ"];
        if (!skipKeywords.some(kw => cellLabel.includes(kw))) {
          let bestMatch = null;
          let bestScore = 0;
          for (const configKey of allConfigKeys) {
            let score = 0;
            if (cellLabel.includes(configKey) || configKey.includes(cellLabel)) {
              const shorter = Math.min(cellLabel.length, configKey.length);
              const longer = Math.max(cellLabel.length, configKey.length);
              if (shorter / longer >= 0.5) score = shorter;
            }
            if (score > bestScore) {
              bestScore = score;
              bestMatch = configKey;
            }
          }
          if (bestMatch) matchedKey = bestMatch;
        }
      }

      if (!matchedKey) continue;

      let bestVal = null;
      for (let c = labelCol + 1; c < row.length; c++) {
        const val = row[c];
        if (typeof val === "number" && !isNaN(val)) {
          bestVal = val;
        } else if (typeof val === "string" && val.trim() !== "") {
          const num = parseFloat(val.replace(/,/g, ""));
          if (!isNaN(num)) bestVal = num;
        }
      }

      if (bestVal !== null && !(matchedKey in extracted)) {
        extracted[matchedKey] = bestVal;
      }
    }
  }

  return extracted;
}

// ======================================================
// 11. GLOBAL STATE
// ======================================================
let mergedData = [];
let tables = {};
let wpTaxInfo = null;
let priorYearSummary = null;
let priorYearLabel = "";

// ======================================================
// 12. FIND DYNAMIC CONFIG LABELS
// ======================================================
function findShareCapitalLabel() {
  const items = getFsLineItems();
  for (const [name, rules] of Object.entries(items)) {
    if (rules.side === "bs_credit" && ["ทุน", "หุ้น", "capital"].some(kw => name.includes(kw))) {
      if (!name.includes("กู้") && !name.includes("เจ้าหนี้")) return name;
    }
  }
  return "ทุนเรือนหุ้น";
}

function findRetainedEarningsLabel() {
  const items = getFsLineItems();
  for (const [name] of Object.entries(items)) {
    if (name.includes("กำไร") && name.includes("สะสม")) return name;
  }
  return "กำไร(ขาดทุน)สะสม";
}

// ======================================================
// 13. UI — INITIALIZATION
// ======================================================
document.addEventListener("DOMContentLoaded", () => {
  // Theme toggle
  const themeToggle = document.getElementById("themeToggle");
  if (themeToggle) {
    const savedTheme = localStorage.getItem("theme") || "dark";
    document.documentElement.setAttribute("data-theme", savedTheme);
    themeToggle.textContent = savedTheme === "dark" ? "🌙" : "☀️";
    themeToggle.addEventListener("click", () => {
      const current = document.documentElement.getAttribute("data-theme");
      const next = current === "dark" ? "light" : "dark";
      document.documentElement.setAttribute("data-theme", next);
      localStorage.setItem("theme", next);
      themeToggle.textContent = next === "dark" ? "🌙" : "☀️";
    });
  }

  // Run button
  document.getElementById("runBtn")?.addEventListener("click", runReconciliation);

  // Tab switching
  document.querySelectorAll(".tab-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));
      document.querySelectorAll(".tab-content > .tab-pane").forEach(p => p.classList.remove("active"));
      btn.classList.add("active");
      const target = document.getElementById(btn.dataset.target);
      if (target) target.classList.add("active");
      // Redraw tables
      Object.values(tables).forEach(t => { try { t.redraw(); } catch(e) {} });
    });
  });

  // Sub-tab switching
  document.querySelectorAll(".sub-tab-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".sub-tab-btn").forEach(b => b.classList.remove("active"));
      document.querySelectorAll(".sub-tab-content > .tab-pane").forEach(p => p.classList.remove("active"));
      btn.classList.add("active");
      const target = document.getElementById(btn.dataset.target);
      if (target) target.classList.add("active");
    });
  });

  // Export buttons
  document.getElementById("exportReconBtn")?.addEventListener("click", exportReconciliation);
  document.getElementById("exportFsBtn")?.addEventListener("click", exportFinancialStatements);

  // Config buttons
  document.getElementById("cfgDeployBtn")?.addEventListener("click", deployConfig);
  document.getElementById("cfgExportBtn")?.addEventListener("click", exportConfig);
  document.getElementById("cfgResetBtn")?.addEventListener("click", resetDefaultConfig);
  document.getElementById("cfgImportFile")?.addEventListener("change", importConfigFile);

  // Capital info live update
  ["regShares", "regPar", "paidShares", "paidPar"].forEach(id => {
    document.getElementById(id)?.addEventListener("input", updateCapitalInfo);
  });
  updateCapitalInfo();

  // File upload dropzone initialization
  setupDropZones();

  // Initialize config editors
  initConfigEditors();
});

// ======================================================
// 13.1 DROP ZONE & FILE STATUS HANDLING
// ======================================================
function setupDropZones() {
  const fileInputs = ["tbFile", "glFile", "tbPdfFile", "priorFile"];

  fileInputs.forEach(id => {
    const input = document.getElementById(id);
    const zone = document.getElementById(`drop-${id}`);
    if (!input || !zone) return;

    // Click on dropzone triggers file input click
    zone.addEventListener("click", (e) => {
      if (e.target.closest(".btn-clear-file")) return;
      input.click();
    });

    // Handle input change
    input.addEventListener("change", () => {
      updateDropZoneStatus(input, zone);
    });

    // Drag and Drop events
    ["dragenter", "dragover"].forEach(eventName => {
      zone.addEventListener(eventName, (e) => {
        e.preventDefault();
        e.stopPropagation();
        zone.classList.add("dragover");
      }, false);
    });

    ["dragleave", "drop"].forEach(eventName => {
      zone.addEventListener(eventName, (e) => {
        e.preventDefault();
        e.stopPropagation();
        zone.classList.remove("dragover");
      }, false);
    });

    zone.addEventListener("drop", (e) => {
      const dt = e.dataTransfer;
      const files = dt.files;
      if (files && files.length > 0) {
        input.files = files;
        updateDropZoneStatus(input, zone);
      }
    });
  });

  // Clear file buttons
  document.querySelectorAll(".btn-clear-file").forEach(btn => {
    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      const targetId = btn.dataset.target;
      const input = document.getElementById(targetId);
      const zone = document.getElementById(`drop-${targetId}`);
      if (input) input.value = "";
      if (zone) {
        zone.classList.remove("has-file");
        const nameEl = zone.querySelector(".file-name-text");
        const sizeEl = zone.querySelector(".file-size-text");
        if (nameEl) nameEl.textContent = "";
        if (sizeEl) sizeEl.textContent = "";
      }
    });
  });
}

function updateDropZoneStatus(input, zone) {
  if (input.files && input.files.length > 0) {
    const file = input.files[0];
    zone.classList.add("has-file");
    const nameEl = zone.querySelector(".file-name-text");
    const sizeEl = zone.querySelector(".file-size-text");
    if (nameEl) nameEl.textContent = `✅ ${file.name}`;
    if (sizeEl) sizeEl.textContent = `ขนาดไฟล์: ${formatFileSize(file.size)}`;
  } else {
    zone.classList.remove("has-file");
  }
}

function formatFileSize(bytes) {
  if (bytes === 0) return "0 Bytes";
  const k = 1024;
  const sizes = ["Bytes", "KB", "MB", "GB"];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + " " + sizes[i];
}

function updateCapitalInfo() {
  const regShares = parseInt(document.getElementById("regShares")?.value) || 0;
  const regPar = parseFloat(document.getElementById("regPar")?.value) || 0;
  const paidShares = parseInt(document.getElementById("paidShares")?.value) || 0;
  const paidPar = parseFloat(document.getElementById("paidPar")?.value) || 0;

  const regInfo = document.getElementById("regCapitalInfo");
  if (regInfo) regInfo.innerHTML = `💡 ทุนจดทะเบียนรวม: <strong>${formatMoney(regShares * regPar)}</strong> บาท`;

  const paidInfo = document.getElementById("paidCapitalInfo");
  if (paidInfo) paidInfo.innerHTML = `💡 ทุนที่เรียกชำระแล้วรวม: <strong>${formatMoney(paidShares * paidPar)}</strong> บาท`;
}

// ======================================================
// 14. TOAST NOTIFICATION
// ======================================================
function showToast(message, type = "success", duration = 4000) {
  const toast = document.getElementById("toast");
  if (!toast) return;
  toast.textContent = message;
  toast.className = `toast toast--${type} show`;
  setTimeout(() => { toast.classList.remove("show"); }, duration);
}

// ======================================================
// 15. MAIN RECONCILIATION FLOW
// ======================================================
async function runReconciliation() {
  const tbFile = document.getElementById("tbFile")?.files[0];
  const glFile = document.getElementById("glFile")?.files[0];
  const tbPdfFile = document.getElementById("tbPdfFile")?.files[0];

  if (!tbFile && !glFile) {
    showToast("กรุณาอัปโหลดไฟล์ในช่องที่ 1 (งบทดลอง/กระดาษทำการ) และช่องที่ 2 (บัญชีแยกประเภท GL)", "error", 6000);
    return;
  }
  if (!tbFile) {
    showToast("กรุณาอัปโหลดไฟล์ในช่องที่ 1: งบทดลองสิ้นปี (TB) หรือกระดาษทำการ", "error", 6000);
    return;
  }
  if (!glFile) {
    showToast("กรุณาอัปโหลดไฟล์ในช่องที่ 2: บัญชีแยกประเภท (GL) PDF", "error", 6000);
    return;
  }

  const btn = document.getElementById("runBtn");
  const loading = document.getElementById("loading");
  const results = document.getElementById("resultsArea");

  btn.disabled = true;
  loading.classList.remove("hidden");
  results.classList.add("hidden");

  try {
    // Parse TB
    let tbData;
    if (tbFile.name.toLowerCase().endsWith(".pdf")) {
      tbData = await parseTbWorkingPaperPdf(tbFile);
    } else {
      tbData = await parseTbExcel(tbFile);
    }

    // Try extract tax from working paper (Excel only)
    wpTaxInfo = null;
    if (!tbFile.name.toLowerCase().endsWith(".pdf")) {
      try {
        wpTaxInfo = await extractTaxFromWorkingPaper(tbFile);
      } catch (e) { wpTaxInfo = null; }
    }

    // Parse GL
    const glResult = await parseGlPdf(glFile);
    const glData = glResult.glData;

    // Auto-fill company info from GL
    if (glResult.extractedCompany) {
      const companyInput = document.getElementById("companyName");
      if (companyInput && !companyInput.value) companyInput.value = glResult.extractedCompany;
    }
    if (glResult.extractedYear) {
      const yearInput = document.getElementById("currentYear");
      if (yearInput && (yearInput.value === "พ.ศ. 2568" || !yearInput.value)) yearInput.value = glResult.extractedYear;
      // Auto-fill prior year
      const yrMatch = glResult.extractedYear.match(/\d{4}/);
      if (yrMatch) {
        const priorInput = document.getElementById("priorYear");
        if (priorInput && (priorInput.value === "พ.ศ. 2567" || !priorInput.value)) {
          priorInput.value = `พ.ศ. ${parseInt(yrMatch[0]) - 1}`;
        }
      }
    }

    // Parse TB BF PDF
    let tbBfData = [];
    if (tbPdfFile) {
      tbBfData = await parseTbPdf(tbPdfFile);
    }

    // Parse Prior Year File if uploaded
    priorYearSummary = null;
    priorYearLabel = document.getElementById("priorYear")?.value || "พ.ศ. 2567";
    const priorFile = document.getElementById("priorFile")?.files[0];
    const priorUploadMode = document.getElementById("priorUploadMode")?.value || "working_paper";

    if (priorFile) {
      try {
        if (priorUploadMode === "fs") {
          priorYearSummary = await parsePriorFsExcel(priorFile);
          const matchCount = Object.keys(priorYearSummary).length;
          showToast(`✅ โหลดข้อมูลปีก่อนหน้า (${priorYearLabel}) จากงบการเงินเรียบร้อย (${matchCount} รายการ)`, "success");
        } else {
          let priorTbData;
          if (priorFile.name.toLowerCase().endsWith(".pdf")) {
            priorTbData = await parseTbWorkingPaperPdf(priorFile);
          } else {
            priorTbData = await parseTbExcel(priorFile);
          }
          if (priorTbData && priorTbData.length > 0) {
            priorYearSummary = {};
            for (const r of priorTbData) {
              r["FS Line Item"] = autoMap(r);
              r["FS Value"] = getFsValue(r);
              const line = r["FS Line Item"];
              if (line && line !== "ไม่จัดประเภท (Unmapped)") {
                priorYearSummary[line] = (priorYearSummary[line] || 0) + (r["FS Value"] || 0);
              }
            }
            showToast(`✅ โหลดข้อมูลปีก่อนหน้า (${priorYearLabel}) จากกระดาษทำการเรียบร้อย (${priorTbData.length} บัญชี)`, "success");
          }
        }
      } catch (err) {
        console.error("Could not parse prior year file:", err);
        showToast("⚠️ ไม่สามารถอ่านไฟล์ปีก่อนหน้าได้: " + err.message, "warning");
      }
    }

    // Run reconciliation
    mergedData = processReconciliation(tbData, glData, tbBfData);

    // Corporate tax
    const taxMethod = document.getElementById("taxMethod")?.value || "auto";
    const taxResult = calculateCorporateTax(mergedData, wpTaxInfo, taxMethod);

    if (taxResult.corporateTax > 0 && (taxResult.taxSource === "working_paper" || taxResult.taxSource === "auto_calc")) {
      // Inject tax rows
      const activeItems = getFsLineItems();
      let taxFsName = "ภาษีเงินได้";
      let otherCaFsName = "สินทรัพย์หมุนเวียนอื่น";
      for (const [name, rules] of Object.entries(activeItems)) {
        if (name.includes("ภาษีเงินได้") && rules.side === "pl_debit") taxFsName = name;
        if (rules.side === "bs_debit" && ["หมุนเวียนอื่น", "สินทรัพย์หมุนเวียนอื่น"].some(kw => name.includes(kw))) otherCaFsName = name;
      }
      const srcLabel = taxResult.taxSource === "working_paper" ? "WP" : "Auto";
      if (!mergedData.some(r => r["Account ID"] === "TAX-PL")) {
        mergedData.push({
          "Account ID": "TAX-PL", "Account Name": `ค่าใช้จ่ายภาษีเงินได้ (${srcLabel})`,
          "TB Net Balance": taxResult.corporateTax, "BS Debit": 0, "BS Credit": 0,
          "PL Debit": taxResult.corporateTax, "PL Credit": 0,
          "GL Net Balance": 0, "GL Brought Forward": 0, "GL Debit": 0, "GL Credit": 0,
          "TB Brought Forward Net": 0, "Difference": 0, "BF Difference": 0,
          "FS Line Item": taxFsName, "FS Value": taxResult.corporateTax,
          "Status": "Match", "Missing in TB original": false, "Missing in GL original": false,
          "Max Absolute Balance": taxResult.corporateTax,
        });
        mergedData.push({
          "Account ID": "TAX-BS", "Account Name": `ภาษีเงินได้ค้างจ่าย (${srcLabel})`,
          "TB Net Balance": taxResult.corporateTax, "BS Debit": 0, "BS Credit": taxResult.corporateTax,
          "PL Debit": 0, "PL Credit": 0,
          "GL Net Balance": 0, "GL Brought Forward": 0, "GL Debit": 0, "GL Credit": 0,
          "TB Brought Forward Net": 0, "Difference": 0, "BF Difference": 0,
          "FS Line Item": otherCaFsName, "FS Value": taxResult.corporateTax,
          "Status": "Match", "Missing in TB original": false, "Missing in GL original": false,
          "Max Absolute Balance": taxResult.corporateTax,
        });
      }
    }

    // Display tax info
    displayTaxInfo(taxResult);

    // Update UI
    updateUI();
    results.classList.remove("hidden");
    showToast(`กระทบยอดสำเร็จ! พบ ${mergedData.length} บัญชี`, "success");
  } catch (e) {
    showToast("เกิดข้อผิดพลาด: " + e.message, "error", 8000);
    console.error(e);
  } finally {
    btn.disabled = false;
    loading.classList.add("hidden");
  }
}

function displayTaxInfo(taxResult) {
  const area = document.getElementById("taxInfoArea");
  if (!area) return;
  area.innerHTML = "";

  if (taxResult.taxSource === "working_paper") {
    const info = taxResult.wpTaxInfo;
    let html = `<div class="success-box"><strong>💡 ภาษีจากกระดาษทำการ: ${formatMoney(taxResult.corporateTax)} บาท</strong>`;
    if (info.accounting_profit !== null) html += `<br>• กำไรสุทธิทางบัญชี: ${formatMoney(info.accounting_profit)}`;
    if (info.non_deductible !== null && info.non_deductible !== 0) html += `<br>• ค่าใช้จ่ายต้องห้าม: +${formatMoney(info.non_deductible)}`;
    if (info.taxable_profit !== null) html += `<br>• กำไรสุทธิทางภาษี: ${formatMoney(info.taxable_profit)}`;
    html += `</div>`;
    area.innerHTML = html;
  } else if (taxResult.taxSource === "from_tb") {
    area.innerHTML = `<div class="info-box">💡 พบรายการ 'ภาษีเงินได้นิติบุคคล' ในงบทดลอง — ใช้ยอดจาก TB โดยตรง</div>`;
  } else if (taxResult.taxSource === "auto_calc") {
    let html = `<div class="warning-box">⚠️ ไม่พบข้อมูลภาษีในกระดาษทำการ — คำนวณอัตโนมัติ`;
    html += `<br>• กำไรก่อนภาษี: ${formatMoney(taxResult.preTaxProfit)}`;
    if (taxResult.nonDeductible > 0) html += `<br>• ค่าใช้จ่ายต้องห้าม: +${formatMoney(taxResult.nonDeductible)}`;
    html += `<br>• อัตราภาษี: ${taxResult.appliedRule === "sme" ? "SME" : "Standard (20%)"}`;
    html += `<br>• ภาษีคำนวณได้: ${formatMoney(taxResult.corporateTax)}`;
    html += `</div>`;
    area.innerHTML = html;
  }
}

// ======================================================
// 16. UI UPDATES
// ======================================================
function updateUI() {
  const mismatches = mergedData.filter(d => d.Status === "Amount Mismatch");
  const missingGl = mergedData.filter(d => d.Status === "Missing in GL");
  const missingTb = mergedData.filter(d => d.Status === "Missing in TB");
  const matches = mergedData.filter(d => d.Status.includes("Match"));
  const suspects = [...mergedData].sort((a, b) => b["Max Absolute Balance"] - a["Max Absolute Balance"]).slice(0, 10);

  // Summary cards
  setText("sumTotal", mergedData.length);
  setText("sumMatched", matches.length);
  setText("sumMismatches", mismatches.length);
  setText("sumMissing", missingGl.length + missingTb.length);
  setText("countMismatches", mismatches.length);
  setText("countMissing", missingGl.length + missingTb.length);
  setText("countMatches", matches.length);

  // Common columns
  const colId = { title: "รหัสบัญชี", field: "Account ID", width: 110 };
  const colName = { title: "ชื่อบัญชี", field: "Account Name", widthGrow: 2 };
  const colGlNet = { title: "ยอด GL สุทธิ", field: "GL Net Balance", formatter: moneyFormatter, hozAlign: "right", width: 140 };
  const colTbNet = { title: "ยอด TB สุทธิ", field: "TB Net Balance", formatter: moneyFormatter, hozAlign: "right", width: 140 };
  const colDiff = { title: "ผลต่าง", field: "Difference", formatter: moneyFormatter, hozAlign: "right", width: 130 };
  const colStatus = { title: "สถานะ", field: "Status", width: 180 };
  const colFsValue = { title: "FS Value", field: "FS Value", formatter: moneyFormatter, hozAlign: "right", width: 140 };

  // Create tables
  tables.discrepancies = new Tabulator("#grid-discrepancies", {
    data: mismatches, layout: "fitColumns", height: 400,
    columns: [colId, colName, colGlNet, colTbNet, colDiff,
      { title: "ยอดยกมา GL", field: "GL Brought Forward", formatter: moneyFormatter, hozAlign: "right", width: 130 },
      { title: "ยอดยกมา TB", field: "TB Brought Forward Net", formatter: moneyFormatter, hozAlign: "right", width: 130 },
      { title: "BF Diff", field: "BF Difference", formatter: moneyFormatter, hozAlign: "right", width: 120 },
      colStatus],
  });

  tables.missingGl = new Tabulator("#grid-missing-gl", {
    data: missingGl, layout: "fitColumns", height: 300,
    columns: [colId, colName, colTbNet, colStatus],
  });

  tables.missingTb = new Tabulator("#grid-missing-tb", {
    data: missingTb, layout: "fitColumns", height: 300,
    columns: [colId, colName, colGlNet, colStatus],
  });

  tables.suspects = new Tabulator("#grid-suspects", {
    data: suspects, layout: "fitColumns", height: 400,
    columns: [colId, colName,
      { title: "ยอดคงเหลือสูงสุด", field: "Max Absolute Balance", formatter: moneyFormatter, hozAlign: "right" },
      colStatus],
  });

  tables.matches = new Tabulator("#grid-matches", {
    data: matches, layout: "fitColumns", height: 400,
    columns: [colId, colName, colGlNet, colTbNet,
      { title: "ยอดยกมา GL", field: "GL Brought Forward", formatter: moneyFormatter, hozAlign: "right" },
      { title: "ยอดยกมา TB", field: "TB Brought Forward Net", formatter: moneyFormatter, hozAlign: "right" },
      colStatus],
  });

  // Mapping editor
  const lineItemOptions = Object.fromEntries(
    [...Object.keys(getFsLineItems()), "ไม่จัดประเภท (Unmapped)"].map(k => [k, k])
  );

  const mappingData = mergedData.map(row => ({
    "Account ID": row["Account ID"],
    "Account Name": row["Account Name"],
    "TB Net Balance": row["TB Net Balance"],
    "FS Line Item": row["FS Line Item"],
    "FS Value": row["FS Value"],
  }));
  mappingData.sort((a, b) => (a["FS Line Item"] || "").localeCompare(b["FS Line Item"] || ""));

  tables.mapping = new Tabulator("#grid-mapping", {
    data: mappingData,
    layout: "fitColumns",
    height: 500,
    columns: [
      colId,
      colName,
      { title: "TB Net Balance", field: "TB Net Balance", formatter: moneyFormatter, hozAlign: "right", width: 140 },
      {
        title: "📌 FS Line Item", field: "FS Line Item", editor: "list",
        editorParams: { values: lineItemOptions },
        cellEdited: function (cell) {
          const accId = cell.getRow().getData()["Account ID"];
          const newFsLine = cell.getValue();
          // Update mergedData
          const row = mergedData.find(r => r["Account ID"] === accId);
          if (row) {
            row["FS Line Item"] = newFsLine;
            row["FS Value"] = getFsValue(row);
          }
          updateFsSummary();
          updateFsPreview();
        },
      },
      colFsValue,
    ],
  });

  updateFsSummary();
  updateFsPreview();
}

function setText(id, value) {
  const el = document.getElementById(id);
  if (el) el.textContent = value;
}

// ======================================================
// 17. FS SUMMARY & PREVIEW
// ======================================================
function updateFsSummary() {
  const area = document.getElementById("fsSummaryArea");
  if (!area) return;

  const summary = {};
  for (const row of mergedData) {
    const line = row["FS Line Item"] || "";
    if (!summary[line]) summary[line] = 0;
    summary[line] += row["FS Value"] || 0;
  }

  let html = '<table style="width:100%;border-collapse:collapse;">';
  html += '<tr style="border-bottom:2px solid var(--border-color);"><th style="text-align:left;padding:8px;">รายการ</th><th style="text-align:right;padding:8px;">FS Value</th></tr>';
  for (const [line, total] of Object.entries(summary).sort()) {
    const isUnmapped = line === "ไม่จัดประเภท (Unmapped)";
    const bg = isUnmapped && total > 0 ? "background:rgba(239,68,68,0.15);" : "";
    html += `<tr style="${bg}border-bottom:1px solid var(--border-color);">
      <td style="padding:6px 8px;font-size:0.85rem;">${line}</td>
      <td style="padding:6px 8px;text-align:right;font-size:0.85rem;">${formatMoney(total)}</td></tr>`;
  }
  html += "</table>";
  area.innerHTML = html;
}

function updateFsPreview() {
  const activeItems = getFsLineItems();
  const bsStructure = getBsStructure();
  const plStructure = getPlStructure();

  // Compute FS summary from merged data
  const fsSummary = {};
  for (const row of mergedData) {
    const line = row["FS Line Item"] || "";
    if (line === "ไม่จัดประเภท (Unmapped)") continue;
    if (!fsSummary[line]) fsSummary[line] = 0;
    fsSummary[line] += row["FS Value"] || 0;
  }

  // Split BS vs PL
  const plSummary = {};
  const bsSummary = {};
  for (const [k, v] of Object.entries(fsSummary)) {
    const rules = activeItems[k];
    if (!rules) continue;
    if (rules.side.startsWith("pl")) plSummary[k] = v;
    else bsSummary[k] = v;
  }

  // Compute PL first (need net profit for BS)
  const plComputed = buildFsFromMapping(plSummary, plStructure);
  const netProfit = plComputed["กำไร(ขาดทุน)สุทธิ"] || 0;

  // Adjust BS: retained earnings + net profit, share capital
  const retainedLabel = findRetainedEarningsLabel();
  const shareCapLabel = findShareCapitalLabel();
  const paidShares = parseInt(document.getElementById("paidShares")?.value) || 0;
  const paidPar = parseFloat(document.getElementById("paidPar")?.value) || 0;

  bsSummary[retainedLabel] = (bsSummary[retainedLabel] || 0) + netProfit;
  if (paidShares * paidPar > 0) bsSummary[shareCapLabel] = paidShares * paidPar;

  const bsComputed = buildFsFromMapping(bsSummary, bsStructure);

  // Compute prior year FS for preview if available
  let bsComputedPrior = null;
  let plComputedPrior = null;

  if (priorYearSummary) {
    const plSummaryPrior = {};
    const bsSummaryPriorRaw = {};
    for (const [k, v] of Object.entries(priorYearSummary)) {
      const rules = activeItems[k];
      if (!rules) continue;
      if (rules.side.startsWith("pl")) plSummaryPrior[k] = v;
      else bsSummaryPriorRaw[k] = v;
    }
    plComputedPrior = buildFsFromMapping(plSummaryPrior, plStructure);
    const netProfitPrior = plComputedPrior["กำไร(ขาดทุน)สุทธิ"] || 0;

    const bsSummaryPriorFinal = { ...bsSummaryPriorRaw };
    bsSummaryPriorFinal[retainedLabel] = (bsSummaryPriorFinal[retainedLabel] || 0) + netProfitPrior;
    if (paidShares * paidPar > 0) bsSummaryPriorFinal[shareCapLabel] = paidShares * paidPar;
    bsComputedPrior = buildFsFromMapping(bsSummaryPriorFinal, bsStructure);
  }

  // Render BS Preview
  renderFsStructurePreview("bsPreviewArea", bsStructure, bsComputed, bsComputedPrior, "🏛️", activeItems);

  // Render PL Preview
  renderFsStructurePreview("plPreviewArea", plStructure, plComputed, plComputedPrior, "📈", activeItems);

  // Render EQ Preview
  renderEqPreview(bsComputed, plComputed, netProfit);
}

function renderFsStructurePreview(containerId, structure, computedCurrent, computedPrior, icon, activeItems) {
  const container = document.getElementById(containerId);
  if (!container) return;
  const hasPrior = computedPrior !== null && computedPrior !== undefined;
  const currentYearStr = document.getElementById("currentYear")?.value || "ปีปัจจุบัน";
  const priorYearStr = priorYearLabel || "ปีก่อนหน้า";
  let html = "";

  for (const [rowType, label, note, key] of structure) {
    if (rowType === "spacer") {
      html += '<div class="fs-spacer-row"></div>';
      continue;
    }
    if (rowType === "header") {
      html += `<div class="fs-header-row">${icon} ${label}</div>`;
      continue;
    }
    if (rowType === "subtotal") {
      const valCur = computedCurrent[label] || 0;
      const valPri = hasPrior ? (computedPrior[label] || 0) : 0;
      const formulaParts = key.map(k => k.startsWith("-") ? `- ${k.substring(1)}` : `+ ${k}`);
      const formulaText = formulaParts.join(" ").replace(/^\+ /, "");
      const isGrandTotal = label.includes("รวมหนี้สินและส่วนของเจ้าของ") || label === "รวมสินทรัพย์" || label === "กำไร(ขาดทุน)สุทธิ";

      const valDisplay = hasPrior
        ? `<span>${currentYearStr}: <strong>${formatMoney(valCur)}</strong> บาท</span> | <span style="color:var(--text-secondary);">${priorYearStr}: ${formatMoney(valPri)} บาท</span>`
        : `${formatMoney(valCur)} บาท`;

      if (isGrandTotal) {
        const borderColor = label === "กำไร(ขาดทุน)สุทธิ" ? (valCur >= 0 ? "var(--success-color)" : "var(--error-color)") : "var(--success-color)";
        html += `<div class="fs-grand-total" style="border-left-color:${borderColor}" title="วิธีคำนวณ: ${formulaText}">
          <strong>${label}</strong><br><span class="fs-grand-total-value">${valDisplay}</span></div>`;
      } else {
        const subDisplay = hasPrior
          ? `<strong>∑ ${label}</strong> — ${currentYearStr}: ${formatMoney(valCur)} | ${priorYearStr}: ${formatMoney(valPri)}`
          : `<strong>∑ ${label}</strong> : ${formatMoney(valCur)}`;
        html += `<div class="fs-subtotal-row" title="วิธีคำนวณ: ${formulaText}">${subDisplay}</div>`;
      }
      continue;
    }
    if (rowType === "item") {
      const itemTotalCur = computedCurrent[label] || 0;
      const itemTotalPri = hasPrior ? (computedPrior[label] || 0) : 0;
      // Zero-row suppression (both years zero)
      if (itemTotalCur === 0 && itemTotalPri === 0) continue;

      const noteStr = note ? `<span class="badge">${note}</span>` : "";
      const itemDisplay = hasPrior
        ? `📄 <strong>${label}</strong> ${noteStr} — ${currentYearStr}: <strong>${formatMoney(itemTotalCur)}</strong> | ${priorYearStr}: ${formatMoney(itemTotalPri)}`
        : `📄 <strong>${label}</strong> ${noteStr} : ${formatMoney(itemTotalCur)}`;

      // Build drill-down content
      let drillHtml = "";
      if (Array.isArray(key)) {
        drillHtml += '<div class="fs-drill-detail"><em>คำนวณจาก:</em>';
        for (const ref of key) {
          const cleanRef = ref.replace(/^-/, "");
          const sign = ref.startsWith("-") ? "➖" : "➕";
          const refRows = mergedData.filter(r => r["FS Line Item"] === cleanRef);
          const refTotal = refRows.reduce((sum, r) => sum + (r["FS Value"] || 0), 0);
          drillHtml += `<div style="margin:4px 0;"><strong>${sign} ${cleanRef}</strong> : ${formatMoney(refTotal)}</div>`;
          if (refRows.length > 0) {
            drillHtml += buildAccountTable(refRows);
          }
        }
        drillHtml += "</div>";
      } else {
        const lineRows = mergedData.filter(r => r["FS Line Item"] === key);
        if (lineRows.length > 0) drillHtml = '<div class="fs-drill-detail">' + buildAccountTable(lineRows) + "</div>";
        else drillHtml = '<div class="fs-drill-detail"><em>ยังไม่มีบัญชีที่ผูกกับรายการนี้</em></div>';
      }

      html += `<div class="accordion">
        <div class="accordion-header" onclick="this.parentElement.classList.toggle('open')">
          <span>${itemDisplay}</span>
          <span class="collapsible-icon">▶</span>
        </div>
        <div class="accordion-content">${drillHtml}</div>
      </div>`;
    }
  }
  container.innerHTML = html;
}

function buildAccountTable(rows) {
  let html = '<table style="width:100%;border-collapse:collapse;font-size:0.85rem;margin-top:4px;">';
  html += '<tr style="border-bottom:1px solid var(--border-color);"><th style="text-align:left;padding:4px;">รหัสบัญชี</th><th style="text-align:left;padding:4px;">ชื่อบัญชี</th><th style="text-align:right;padding:4px;">ยอดเงิน</th></tr>';
  for (const r of rows) {
    html += `<tr style="border-bottom:1px solid var(--border-color);">
      <td style="padding:4px;">${r["Account ID"]}</td>
      <td style="padding:4px;">${r["Account Name"]}</td>
      <td style="padding:4px;text-align:right;">${formatMoney(r["FS Value"])}</td></tr>`;
  }
  html += "</table>";
  return html;
}

function renderEqPreview(bsComputed, plComputed, netProfit) {
  const container = document.getElementById("eqPreviewArea");
  if (!container) return;

  const currentYear = document.getElementById("currentYear")?.value || "พ.ศ. 2568";
  const paidShares = parseInt(document.getElementById("paidShares")?.value) || 0;
  const paidPar = parseFloat(document.getElementById("paidPar")?.value) || 0;
  const retainedLabel = findRetainedEarningsLabel();

  const openCap = paidShares * paidPar;
  const retainedFromBs = bsComputed[retainedLabel] || 0;
  const openRe = retainedFromBs - netProfit;

  let html = `<h4>⚖️ งบการเปลี่ยนแปลงส่วนของเจ้าของ</h4>
  <p>สำหรับปีสิ้นสุดวันที่ 31 ธันวาคม ${currentYear}</p>
  <table style="width:100%;border-collapse:collapse;margin-top:1rem;">
  <tr style="border-bottom:2px solid var(--border-color);">
    <th style="text-align:left;padding:8px;width:40%;">รายการ</th>
    <th style="text-align:right;padding:8px;">ทุนที่ออกและเรียกชำระแล้ว</th>
    <th style="text-align:right;padding:8px;">กำไร(ขาดทุน)สะสม</th>
    <th style="text-align:right;padding:8px;">รวมส่วนของเจ้าของ</th>
  </tr>`;

  const rows = [
    { label: `ยอด ณ วันต้นปี ${currentYear}`, cap: openCap, ret: openRe, bold: true },
    { label: "กำไร(ขาดทุน)สุทธิสำหรับปี", cap: 0, ret: netProfit, bold: false },
    { label: `ยอดคงเหลือ ณ สิ้นปี ${currentYear}`, cap: openCap, ret: openRe + netProfit, bold: true },
  ];

  for (const r of rows) {
    const style = r.bold ? "font-weight:700;" : "";
    html += `<tr style="border-bottom:1px solid var(--border-color);${style}">
      <td style="padding:8px;">${r.label}</td>
      <td style="padding:8px;text-align:right;">${formatMoney(r.cap)}</td>
      <td style="padding:8px;text-align:right;">${formatMoney(r.ret)}</td>
      <td style="padding:8px;text-align:right;">${formatMoney(r.cap + r.ret)}</td>
    </tr>`;
  }
  html += "</table>";
  container.innerHTML = html;
}

// ======================================================
// 18. EXPORT — RECONCILIATION REPORT
// ======================================================
async function exportReconciliation() {
  if (mergedData.length === 0) {
    showToast("ยังไม่มีข้อมูล กรุณากระทบยอดก่อน", "error");
    return;
  }
  const wb = new ExcelJS.Workbook();
  const allCols = [
    "Account ID", "Account Name", "GL Net Balance", "TB Net Balance",
    "Difference", "GL Brought Forward", "TB Brought Forward Net", "BF Difference",
    "PL Debit", "PL Credit", "BS Debit", "BS Credit",
    "FS Line Item", "FS Value", "Status",
  ];

  function addSheet(name, data) {
    const ws = wb.addWorksheet(name);
    ws.columns = allCols.map(c => ({ header: c, key: c, width: 18 }));
    for (const row of data) {
      const obj = {};
      for (const c of allCols) obj[c] = row[c] ?? "";
      ws.addRow(obj);
    }
  }

  addSheet("All Accounts", mergedData);
  addSheet("Discrepancies", mergedData.filter(d => d.Status === "Amount Mismatch"));
  addSheet("Missing in TB", mergedData.filter(d => d.Status === "Missing in TB"));
  addSheet("Missing in GL", mergedData.filter(d => d.Status === "Missing in GL"));
  addSheet("Top 10 Suspects", [...mergedData].sort((a, b) => b["Max Absolute Balance"] - a["Max Absolute Balance"]).slice(0, 10));

  const buffer = await wb.xlsx.writeBuffer();
  saveAs(new Blob([buffer]), "Reconciliation_Report.xlsx");
  showToast("ดาวน์โหลด Reconciliation Report สำเร็จ!", "success");
}

// ======================================================
// 19. EXPORT — FINANCIAL STATEMENTS
// ======================================================
async function exportFinancialStatements() {
  if (mergedData.length === 0) {
    showToast("ยังไม่มีข้อมูล กรุณากระทบยอดก่อน", "error");
    return;
  }

  const companyName = document.getElementById("companyName")?.value || "บริษัท";
  const currentYear = document.getElementById("currentYear")?.value || "พ.ศ. 2568";
  const regShares = parseInt(document.getElementById("regShares")?.value) || 0;
  const regPar = parseFloat(document.getElementById("regPar")?.value) || 0;
  const paidShares = parseInt(document.getElementById("paidShares")?.value) || 0;
  const paidPar = parseFloat(document.getElementById("paidPar")?.value) || 0;

  const activeItems = getFsLineItems();
  const bsStructure = getBsStructure();
  const plStructure = getPlStructure();
  const eqStructure = getEqStructure();

  // Compute FS values
  const fsSummary = {};
  for (const row of mergedData) {
    const line = row["FS Line Item"];
    if (!line || line === "ไม่จัดประเภท (Unmapped)") continue;
    if (!fsSummary[line]) fsSummary[line] = 0;
    fsSummary[line] += row["FS Value"] || 0;
  }

  const plSummary = {}, bsSummaryRaw = {};
  for (const [k, v] of Object.entries(fsSummary)) {
    const rules = activeItems[k];
    if (!rules) continue;
    if (rules.side.startsWith("pl")) plSummary[k] = v;
    else bsSummaryRaw[k] = v;
  }

  const plComputed = buildFsFromMapping(plSummary, plStructure);
  const netProfit = plComputed["กำไร(ขาดทุน)สุทธิ"] || 0;

  const retainedLabel = findRetainedEarningsLabel();
  const shareCapLabel = findShareCapitalLabel();
  const bsSummaryFinal = { ...bsSummaryRaw };
  bsSummaryFinal[retainedLabel] = (bsSummaryFinal[retainedLabel] || 0) + netProfit;
  if (paidShares * paidPar > 0) bsSummaryFinal[shareCapLabel] = paidShares * paidPar;
  const bsComputed = buildFsFromMapping(bsSummaryFinal, bsStructure);

  // Compute prior year FS if available
  let bsComputedPrior = null;
  let plComputedPrior = null;

  if (priorYearSummary) {
    const plSummaryPrior = {};
    const bsSummaryPriorRaw = {};
    for (const [k, v] of Object.entries(priorYearSummary)) {
      const rules = activeItems[k];
      if (!rules) continue;
      if (rules.side.startsWith("pl")) plSummaryPrior[k] = v;
      else bsSummaryPriorRaw[k] = v;
    }
    plComputedPrior = buildFsFromMapping(plSummaryPrior, plStructure);
    const netProfitPrior = plComputedPrior["กำไร(ขาดทุน)สุทธิ"] || 0;

    const bsSummaryPriorFinal = { ...bsSummaryPriorRaw };
    bsSummaryPriorFinal[retainedLabel] = (bsSummaryPriorFinal[retainedLabel] || 0) + netProfitPrior;
    if (paidShares * paidPar > 0) bsSummaryPriorFinal[shareCapLabel] = paidShares * paidPar;
    bsComputedPrior = buildFsFromMapping(bsSummaryPriorFinal, bsStructure);
  }

  const wb = new ExcelJS.Workbook();
  const NUM_FMT = '#,##0.00;[Red]-#,##0.00';

  // Helper: write FS sheet
  function writeFsSheet(ws, structure, computedCurrent, computedPrior, title, periodLabel) {
    const hasPrior = computedPrior !== null && computedPrior !== undefined;
    ws.getColumn(1).width = 42;
    ws.getColumn(2).width = 10;
    ws.getColumn(3).width = 18;
    if (hasPrior) ws.getColumn(4).width = 18;

    let r = 1;
    ws.getCell(r, 1).value = companyName;
    ws.getCell(r, 1).font = { name: "Cordia New", size: 14, bold: true };
    r++;
    ws.getCell(r, 1).value = title;
    ws.getCell(r, 1).font = { name: "Cordia New", size: 14, bold: true };
    r++;
    ws.getCell(r, 1).value = periodLabel;
    ws.getCell(r, 1).font = { name: "Cordia New", size: 14 };
    r++;
    ws.getCell(r, 1).value = "หน่วย : บาท";
    ws.getCell(r, 1).font = { name: "Cordia New", size: 14, italic: true };
    r++;

    ws.getCell(r, 3).value = currentYear;
    ws.getCell(r, 3).font = { name: "Cordia New", size: 14, bold: true };
    ws.getCell(r, 3).alignment = { horizontal: "center" };

    if (hasPrior) {
      ws.getCell(r, 4).value = priorYearLabel;
      ws.getCell(r, 4).font = { name: "Cordia New", size: 14, bold: true };
      ws.getCell(r, 4).alignment = { horizontal: "center" };
    }
    r++;

    for (const [rowType, label, note, key] of structure) {
      if (rowType === "spacer") { r++; continue; }
      if (rowType === "item" && key) {
        const valCur = computedCurrent[label] || 0;
        const valPri = hasPrior ? (computedPrior[label] || 0) : 0;
        if (valCur === 0 && valPri === 0) continue; // Zero-row suppression (both years zero)
      }

      const isHeader = rowType === "header";
      const isSubtotal = rowType === "subtotal";
      const cell = ws.getCell(r, 1);
      cell.value = label;
      cell.font = { name: "Cordia New", size: 14, bold: isHeader || isSubtotal };
      cell.alignment = { indent: isHeader ? 0 : (isSubtotal ? 1 : 2) };

      if (isHeader) {
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDDEEFF" } };
      }

      if (note) {
        const nc = ws.getCell(r, 2);
        nc.value = note;
        nc.alignment = { horizontal: "center" };
        nc.font = { name: "Cordia New", size: 14 };
      }

      if (rowType === "item" || rowType === "subtotal") {
        const valCur = computedCurrent[label];
        if (valCur !== null && valCur !== undefined) {
          const vc = ws.getCell(r, 3);
          vc.value = valCur;
          vc.numFmt = NUM_FMT;
          vc.alignment = { horizontal: "right" };
          vc.font = { name: "Cordia New", size: 14, bold: isSubtotal };
          if (isSubtotal) vc.border = { top: { style: "double" } };
        }

        if (hasPrior) {
          const valPri = computedPrior[label];
          if (valPri !== null && valPri !== undefined) {
            const vp = ws.getCell(r, 4);
            vp.value = valPri;
            vp.numFmt = NUM_FMT;
            vp.alignment = { horizontal: "right" };
            vp.font = { name: "Cordia New", size: 14, bold: isSubtotal };
            if (isSubtotal) vp.border = { top: { style: "double" } };
          }
        }
      }
      r++;
    }

    ws.getCell(r, 1).value = "หมายเหตุประกอบงบการเงินเป็นส่วนหนึ่งของงบการเงินนี้";
    ws.getCell(r, 1).font = { name: "Cordia New", size: 14, italic: true };
  }

  // BS Sheet
  const wsBs = wb.addWorksheet("งบฐานะการเงิน");
  writeFsSheet(wsBs, bsStructure, bsComputed, bsComputedPrior, "งบฐานะการเงิน", `ณ วันที่ 31 ธันวาคม ${currentYear}`);

  // PL Sheet
  const wsPl = wb.addWorksheet("งบกำไรขาดทุน");
  writeFsSheet(wsPl, plStructure, plComputed, plComputedPrior, "งบกำไรขาดทุน", `สำหรับปีสิ้นสุดวันที่ 31 ธันวาคม ${currentYear}`);

  // EQ Sheet
  const wsEq = wb.addWorksheet(eqStructure.sheet_title || "งบส่วนของเจ้าของ");
  wsEq.getColumn(1).width = 46;
  wsEq.getColumn(2).width = 2;
  wsEq.getColumn(3).width = 18;
  wsEq.getColumn(4).width = 2;
  wsEq.getColumn(5).width = 18;
  wsEq.getColumn(6).width = 2;
  wsEq.getColumn(7).width = 18;

  let eqR = 1;
  wsEq.getCell(eqR, 1).value = companyName;
  wsEq.getCell(eqR, 1).font = { name: "Cordia New", size: 14, bold: true }; eqR++;
  wsEq.getCell(eqR, 1).value = eqStructure.sheet_title || "งบการเปลี่ยนแปลงส่วนของเจ้าของ";
  wsEq.getCell(eqR, 1).font = { name: "Cordia New", size: 14, bold: true }; eqR++;
  wsEq.getCell(eqR, 1).value = `สำหรับปีสิ้นสุดวันที่ 31 ธันวาคม ${currentYear}`;
  wsEq.getCell(eqR, 1).font = { name: "Cordia New", size: 14 }; eqR++;
  wsEq.getCell(eqR, 1).value = "หน่วย : บาท";
  wsEq.getCell(eqR, 1).font = { name: "Cordia New", size: 14, italic: true }; eqR++;
  eqR++;

  const cols = eqStructure.columns || ["ทุนที่ออกและเรียกชำระแล้ว", "กำไร(ขาดทุน)สะสม", "รวมส่วนของเจ้าของ"];
  const colPositions = [3, 5, 7];
  for (let ci = 0; ci < cols.length; ci++) {
    wsEq.getCell(eqR, colPositions[ci]).value = cols[ci];
    wsEq.getCell(eqR, colPositions[ci]).font = { name: "Cordia New", size: 14, bold: true };
    wsEq.getCell(eqR, colPositions[ci]).alignment = { horizontal: "center" };
  }
  eqR++;

  const openCap = paidShares * paidPar;
  const retainedFromBs = bsComputed[retainedLabel] || 0;
  const openRe = retainedFromBs - netProfit;

  for (const eqRow of (eqStructure.rows || DEFAULT_CONFIG.eq_structure.rows)) {
    const labelTmpl = eqRow.label_template || eqRow.label || "";
    const rowLabel = labelTmpl.replace("{year}", currentYear);
    const isBold = eqRow.type === "opening" || eqRow.type === "closing";

    wsEq.getCell(eqR, 1).value = rowLabel;
    wsEq.getCell(eqR, 1).font = { name: "Cordia New", size: 14, bold: isBold };

    let capVal = eqRow.values.capital;
    if (capVal === "open_capital") capVal = openCap;
    else if (capVal === "close_capital") capVal = openCap;
    else if (capVal === "issued_capital") capVal = 0;
    else if (typeof capVal === "string") capVal = 0;

    let retVal = eqRow.values.retained;
    if (retVal === "open_retained") retVal = openRe;
    else if (retVal === "close_retained") retVal = openRe + netProfit;
    else if (retVal === "net_profit") retVal = netProfit;
    else if (typeof retVal === "string") retVal = 0;

    let totalVal = eqRow.values.total;
    if (totalVal === "auto_sum") totalVal = capVal + retVal;

    if (capVal !== 0) {
      wsEq.getCell(eqR, 3).value = capVal;
      wsEq.getCell(eqR, 3).numFmt = NUM_FMT;
      wsEq.getCell(eqR, 3).alignment = { horizontal: "right" };
    }
    if (retVal !== 0) {
      wsEq.getCell(eqR, 5).value = retVal;
      wsEq.getCell(eqR, 5).numFmt = NUM_FMT;
      wsEq.getCell(eqR, 5).alignment = { horizontal: "right" };
    }
    wsEq.getCell(eqR, 7).value = totalVal;
    wsEq.getCell(eqR, 7).numFmt = NUM_FMT;
    wsEq.getCell(eqR, 7).alignment = { horizontal: "right" };
    eqR++;
  }

  eqR++;
  wsEq.getCell(eqR, 1).value = "หมายเหตุประกอบงบการเงินเป็นส่วนหนึ่งของงบการเงินนี้";
  wsEq.getCell(eqR, 1).font = { name: "Cordia New", size: 14, italic: true };

  const buffer = await wb.xlsx.writeBuffer();
  const safeName = companyName.replace(/[^\w\u0e00-\u0e7f]/g, "_").substring(0, 30);
  saveAs(new Blob([buffer]), `FS_${safeName}.xlsx`);
  showToast("ดาวน์โหลดงบการเงินสำเร็จ!", "success");
}

// ======================================================
// 20. FILE SAVE UTILITY
// ======================================================
function saveAs(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// ======================================================
// 21. CONFIG EDITOR UI
// ======================================================

// Ace Editor instances
let aceEditors = {};

function initConfigEditors() {
  // Initialize Ace Editors
  const editorIds = ["cfgItems", "cfgBs", "cfgPl", "cfgEq"];
  editorIds.forEach(id => {
    if (!aceEditors[id] && document.getElementById(id)) {
      const editor = ace.edit(id);
      editor.setTheme("ace/theme/tomorrow_night_eighties");
      editor.session.setMode("ace/mode/json");
      editor.setOptions({
        fontSize: "14px",
        showPrintMargin: false,
        wrap: true,
        tabSize: 2
      });
      aceEditors[id] = editor;
    }
  });

  const cfg = getActiveConfig();
  setEditorValue("cfgItems", cfg.fs_line_items || DEFAULT_CONFIG.fs_line_items);
  setEditorValue("cfgBs", cfg.bs_structure || DEFAULT_CONFIG.bs_structure);
  setEditorValue("cfgPl", cfg.pl_structure || DEFAULT_CONFIG.pl_structure);
  setEditorValue("cfgEq", cfg.eq_structure || DEFAULT_CONFIG.eq_structure);

  const sourceInfo = document.getElementById("configSourceInfo");
  if (sourceInfo) {
    const hasCustom = localStorage.getItem(CONFIG_STORAGE_KEY) !== null;
    sourceInfo.innerHTML = `📍 Config source: <strong>${hasCustom ? "localStorage (custom)" : "default"}</strong>`;
  }
}

function setEditorValue(id, obj) {
  const editor = aceEditors[id];
  if (editor) {
    editor.setValue(JSON.stringify(obj, null, 2), -1); // -1 moves cursor to start
  }
}

function getEditorValue(id) {
  const editor = aceEditors[id];
  if (!editor) return null;
  return JSON.parse(editor.getValue());
}

function deployConfig() {
  const resultArea = document.getElementById("cfgValidationResults");
  try {
    const parsed = {
      fs_line_items: getEditorValue("cfgItems"),
      bs_structure: getEditorValue("cfgBs"),
      pl_structure: getEditorValue("cfgPl"),
      eq_structure: getEditorValue("cfgEq"),
    };

    const issues = validateConfig(parsed);
    const errors = issues.filter(i => i.severity === "error");
    const warnings = issues.filter(i => i.severity === "warning");

    if (errors.length > 0) {
      let html = '<div class="error-box"><h4>❌ Deploy Failed — Validation Errors</h4>';
      for (const e of errors) html += `<div>❌ ${e.message}</div>`;
      if (warnings.length > 0) {
        html += '<h4>⚠️ Warnings</h4>';
        for (const w of warnings) html += `<div>⚠️ ${w.message}</div>`;
      }
      html += "</div>";
      if (resultArea) resultArea.innerHTML = html;
      showToast("Deploy Failed — มี validation errors", "error");
      return;
    }

    saveConfig(parsed);
    let html = '<div class="success-box"><h4>✅ Deployed Successfully!</h4><div>Config saved to localStorage and activated.</div>';
    if (warnings.length > 0) {
      html += `<h4>⚠️ ${warnings.length} Warning(s)</h4>`;
      for (const w of warnings) html += `<div>⚠️ ${w.message}</div>`;
    }
    html += "</div>";
    if (resultArea) resultArea.innerHTML = html;

    // Update source info
    const sourceInfo = document.getElementById("configSourceInfo");
    if (sourceInfo) sourceInfo.innerHTML = '📍 Config source: <strong>localStorage (deployed)</strong>';

    showToast("✅ Config deployed successfully!", "success");
  } catch (e) {
    if (resultArea) resultArea.innerHTML = `<div class="error-box"><h4>❌ JSON Parse Error</h4><div>${e.message}</div></div>`;
    showToast("JSON Parse Error: " + e.message, "error");
  }
}

function exportConfig() {
  const cfg = getActiveConfig();
  const blob = new Blob([JSON.stringify(cfg, null, 2)], { type: "application/json" });
  saveAs(blob, "fs_config.json");
  showToast("Config exported!", "success");
}

function resetDefaultConfig() {
  if (!confirm("รีเซ็ตกลับเป็นค่า default ทั้งหมด?")) return;
  resetConfig();
  initConfigEditors();
  showToast("✅ Reset to defaults!", "success");
}

function importConfigFile(event) {
  const file = event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function (e) {
    try {
      const loaded = JSON.parse(e.target.result);
      if (typeof loaded !== "object") {
        showToast("Invalid config file", "error");
        return;
      }
      if (loaded.fs_line_items) setEditorValue("cfgItems", loaded.fs_line_items);
      if (loaded.bs_structure) setEditorValue("cfgBs", loaded.bs_structure);
      if (loaded.pl_structure) setEditorValue("cfgPl", loaded.pl_structure);
      if (loaded.eq_structure) setEditorValue("cfgEq", loaded.eq_structure);
      showToast("📂 Config loaded into editors. Click Deploy & Apply to activate.", "success");
    } catch (err) {
      showToast("Invalid JSON file: " + err.message, "error");
    }
  };
  reader.readAsText(file);
  // Reset file input so same file can be loaded again
  event.target.value = "";
}
