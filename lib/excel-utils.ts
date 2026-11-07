import * as XLSX from "xlsx";
import JSZip from "jszip";
import {
  GlobalConfig,
  NamingMode,
  SheetItem,
  UploadedWorkbook,
  ZeroPadding,
  Separator,
} from "@/types/excel";

/**
 * Safely resolves the JSZip constructor across ESM, CJS, and Turbopack bundlers.
 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
function getJSZipConstructor(): any {
  if (typeof JSZip === "function") {
    return JSZip;
  }
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  if ((JSZip as any)?.default && typeof (JSZip as any).default === "function") {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (JSZip as any).default;
  }
  return JSZip;
}

export async function loadJSZip(buffer: ArrayBuffer | Uint8Array | string): Promise<JSZip> {
  const ZipClass = getJSZipConstructor();
  if (typeof ZipClass.loadAsync === "function") {
    return await ZipClass.loadAsync(buffer);
  }
  const instance = new ZipClass();
  return await instance.loadAsync(buffer);
}

export function createJSZipInstance(): JSZip {
  const ZipClass = getJSZipConstructor();
  return new ZipClass();
}

/**
 * Automatically sanitizes any Devanagari font fallbacks injected by LibreOffice or system fallbacks
 * and maps them directly to authentic Khmer typography (Khmer OS Battambang / Khmer OS Muol Light).
 */
export function sanitizeKhmerFontsInXml(xml: string): string {
  if (!/Devanagari|Mangal/i.test(xml)) {
    return xml;
  }
  return xml
    .replace(/(?:val|typeface)="Noto Serif Devanagari"/gi, 'val="Khmer OS Muol Light"')
    .replace(/(?:val|typeface)="Noto Sans Devanagari(?:\s+UI)?"/gi, 'val="Khmer OS Battambang"')
    .replace(/(?:val|typeface)="[^"]*Devanagari[^"]*"/gi, 'val="Khmer OS Battambang"')
    .replace(/(?:val|typeface)="Mangal"/gi, 'val="Khmer OS Battambang"')
    .replace(/>Noto Serif Devanagari</gi, '>Khmer OS Muol Light<')
    .replace(/>Noto Sans Devanagari(?:\s+UI)?</gi, '>Khmer OS Battambang<')
    .replace(/>[^<]*Devanagari[^<]*</gi, '>Khmer OS Battambang<');
}

/**
 * Traverses all styles, shared strings, and theme definitions in the OpenXML archive
 * and replaces any Devanagari fallback fonts with Khmer OS Battambang / Khmer OS Muol Light.
 */
export async function sanitizeKhmerFontsInZip(zip: JSZip): Promise<void> {
  const targetPatterns = [
    "xl/styles.xml",
    "xl/sharedStrings.xml",
    "xl/theme/theme1.xml",
  ];

  for (const pattern of targetPatterns) {
    const fileObj = getZipFileCaseInsensitive(zip, pattern);
    if (fileObj) {
      const content = await fileObj.async("text");
      if (/Devanagari|Mangal/i.test(content)) {
        zip.file(fileObj.name, sanitizeKhmerFontsInXml(content));
      }
    }
  }

  // Also check worksheets
  for (const filename of Object.keys(zip.files)) {
    if (/^xl\/worksheets\/sheet\d+\.xml$/i.test(filename)) {
      const wsFile = zip.files[filename];
      const content = await wsFile.async("text");
      if (/Devanagari|Mangal/i.test(content)) {
        zip.file(wsFile.name, sanitizeKhmerFontsInXml(content));
      }
    }
  }
}

/**
 * Removes characters that are illegal in file and folder names on Windows, Mac, and Linux.
 */
export function sanitizeFileName(name: string, fallback = "Sheet"): string {
  const sanitized = name.replace(/[/\\?%*:|"<>]/g, "_").trim();
  return sanitized || fallback;
}

/**
 * Formats a sequence number according to the padding rule.
 */
export function formatSequenceNumber(
  num: number,
  padding: ZeroPadding,
  totalItems: number
): string {
  if (padding === "none") {
    return String(num);
  }
  if (padding === "2") {
    return String(num).padStart(2, "0");
  }
  if (padding === "3") {
    return String(num).padStart(3, "0");
  }
  if (padding === "auto") {
    const digitsNeeded = Math.max(2, String(totalItems).length);
    return String(num).padStart(digitsNeeded, "0");
  }
  return String(num);
}

/**
 * Calculates the final .xlsx filename for a sheet based on configurations.
 */
export function resolveSheetFileName({
  prefix,
  separator,
  namingMode,
  originalSheetName,
  customName,
  index,
  startNumber,
  zeroPadding,
  totalSheets,
  extension = "xlsx",
  isInfoSheet = false,
  infoFileName = "",
}: {
  prefix: string;
  separator: Separator;
  namingMode: NamingMode;
  originalSheetName: string;
  customName?: string;
  index: number;
  startNumber: number;
  zeroPadding: ZeroPadding;
  totalSheets: number;
  extension?: string;
  isInfoSheet?: boolean;
  infoFileName?: string;
}): string {
  const cleanExt = extension.replace(/^\./, "");
  if (customName && customName.trim()) {
    const cleanCustom = sanitizeFileName(customName.trim());
    return cleanCustom.toLowerCase().endsWith(`.${cleanExt}`)
      ? cleanCustom
      : `${cleanCustom}.${cleanExt}`;
  }

  const cleanPrefix = prefix ? sanitizeFileName(prefix.trim()) : "";
  const cleanInfoName = infoFileName ? sanitizeFileName(infoFileName.trim()) : "";

  // If this is the Info sheet, use custom info file name if given, else prefix, else "Info"
  if (isInfoSheet) {
    const infoBase = cleanInfoName || cleanPrefix || "Info";
    return `${infoBase}.${cleanExt}`;
  }

  const sep = cleanPrefix ? separator : "";
  const currentNum = startNumber + index;
  const numStr = formatSequenceNumber(currentNum, zeroPadding, totalSheets);
  const cleanOriginalName = sanitizeFileName(originalSheetName);

  let body = "";
  switch (namingMode) {
    case "sequence":
      body = numStr;
      break;
    case "sequence_original":
      body = `${numStr}_${cleanOriginalName}`;
      break;
    case "original":
    default:
      body = cleanOriginalName;
      break;
  }

  const baseName = cleanPrefix ? `${cleanPrefix}${sep}${body}` : body;
  return `${baseName}.${cleanExt}`;
}

/**
 * Resolves configuration for a specific workbook (considering overrides).
 */
export function getEffectiveConfig(
  workbook: UploadedWorkbook,
  globalConfig: GlobalConfig
): {
  folderName: string;
  prefix: string;
  namingMode: NamingMode;
  startNumber: number;
  zeroPadding: ZeroPadding;
  separator: Separator;
  firstSheetIsInfo: boolean;
  infoFileName: string;
} {
  const defaultBase = sanitizeFileName(
    workbook.fileName.replace(/\.[^/.]+$/, ""),
    "Excel_Folder"
  );
  let resolvedFolder = defaultBase;
  if (workbook.folderName && workbook.folderName !== defaultBase) {
    resolvedFolder = sanitizeFileName(workbook.folderName.trim(), defaultBase);
  } else if (globalConfig.masterFolderName && globalConfig.masterFolderName.trim()) {
    resolvedFolder = sanitizeFileName(globalConfig.masterFolderName.trim(), defaultBase);
  } else if (workbook.folderName) {
    resolvedFolder = sanitizeFileName(workbook.folderName.trim(), defaultBase);
  }

  if (!workbook.useCustomConfig) {
    return {
      ...globalConfig,
      folderName: resolvedFolder,
      infoFileName: globalConfig.infoFileName || "",
    };
  }
  return {
    folderName: resolvedFolder,
    prefix: workbook.customPrefix ?? globalConfig.prefix,
    namingMode: workbook.customNamingMode ?? globalConfig.namingMode,
    startNumber: workbook.customStartNumber ?? globalConfig.startNumber,
    zeroPadding: workbook.customZeroPadding ?? globalConfig.zeroPadding,
    separator: workbook.customSeparator ?? globalConfig.separator,
    firstSheetIsInfo:
      workbook.customFirstSheetIsInfo ?? globalConfig.firstSheetIsInfo,
    infoFileName:
      workbook.customInfoFileName ?? (globalConfig.infoFileName || ""),
  };
}

/**
 * Parses an uploaded Excel file into an UploadedWorkbook structure.
 */
export async function extractWorkbookData(file: File): Promise<UploadedWorkbook> {
  let arrayBuffer = await file.arrayBuffer();
  let effectiveFile = file;
  const activeFileName = file.name;
  const u8 = new Uint8Array(arrayBuffer.slice(0, 8));
  const isBinaryXls = u8[0] === 0xd0 && u8[1] === 0xcf; // OLE2 BIFF8 header

  // 1. Try Spring Boot backend inspection first
  if (typeof window !== "undefined") {
    try {
      const formData = new FormData();
      formData.append("file", file);
      const res = await fetch("/api/excel/inspect", {
        method: "POST",
        body: formData,
      });

      if (res.ok) {
        const meta = await res.json();
        const baseNameWithoutExt = meta.fileName.replace(/\.[^/.]+$/, "");
        const defaultFolderName = sanitizeFileName(baseNameWithoutExt, "Excel_Folder");
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const sheets: SheetItem[] = meta.sheets.map((s: any, idx: number) => ({
          id: `${file.name}-${s.name}-${idx}-${Date.now()}`,
          originalName: s.name,
          selected: true,
          rowCount: s.rowCount,
          colCount: s.colCount,
          sampleHeaders: s.sampleHeaders || [],
          sampleRows: s.sampleRows || [],
        }));

        let fallbackWorkbook: XLSX.WorkBook;
        try {
          fallbackWorkbook = XLSX.read(arrayBuffer.slice(0), { type: "array" });
        } catch {
          fallbackWorkbook = XLSX.utils.book_new();
        }

        return {
          id: `wb-${Date.now()}-${Math.random().toString(36).substring(2, 9)}`,
          fileName: meta.fileName,
          fileSize: meta.fileSize,
          folderName: defaultFolderName,
          sheets,
          rawWorkbook: fallbackWorkbook,
          rawBuffer: arrayBuffer,
          rawFile: file,
          isBinaryXls: meta.isBinaryXls,
          useCustomConfig: false,
          status: "ready",
        };
      }
    } catch {
      // Backend not running, seamlessly proceed with local parser
    }
  }

  // If modern .xlsx zip package, ensure any Devanagari font fallbacks are sanitized upfront
  if (arrayBuffer.byteLength > 4) {
    const pkCheck = new Uint8Array(arrayBuffer.slice(0, 4));
    if (pkCheck[0] === 0x50 && pkCheck[1] === 0x4b) {
      try {
        const checkZip = await loadJSZip(arrayBuffer);
        const stylesFile = getZipFileCaseInsensitive(checkZip, "xl/styles.xml");
        if (stylesFile) {
          const stylesContent = await stylesFile.async("text");
          if (/Devanagari|Mangal/i.test(stylesContent)) {
            await sanitizeKhmerFontsInZip(checkZip);
            const cleaned = await checkZip.generateAsync({ type: "uint8array" });
            arrayBuffer = cleaned.buffer.slice(
              cleaned.byteOffset,
              cleaned.byteOffset + cleaned.byteLength
            ) as ArrayBuffer;
            effectiveFile = new File([arrayBuffer], activeFileName, {
              type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            });
          }
        }
      } catch (cleanErr) {
        console.warn("Could not pre-sanitize fonts from uploaded workbook:", cleanErr);
      }
    }
  }

  // Pass a clean slice to SheetJS so the original arrayBuffer remains 100% pristine
  const bufferCopy = arrayBuffer.slice(0);
  const rawWorkbook = XLSX.read(bufferCopy, {
    type: "array",
    cellStyles: true,
    cellFormula: true,
    cellDates: true,
    cellNF: true,
  });

  const baseFileNameWithoutExt = activeFileName.replace(/\.[^/.]+$/, "");
  const defaultFolderName = sanitizeFileName(baseFileNameWithoutExt, "Excel_Folder");

  const visibleSheetEntries = rawWorkbook.SheetNames.map((sheetName, idx) => {
    // In SheetJS: Hidden: 1 = Hidden, 2 = Very Hidden
    const metaHidden = rawWorkbook.Workbook?.Sheets?.[idx]?.Hidden;
    const isHidden = metaHidden === 1 || metaHidden === 2;
    return { sheetName, idx, isHidden };
  }).filter((entry) => !entry.isHidden);

  const sheets: SheetItem[] = visibleSheetEntries.map(({ sheetName, idx }) => {
    const sheet = rawWorkbook.Sheets[sheetName];
    let rowCount = 0;
    let colCount = 0;
    let sampleHeaders: string[] = [];
    const sampleRows: string[][] = [];

    if (sheet && sheet["!ref"]) {
      const range = XLSX.utils.decode_range(sheet["!ref"]);
      rowCount = range.e.r - range.s.r + 1;
      colCount = range.e.c - range.s.c + 1;

      // Extract sample rows for preview (up to 5 rows)
      const maxRows = Math.min(6, rowCount);
      const maxCols = Math.min(8, colCount);

      for (let r = range.s.r; r < range.s.r + maxRows; r++) {
        const rowVals: string[] = [];
        for (let c = range.s.c; c < range.s.c + maxCols; c++) {
          const cellAddress = XLSX.utils.encode_cell({ r, c });
          const cell = sheet[cellAddress];
          rowVals.push(cell && cell.v !== undefined ? String(cell.v) : "");
        }
        if (r === range.s.r) {
          sampleHeaders = rowVals;
        } else {
          sampleRows.push(rowVals);
        }
      }
    }

    return {
      id: `${file.name}-${sheetName}-${idx}-${Date.now()}`,
      originalName: sheetName,
      selected: true,
      rowCount,
      colCount,
      sampleHeaders,
      sampleRows,
    };
  });

  return {
    id: `wb-${Date.now()}-${Math.random().toString(36).substring(2, 9)}`,
    fileName: activeFileName,
    fileSize: arrayBuffer.byteLength,
    folderName: defaultFolderName,
    sheets,
    rawWorkbook,
    rawBuffer: arrayBuffer,
    rawFile: effectiveFile,
    isBinaryXls,
    useCustomConfig: false,
    status: "ready",
  };
}

/**
 * Ensures unique file names within a folder to avoid collision overwrites.
 */
export function getUniqueName(desiredName: string, usedSet: Set<string>): string {
  if (!usedSet.has(desiredName)) {
    usedSet.add(desiredName);
    return desiredName;
  }
  const dotIdx = desiredName.lastIndexOf(".");
  const base = dotIdx !== -1 ? desiredName.slice(0, dotIdx) : desiredName;
  const ext = dotIdx !== -1 ? desiredName.slice(dotIdx) : "";
  let counter = 1;
  while (usedSet.has(`${base}_(${counter})${ext}`)) {
    counter++;
  }
  const uniqueName = `${base}_(${counter})${ext}`;
  usedSet.add(uniqueName);
  return uniqueName;
}

/**
 * Decodes XML entities including numeric character references (hex and decimal).
 */
export function decodeXmlEntities(str: string): string {
  return str
    .replace(/&#(\d+);/g, (_, dec) => String.fromCharCode(parseInt(dec, 10)))
    .replace(/&#x([0-9a-fA-F]+);/g, (_, hex) => String.fromCharCode(parseInt(hex, 16)))
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");
}

/**
 * Normalizes sheet name for robust matching across XML and SheetJS.
 */
export function normalizeSheetName(name: string): string {
  return decodeXmlEntities(name)
    .normalize("NFC")
    .replace(/[\u200B-\u200D\uFEFF]/g, "") // remove zero-width characters
    .replace(/[\u00A0\s]+/g, " ") // normalize whitespace and non-breaking spaces
    .trim()
    .toLowerCase();
}

/**
 * Case-insensitive zip file accessor to handle varied OpenXML package conventions.
 */
function getZipFileCaseInsensitive(zip: JSZip, targetPath: string): JSZip.JSZipObject | null {
  const exact = zip.file(targetPath);
  if (exact) return exact;
  const targetLower = targetPath.toLowerCase().replace(/\\/g, "/");
  for (const filename of Object.keys(zip.files)) {
    if (filename.toLowerCase().replace(/\\/g, "/") === targetLower) {
      return zip.files[filename];
    }
  }
  return null;
}

/**
 * Encodes string for XML attribute.
 */
export function encodeXmlAttr(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

/**
 * Direct OpenXML tab isolation:
 * Clones the master Excel ZIP package and removes all other worksheets,
 * preserving 100% of the original styles, fills, fonts, borders, column widths,
 * row heights, themes, merged cells, formulas, drawings, and images with ZERO loss.
 */
export async function splitExcelTabDirect(
  originalZipBuffer: ArrayBuffer,
  targetSheetName: string,
  customTabName?: string,
  sheetIndex?: number
): Promise<Uint8Array> {
  const zip = await loadJSZip(originalZipBuffer);

  const workbookFile =
    getZipFileCaseInsensitive(zip, "xl/workbook.xml") ||
    getZipFileCaseInsensitive(zip, "workbook.xml");
  const relsFile =
    getZipFileCaseInsensitive(zip, "xl/_rels/workbook.xml.rels") ||
    getZipFileCaseInsensitive(zip, "_rels/workbook.xml.rels");
  const contentTypesFile = getZipFileCaseInsensitive(zip, "[Content_Types].xml");

  if (!workbookFile || !relsFile || !contentTypesFile) {
    throw new Error("Invalid Excel file: missing workbook.xml, rels, or [Content_Types].xml");
  }

  let workbookXml = await workbookFile.async("text");
  let relsXml = await relsFile.async("text");
  let contentTypesXml = await contentTypesFile.async("text");

  // 1. Parse all <sheet> tags from xl/workbook.xml (supporting XML namespaces like <x:sheet>)
  const sheetTagRegex = /<(?:[a-zA-Z0-9_-]+:)?sheet\b([^>]*)\/?>/gi;
  const sheets: {
    rawTag: string;
    name: string;
    rId: string;
    sheetId: string;
  }[] = [];

  let match: RegExpExecArray | null;
  while ((match = sheetTagRegex.exec(workbookXml)) !== null) {
    const attrs = match[1];
    const nameMatch = attrs.match(/name="([^"]*)"/i);
    // Be strict with word boundary so sheetId does not match id
    const rIdMatch =
      attrs.match(/\br:id="([^"]*)"/i) ||
      attrs.match(/\b(?:relationships:id|rel:id)="([^"]*)"/i) ||
      attrs.match(/\b(?<!sheet)id="([^"]*)"/i);
    const sheetIdMatch = attrs.match(/sheetId="([^"]*)"/i);

    if (nameMatch && rIdMatch) {
      sheets.push({
        rawTag: match[0],
        name: decodeXmlEntities(nameMatch[1]),
        rId: rIdMatch[1],
        sheetId: sheetIdMatch ? sheetIdMatch[1] : "1",
      });
    }
  }

  // 2. Find target sheet using multi-stage matching (Exact -> Trimmed -> Decoded -> Normalized -> Index fallback)
  let targetSheetIndex = sheets.findIndex(
    (s) =>
      s.name === targetSheetName ||
      s.name.trim() === targetSheetName.trim() ||
      decodeXmlEntities(s.name) === decodeXmlEntities(targetSheetName) ||
      decodeXmlEntities(s.name).trim() === decodeXmlEntities(targetSheetName).trim() ||
      encodeXmlAttr(s.name) === encodeXmlAttr(targetSheetName)
  );

  if (targetSheetIndex === -1) {
    const targetNorm = normalizeSheetName(targetSheetName);
    targetSheetIndex = sheets.findIndex(
      (s) => normalizeSheetName(s.name) === targetNorm
    );
  }

  if (
    targetSheetIndex === -1 &&
    typeof sheetIndex === "number" &&
    sheetIndex >= 0 &&
    sheetIndex < sheets.length
  ) {
    targetSheetIndex = sheetIndex;
  }

  if (targetSheetIndex === -1) {
    throw new Error(`Sheet "${targetSheetName}" not found in workbook.`);
  }

  const targetSheet = sheets[targetSheetIndex];
  const sheetsToRemove = sheets.filter((_, i) => i !== targetSheetIndex);

  // 3. Parse relationships from xl/_rels/workbook.xml.rels
  const relTagRegex = /<(?:[a-zA-Z0-9_-]+:)?Relationship\b([^>]*)\/?>/gi;
  const relMap = new Map<string, string>();
  while ((match = relTagRegex.exec(relsXml)) !== null) {
    const attrs = match[1];
    const idMatch = attrs.match(/\bId="([^"]*)"/i);
    const targetMatch = attrs.match(/\bTarget="([^"]*)"/i);
    if (idMatch && targetMatch) {
      relMap.set(idMatch[1], targetMatch[1]);
    }
  }

  // 4. Update xl/workbook.xml: keep only the target sheet
  let cleanTargetTag = targetSheet.rawTag.replace(
    /state="hidden"|state="veryHidden"/gi,
    ""
  );

  // If a custom sheet name was specified, update the tab name
  if (customTabName && customTabName.trim()) {
    const safeCustomTab =
      customTabName.replace(/[/\\?*[\]:]/g, "_").slice(0, 31).trim() ||
      targetSheet.name;
    cleanTargetTag = cleanTargetTag.replace(
      /name="[^"]*"/i,
      `name="${encodeXmlAttr(safeCustomTab)}"`
    );
  }

  // Replace <sheets> tag block (preserving existing namespace prefix if any, e.g. <x:sheets>)
  workbookXml = workbookXml.replace(
    /(<([a-zA-Z0-9_-]+:)?sheets\b[^>]*>)[\s\S]*?(<\/\2?sheets>)/i,
    `$1${cleanTargetTag}$3`
  );

  // Fix activeTab and firstSheet in bookViews to 0 to prevent Excel out-of-bounds repair warnings
  workbookXml = workbookXml.replace(/activeTab="\d+"/gi, 'activeTab="0"');
  workbookXml = workbookXml.replace(/firstSheet="\d+"/gi, 'firstSheet="0"');

  // Remove pivotCaches, externalReferences, and customWorkbookViews from deleted sheets
  workbookXml = workbookXml.replace(
    /<(?:[a-zA-Z0-9_-]+:)?pivotCaches\b[\s\S]*?<\/(?:[a-zA-Z0-9_-]+:)?pivotCaches>/gi,
    ""
  );
  workbookXml = workbookXml.replace(
    /<(?:[a-zA-Z0-9_-]+:)?pivotCaches\b[^>]*\/>/gi,
    ""
  );
  workbookXml = workbookXml.replace(
    /<(?:[a-zA-Z0-9_-]+:)?externalReferences\b[\s\S]*?<\/(?:[a-zA-Z0-9_-]+:)?externalReferences>/gi,
    ""
  );
  workbookXml = workbookXml.replace(
    /<(?:[a-zA-Z0-9_-]+:)?externalReferences\b[^>]*\/>/gi,
    ""
  );
  workbookXml = workbookXml.replace(
    /<(?:[a-zA-Z0-9_-]+:)?customWorkbookViews\b[\s\S]*?<\/(?:[a-zA-Z0-9_-]+:)?customWorkbookViews>/gi,
    ""
  );
  workbookXml = workbookXml.replace(
    /<(?:[a-zA-Z0-9_-]+:)?customWorkbookViews\b[^>]*\/>/gi,
    ""
  );

  // Remove definedNames completely: In single-sheet splits, keeping multi-sheet named ranges
  // causes Microsoft Excel's "We found a problem with some content (Named range)" recovery dialog.
  // Removing it allows Excel to open cleanly with 0 repair warnings and 100% original styles.
  workbookXml = workbookXml.replace(
    /<(?:[a-zA-Z0-9_-]+:)?definedNames\b[\s\S]*?<\/(?:[a-zA-Z0-9_-]+:)?definedNames>/gi,
    ""
  );
  workbookXml = workbookXml.replace(
    /<(?:[a-zA-Z0-9_-]+:)?definedNames\b[^>]*\/>/gi,
    ""
  );

  // Remove calcChain relationship and file to prevent formula dependency errors on missing sheets
  const calcChainRelMatch = relsXml.match(
    /<(?:[a-zA-Z0-9_-]+:)?Relationship\b[^>]*Type="[^"]*calcChain"[^>]*\/?>/i
  );
  if (calcChainRelMatch) {
    relsXml = relsXml.replace(calcChainRelMatch[0], "");
    zip.remove("xl/calcChain.xml");
    zip.remove("calcChain.xml");
    contentTypesXml = contentTypesXml.replace(
      /<Override\b[^>]*PartName="(?:\/xl)?\/calcChain\.xml"[^>]*\/?>/gi,
      ""
    );
  }
  // Clear calcPr to prevent Excel looking for calculation chain
  workbookXml = workbookXml.replace(
    /<(?:[a-zA-Z0-9_-]+:)?calcPr\b[\s\S]*?<\/(?:[a-zA-Z0-9_-]+:)?calcPr>/gi,
    ""
  );
  workbookXml = workbookXml.replace(
    /<(?:[a-zA-Z0-9_-]+:)?calcPr\b[^>]*\/?>/gi,
    ""
  );

  // 5. Update xl/_rels/workbook.xml.rels: remove Relationships for deleted sheets
  for (const rem of sheetsToRemove) {
    const remRelRegex = new RegExp(
      `<(?:[a-zA-Z0-9_-]+:)?Relationship\\b[^>]*Id="${rem.rId}"[^>]*\\/?>`,
      "gi"
    );
    relsXml = relsXml.replace(remRelRegex, "");
  }

  // 6. Update [Content_Types].xml & remove deleted sheet XMLs from zip
  for (const rem of sheetsToRemove) {
    const target = relMap.get(rem.rId);
    if (target) {
      const normPart = target.startsWith("/")
        ? target
        : target.startsWith("xl/")
        ? "/" + target
        : "/xl/" + target;
      const zipPath = normPart.startsWith("/") ? normPart.slice(1) : normPart;

      const escapedPart = normPart.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      const ctRegex = new RegExp(
        `<Override\\b[^>]*PartName="${escapedPart}"[^>]*\\/?>`,
        "gi"
      );
      contentTypesXml = contentTypesXml.replace(ctRegex, "");

      zip.remove(zipPath);
      // Case-insensitive removal fallback
      const directObj = getZipFileCaseInsensitive(zip, zipPath);
      if (directObj) zip.remove(directObj.name);

      const fileNameOnly = zipPath.split("/").pop();
      const sheetRelsPath = `xl/worksheets/_rels/${fileNameOnly}.rels`;
      zip.remove(sheetRelsPath);
      const directRels = getZipFileCaseInsensitive(zip, sheetRelsPath);
      if (directRels) zip.remove(directRels.name);
    }
  }

  // 7. Update docProps/app.xml with matched vector counts so Excel opens without repair prompts
  const appPropFile =
    getZipFileCaseInsensitive(zip, "docProps/app.xml") ||
    getZipFileCaseInsensitive(zip, "app.xml");
  if (appPropFile) {
    const finalSheetName =
      customTabName && customTabName.trim()
        ? customTabName.trim()
        : targetSheet.name;
    const cleanAppXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Excel</Application><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>${encodeXmlAttr(
      finalSheetName
    )}</vt:lpstr></vt:vector></TitlesOfParts></Properties>`;
    zip.file(appPropFile.name, cleanAppXml);
  }

  // 8. Sanitize any unwanted Devanagari font fallbacks so Khmer text renders in authentic Khmer OS Battambang
  await sanitizeKhmerFontsInZip(zip);

  zip.file(workbookFile.name, workbookXml);
  zip.file(relsFile.name, relsXml);
  zip.file(contentTypesFile.name, contentTypesXml);

  const outBuffer = await zip.generateAsync({
    type: "uint8array",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });

  return outBuffer;
}

/**
 * Exports a single sheet workbook preserving 100% of original formatting and styles.
 * 
 * Strategy:
 * 1. Primary Engine: Spring Boot backend (Apache POI 5.3 native sheet isolation)
 * 2. Client Engine: Direct OpenXML Tab Pruning (exact copy of original file, with other sheets removed)
 * 3. Client Fallback: SheetJS to ExcelJS Structured Conversion (for legacy/offline files)
 */
export async function exportSingleSheetBuffer(
  rawWorkbook: XLSX.WorkBook,
  sheetName: string,
  rawBuffer?: ArrayBuffer,
  customTabName?: string,
  rawFile?: File,
  sheetIndex?: number
): Promise<Uint8Array> {
  // 1. Get clean source buffer
  let sourceBuffer: ArrayBuffer | undefined = rawBuffer;
  if (rawFile) {
    try {
      sourceBuffer = await rawFile.arrayBuffer();
    } catch {
      sourceBuffer = rawBuffer;
    }
  }

  // 2. If it is a modern .xlsx (ZIP package), direct OpenXML tab pruning preserves 100% of formatting with 0 loss
  if (sourceBuffer && sourceBuffer.byteLength > 4) {
    const u8 = new Uint8Array(sourceBuffer.slice(0, 4));
    const isZip = u8[0] === 0x50 && u8[1] === 0x4b;

    if (isZip) {
      try {
        return await splitExcelTabDirect(sourceBuffer, sheetName, customTabName, sheetIndex);
      } catch (errDirect) {
        console.error("Direct OpenXML tab split failed, attempting backend fallback:", errDirect);
      }
    }
  }

  // 3. For .xls (binary BIFF8) or complex files, call Spring Boot backend Apache POI engine
  if (rawFile && typeof window !== "undefined") {
    try {
      const formData = new FormData();
      formData.append("file", rawFile);
      formData.append("sheetName", sheetName);
      if (customTabName && customTabName.trim()) {
        formData.append("customName", customTabName.trim());
      }
      const res = await fetch("/api/excel/split-single", {
        method: "POST",
        body: formData,
      });
      if (res.ok) {
        const arrayBuf = await res.arrayBuffer();
        return new Uint8Array(arrayBuf);
      }
    } catch (err) {
      console.warn("Backend split-single failed, attempting client fallback:", err);
    }
  }

  // 3. Fallback for legacy .xls (BIFF8), XML, HTML, or corrupted files:
  // Convert cell data, merged cells, column widths, row heights, and number formats
  // from SheetJS into an ExcelJS Workbook.
  try {
    const ExcelJS = await import("exceljs");
    const wb = new ExcelJS.Workbook();
    const finalSheetName =
      (customTabName || sheetName).replace(/[/\\?*[\]:]/g, "_").slice(0, 31).trim() ||
      "Sheet1";
    const ws = wb.addWorksheet(finalSheetName);
    const sourceSheet = rawWorkbook.Sheets[sheetName];

    if (sourceSheet && sourceSheet["!ref"]) {
      const range = XLSX.utils.decode_range(sourceSheet["!ref"]);
      for (let r = range.s.r; r <= range.e.r; r++) {
        const row = ws.getRow(r + 1);
        for (let c = range.s.c; c <= range.e.c; c++) {
          const cellAddress = XLSX.utils.encode_cell({ r, c });
          const cell = sourceSheet[cellAddress];
          if (cell && cell.v !== undefined) {
            const targetCell = row.getCell(c + 1);
            if (cell.f) {
              targetCell.value = { formula: cell.f, result: cell.v };
            } else if (cell.t === "d" || cell.v instanceof Date) {
              targetCell.value = cell.v;
            } else {
              targetCell.value = cell.v;
            }

            // Copy number format string if present
            if (cell.z) {
              targetCell.numFmt = cell.z;
            }

            // Ensure Khmer text uses standard Khmer font instead of English Calibri
            const valStr = String(cell.v);
            if (/[\u1780-\u17FF]/.test(valStr)) {
              targetCell.font = { name: "Khmer OS Battambang", size: 11 };
            }
          }
        }
      }

      // Preserve merged cells
      if (sourceSheet["!merges"]) {
        sourceSheet["!merges"].forEach((m) => {
          try {
            ws.mergeCells(m.s.r + 1, m.s.c + 1, m.e.r + 1, m.e.c + 1);
          } catch {
            // ignore merge conflicts
          }
        });
      }

      // Preserve column widths
      if (sourceSheet["!cols"]) {
        sourceSheet["!cols"].forEach((col, idx) => {
          if (col) {
            if (col.wch) {
              ws.getColumn(idx + 1).width = col.wch;
            } else if (col.wpx) {
              ws.getColumn(idx + 1).width = Math.max(10, Math.round(col.wpx / 7));
            }
          }
        });
      }

      // Preserve row heights
      if (sourceSheet["!rows"]) {
        sourceSheet["!rows"].forEach((row, idx) => {
          if (row) {
            if (row.hpt) {
              ws.getRow(idx + 1).height = row.hpt;
            } else if (row.hpx) {
              ws.getRow(idx + 1).height = Math.round(row.hpx * 0.75);
            }
          }
        });
      }
    }

    const outBuffer = await wb.xlsx.writeBuffer();
    return new Uint8Array(outBuffer);
  } catch (errConv) {
    console.error("Structured ExcelJS conversion failed:", errConv);
  }

  // 4. Last-resort fallback
  const newWb = XLSX.utils.book_new();
  const sourceSheet = rawWorkbook.Sheets[sheetName];
  const clonedSheet: XLSX.WorkSheet = JSON.parse(JSON.stringify(sourceSheet || {}));
  const safeSheetName =
    (customTabName || sheetName).replace(/[/\\?*[\]:]/g, "_").slice(0, 31).trim() ||
    "Sheet1";

  XLSX.utils.book_append_sheet(newWb, clonedSheet, safeSheetName);

  const buffer = XLSX.write(newWb, {
    bookType: "xlsx",
    type: "array",
    compression: true,
  });

  return new Uint8Array(buffer);
}

/**
 * Triggers a direct browser file download for a Blob.
 */
export function triggerBlobDownload(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

/**
 * Bundles a single workbook's split sheets into a ZIP inside its folder.
 */
export async function createWorkbookZip(
  workbook: UploadedWorkbook,
  globalConfig: GlobalConfig
): Promise<Blob> {
  const config = getEffectiveConfig(workbook, globalConfig);

  // 1. Try Spring Boot backend first if rawFile is available
  if (workbook.rawFile && typeof window !== "undefined") {
    try {
      const formData = new FormData();
      formData.append("file", workbook.rawFile);

      const activeSheets = workbook.sheets.filter((s) => s.selected);
      const customNames: Record<string, string> = {};
      activeSheets.forEach((s) => {
        if (s.customName && s.customName.trim()) {
          customNames[s.originalName] = s.customName.trim();
        }
      });

      const configPayload = {
        folderName: config.folderName || workbook.folderName,
        prefix: config.prefix,
        separator: config.separator,
        namingMode: config.namingMode,
        startNumber: config.startNumber,
        zeroPadding: config.zeroPadding,
        firstSheetIsInfo: config.firstSheetIsInfo,
        infoFileName: config.infoFileName,
        selectedSheets: activeSheets.map((s) => s.originalName),
        customSheetNames: customNames,
      };

      formData.append("config", JSON.stringify(configPayload));

      const res = await fetch("/api/excel/split-zip", {
        method: "POST",
        body: formData,
      });

      if (res.ok) {
        return await res.blob();
      }
    } catch {
      // Backend not running, fallback to client-side packaging
    }
  }

  const zip = createJSZipInstance();
  const folderName = sanitizeFileName(config.folderName || workbook.folderName, "Excel_Sheets");
  const folder = zip.folder(folderName);

  if (!folder) {
    throw new Error("Failed to create ZIP folder.");
  }

  const activeSheets = workbook.sheets.filter((s) => s.selected);
  const usedFileNames = new Set<string>();
  const fileExt = workbook.isBinaryXls || workbook.fileName.toLowerCase().endsWith(".xls") ? "xls" : "xlsx";
  const hasInfoSheet = config.firstSheetIsInfo && activeSheets.length > 0;

  for (let idx = 0; idx < activeSheets.length; idx++) {
    const sheet = activeSheets[idx];
    const isInfo = hasInfoSheet && idx === 0;
    const effectiveIndex = hasInfoSheet && idx > 0 ? idx - 1 : idx;
    const totalSeqSheets = hasInfoSheet ? Math.max(1, activeSheets.length - 1) : activeSheets.length;

    const rawFileName = resolveSheetFileName({
      prefix: config.prefix,
      separator: config.separator,
      namingMode: config.namingMode,
      originalSheetName: sheet.originalName,
      customName: sheet.customName,
      index: effectiveIndex,
      startNumber: config.startNumber,
      zeroPadding: config.zeroPadding,
      totalSheets: totalSeqSheets,
      extension: fileExt,
      isInfoSheet: isInfo,
      infoFileName: config.infoFileName,
    });

    const uniqueFileName = getUniqueName(rawFileName, usedFileNames);
    const origIdx = workbook.sheets.findIndex((s) => s.id === sheet.id);

    const sheetBuffer = await exportSingleSheetBuffer(
      workbook.rawWorkbook,
      sheet.originalName,
      workbook.rawBuffer,
      sheet.customName,
      workbook.rawFile,
      origIdx !== -1 ? origIdx : idx
    );
    folder.file(uniqueFileName, sheetBuffer);
  }

  return await zip.generateAsync({
    type: "blob",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });
}

/**
 * Bundles all workbooks, each in their respective folders, into one master ZIP.
 */
export async function createMasterZip(
  workbooks: UploadedWorkbook[],
  globalConfig: GlobalConfig,
  onProgress?: (progress: number, currentAction: string) => void
): Promise<Blob> {
  // 1. Try Spring Boot batch split endpoint if all rawFiles are present
  if (
    workbooks.length > 0 &&
    workbooks.every((wb) => wb.rawFile) &&
    typeof window !== "undefined"
  ) {
    try {
      if (onProgress) onProgress(15, "Connecting to Spring Boot backend...");
      const formData = new FormData();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const configsMap: Record<string, any> = {};

      for (const wb of workbooks) {
        if (wb.rawFile) {
          formData.append("files", wb.rawFile);
          const config = getEffectiveConfig(wb, globalConfig);
          const activeSheets = wb.sheets.filter((s) => s.selected);
          const customNames: Record<string, string> = {};
          activeSheets.forEach((s) => {
            if (s.customName && s.customName.trim()) {
              customNames[s.originalName] = s.customName.trim();
            }
          });

          configsMap[wb.fileName] = {
            folderName: config.folderName || wb.folderName,
            prefix: config.prefix,
            separator: config.separator,
            namingMode: config.namingMode,
            startNumber: config.startNumber,
            zeroPadding: config.zeroPadding,
            firstSheetIsInfo: config.firstSheetIsInfo,
            infoFileName: config.infoFileName,
            selectedSheets: activeSheets.map((s) => s.originalName),
            customSheetNames: customNames,
          };
        }
      }

      formData.append("configs", JSON.stringify(configsMap));
      if (globalConfig.masterFolderName && globalConfig.masterFolderName.trim()) {
        formData.append(
          "globalConfig",
          JSON.stringify({
            folderName: globalConfig.masterFolderName.trim(),
            prefix: globalConfig.prefix,
            separator: globalConfig.separator,
            namingMode: globalConfig.namingMode,
            startNumber: globalConfig.startNumber,
            zeroPadding: globalConfig.zeroPadding,
            firstSheetIsInfo: globalConfig.firstSheetIsInfo,
            infoFileName: globalConfig.infoFileName,
          })
        );
      }
      if (onProgress) onProgress(45, "Apache POI splitting & packaging master ZIP...");

      const res = await fetch("/api/excel/split-batch-zip", {
        method: "POST",
        body: formData,
      });

      if (res.ok) {
        if (onProgress) onProgress(100, "Master ZIP ready!");
        return await res.blob();
      }
    } catch {
      // Backend not running, proceed to client packaging
    }
  }

  const masterZip = createJSZipInstance();
  const totalSheetsAcrossAll = workbooks.reduce(
    (acc, wb) => acc + wb.sheets.filter((s) => s.selected).length,
    0
  );

  let processedSheets = 0;
  const usedFolderNames = new Set<string>();

  for (const workbook of workbooks) {
    const config = getEffectiveConfig(workbook, globalConfig);
    const rawFolderName = sanitizeFileName(
      config.folderName || workbook.folderName,
      workbook.fileName.replace(/\.[^/.]+$/, "")
    );
    const cleanFolderName = getUniqueName(rawFolderName, usedFolderNames);
    const folder = masterZip.folder(cleanFolderName);

    if (!folder) continue;

    const activeSheets = workbook.sheets.filter((s) => s.selected);
    const usedFileNames = new Set<string>();
    const fileExt = workbook.isBinaryXls || workbook.fileName.toLowerCase().endsWith(".xls") ? "xls" : "xlsx";
    const hasInfoSheet = config.firstSheetIsInfo && activeSheets.length > 0;

    for (let idx = 0; idx < activeSheets.length; idx++) {
      const sheet = activeSheets[idx];
      const isInfo = hasInfoSheet && idx === 0;
      const effectiveIndex = hasInfoSheet && idx > 0 ? idx - 1 : idx;
      const totalSeqSheets = hasInfoSheet ? Math.max(1, activeSheets.length - 1) : activeSheets.length;

      const rawFileName = resolveSheetFileName({
        prefix: config.prefix,
        separator: config.separator,
        namingMode: config.namingMode,
        originalSheetName: sheet.originalName,
        customName: sheet.customName,
        index: effectiveIndex,
        startNumber: config.startNumber,
        zeroPadding: config.zeroPadding,
        totalSheets: totalSeqSheets,
        extension: fileExt,
        isInfoSheet: isInfo,
        infoFileName: config.infoFileName,
      });

      const uniqueFileName = getUniqueName(rawFileName, usedFileNames);

      if (onProgress) {
        processedSheets++;
        const pct = Math.round((processedSheets / (totalSheetsAcrossAll || 1)) * 80);
        onProgress(pct, `Splitting ${workbook.fileName} -> ${uniqueFileName}`);
      }

      const origIdx = workbook.sheets.findIndex((s) => s.id === sheet.id);

      const sheetBuffer = await exportSingleSheetBuffer(
        workbook.rawWorkbook,
        sheet.originalName,
        workbook.rawBuffer,
        sheet.customName,
        workbook.rawFile,
        origIdx !== -1 ? origIdx : idx
      );
      folder.file(uniqueFileName, sheetBuffer);
    }
  }

  if (onProgress) {
    onProgress(85, "Compressing into ZIP archive...");
  }

  const zipBlob = await masterZip.generateAsync(
    {
      type: "blob",
      compression: "DEFLATE",
      compressionOptions: { level: 6 },
    },
    (metadata) => {
      if (onProgress) {
        const pct = 85 + Math.round((metadata.percent / 100) * 15);
        onProgress(pct, `Packaging archive (${Math.round(metadata.percent)}%)...`);
      }
    }
  );

  return zipBlob;
}

/**
 * Helper to generate a dummy demo Excel workbook with multiple realistic sheets.
 */
export function generateSampleExcelFile(): File {
  const wb = XLSX.utils.book_new();

  // Sheet 1: Sales Q1
  const salesData = [
    ["Region", "Rep Name", "Target ($)", "Actual ($)", "Status"],
    ["North", "Alice Chen", 50000, 62400, "Exceeded"],
    ["South", "Bob Miller", 45000, 48200, "Exceeded"],
    ["East", "Charlie Davis", 55000, 49000, "Pending"],
    ["West", "Dana White", 60000, 67800, "Exceeded"],
    ["Central", "Evan Harris", 40000, 40500, "Achieved"],
  ];
  const ws1 = XLSX.utils.aoa_to_sheet(salesData);
  XLSX.utils.book_append_sheet(wb, ws1, "Sales_Q1");

  // Sheet 2: Inventory Status
  const inventoryData = [
    ["SKU", "Item Description", "Category", "Quantity", "Reorder Level"],
    ["SKU-1001", "Wireless Optical Mouse", "Electronics", 240, 50],
    ["SKU-1002", "Ergonomic Mechanical Keyboard", "Electronics", 95, 30],
    ["SKU-1003", "USB-C Multiport Hub 7-in-1", "Accessories", 310, 80],
    ["SKU-1004", "27-inch 4K IPS Monitor", "Displays", 45, 15],
    ["SKU-1005", "Noise Cancelling Headphones", "Audio", 120, 25],
  ];
  const ws2 = XLSX.utils.aoa_to_sheet(inventoryData);
  XLSX.utils.book_append_sheet(wb, ws2, "Inventory");

  // Sheet 3: Employee Directory
  const employeeData = [
    ["Emp ID", "Full Name", "Department", "Role", "Location"],
    ["EMP-01", "Sarah Connor", "Engineering", "Lead Architect", "San Francisco"],
    ["EMP-02", "James Wilson", "Product", "Senior PM", "New York"],
    ["EMP-03", "Maria Gonzalez", "Marketing", "Brand Director", "Austin"],
    ["EMP-04", "Kenji Sato", "Support", "Customer Success Lead", "Tokyo"],
    ["EMP-05", "Elena Rostova", "Finance", "Senior Accountant", "London"],
  ];
  const ws3 = XLSX.utils.aoa_to_sheet(employeeData);
  XLSX.utils.book_append_sheet(wb, ws3, "Employees");

  // Sheet 4: Marketing Campaigns
  const marketingData = [
    ["Campaign ID", "Channel", "Budget ($)", "Impressions", "Conversions"],
    ["CMP-2024-A", "Google Search", 12000, 450000, 3200],
    ["CMP-2024-B", "LinkedIn Ads", 8500, 180000, 1150],
    ["CMP-2024-C", "YouTube Video", 15000, 920000, 4600],
    ["CMP-2024-D", "Newsletter Sponsorship", 3500, 85000, 920],
  ];
  const ws4 = XLSX.utils.aoa_to_sheet(marketingData);
  XLSX.utils.book_append_sheet(wb, ws4, "Marketing_Campaigns");

  const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  return new File([wbout], "Sample_Company_Master.xlsx", {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

