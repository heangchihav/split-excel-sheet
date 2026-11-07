"use client";

import React, { useState } from "react";
import {
  FileSpreadsheet,
  ChevronDown,
  ChevronUp,
  Download,
  Trash2,
  Eye,
  Sliders,
  CheckSquare,
  Square,
  RefreshCw,
  FolderOpen,
  ArrowRight,
  Edit3,
  X,
  CheckCircle2,
} from "lucide-react";
import {
  GlobalConfig,
  UploadedWorkbook,
  SheetItem,
  NamingMode,
  ZeroPadding,
  Separator,
} from "@/types/excel";
import {
  resolveSheetFileName,
  getEffectiveConfig,
  sanitizeFileName,
  exportSingleSheetBuffer,
  triggerBlobDownload,
  createWorkbookZip,
} from "@/lib/excel-utils";

interface WorkbookItemCardProps {
  workbook: UploadedWorkbook;
  globalConfig: GlobalConfig;
  onUpdateWorkbook: (updated: UploadedWorkbook) => void;
  onRemoveWorkbook: (id: string) => void;
  onPreviewSheet: (sheet: SheetItem, masterName: string) => void;
}

export function WorkbookItemCard({
  workbook,
  globalConfig,
  onUpdateWorkbook,
  onRemoveWorkbook,
  onPreviewSheet,
}: WorkbookItemCardProps) {
  const [isExpanded, setIsExpanded] = useState(true);
  const [isDownloadingFolder, setIsDownloadingFolder] = useState(false);
  const [editingSheetId, setEditingSheetId] = useState<string | null>(null);

  const effectiveConfig = getEffectiveConfig(workbook, globalConfig);

  const defaultMasterFolderName = sanitizeFileName(
    workbook.fileName.replace(/\.[^/.]+$/, ""),
    "Excel_Folder"
  );

  const handleFolderNameChange = (newVal: string) => {
    onUpdateWorkbook({
      ...workbook,
      folderName: newVal,
    });
  };

  const handleResetFolderName = () => {
    onUpdateWorkbook({
      ...workbook,
      folderName: defaultMasterFolderName,
    });
  };

  const toggleSelectAll = (select: boolean) => {
    onUpdateWorkbook({
      ...workbook,
      sheets: workbook.sheets.map((s) => ({ ...s, selected: select })),
    });
  };

  const toggleSheetSelected = (sheetId: string) => {
    onUpdateWorkbook({
      ...workbook,
      sheets: workbook.sheets.map((s) =>
        s.id === sheetId ? { ...s, selected: !s.selected } : s
      ),
    });
  };

  const updateSheetCustomName = (sheetId: string, customName: string) => {
    onUpdateWorkbook({
      ...workbook,
      sheets: workbook.sheets.map((s) =>
        s.id === sheetId ? { ...s, customName } : s
      ),
    });
  };

  const toggleUseCustomConfig = (useCustom: boolean) => {
    onUpdateWorkbook({
      ...workbook,
      useCustomConfig: useCustom,
      customPrefix: useCustom ? workbook.customPrefix ?? globalConfig.prefix : undefined,
      customNamingMode: useCustom
        ? workbook.customNamingMode ?? globalConfig.namingMode
        : undefined,
      customStartNumber: useCustom
        ? workbook.customStartNumber ?? globalConfig.startNumber
        : undefined,
      customZeroPadding: useCustom
        ? workbook.customZeroPadding ?? globalConfig.zeroPadding
        : undefined,
      customSeparator: useCustom
        ? workbook.customSeparator ?? globalConfig.separator
        : undefined,
      customFirstSheetIsInfo: useCustom
        ? workbook.customFirstSheetIsInfo ?? globalConfig.firstSheetIsInfo
        : undefined,
      customInfoFileName: useCustom
        ? workbook.customInfoFileName ?? globalConfig.infoFileName
        : undefined,
    });
  };

  const updateCustomSetting = <K extends keyof UploadedWorkbook>(
    key: K,
    val: UploadedWorkbook[K]
  ) => {
    onUpdateWorkbook({
      ...workbook,
      [key]: val,
    });
  };

  // Download this master workbook's split sheets as a folder in a ZIP
  const handleDownloadFolderZip = async () => {
    try {
      setIsDownloadingFolder(true);
      const zipBlob = await createWorkbookZip(workbook, globalConfig);
      const cleanFolder = sanitizeFileName(effectiveConfig.folderName || workbook.folderName, defaultMasterFolderName);
      triggerBlobDownload(zipBlob, `${cleanFolder}.zip`);
    } catch (err) {
      console.error("Error creating folder zip:", err);
      alert("Failed to create zip for this workbook.");
    } finally {
      setIsDownloadingFolder(false);
    }
  };

  // Download a single sheet as .xlsx or .xls
  const handleDownloadSingleSheet = async (sheet: SheetItem, index: number) => {
    try {
      const activeSheets = workbook.sheets.filter((s) => s.selected);
      const isXls = workbook.isBinaryXls || workbook.fileName.toLowerCase().endsWith(".xls");
      const fileExt = isXls ? "xls" : "xlsx";
      const isInfo = effectiveConfig.firstSheetIsInfo && index === 0;
      const effectiveIndex = effectiveConfig.firstSheetIsInfo && index > 0 ? index - 1 : index;
      const totalSeqSheets = effectiveConfig.firstSheetIsInfo
        ? Math.max(1, activeSheets.length - 1)
        : activeSheets.length;

      const fileName = resolveSheetFileName({
        prefix: effectiveConfig.prefix,
        separator: effectiveConfig.separator,
        namingMode: effectiveConfig.namingMode,
        originalSheetName: sheet.originalName,
        customName: sheet.customName,
        index: effectiveIndex,
        startNumber: effectiveConfig.startNumber,
        zeroPadding: effectiveConfig.zeroPadding,
        totalSheets: totalSeqSheets,
        extension: fileExt,
        isInfoSheet: isInfo,
        infoFileName: effectiveConfig.infoFileName,
      });

      const origIdx = workbook.sheets.findIndex((s) => s.id === sheet.id);

      const buffer = await exportSingleSheetBuffer(
        workbook.rawWorkbook,
        sheet.originalName,
        workbook.rawBuffer,
        sheet.customName,
        workbook.rawFile,
        origIdx !== -1 ? origIdx : index
      );
      const mimeType = isXls
        ? "application/vnd.ms-excel"
        : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
      const blob = new Blob([buffer as BlobPart], {
        type: mimeType,
      });
      triggerBlobDownload(blob, fileName);
    } catch (err) {
      console.error("Error exporting single sheet:", err);
      alert("Failed to export sheet.");
    }
  };

  const selectedCount = workbook.sheets.filter((s) => s.selected).length;
  const totalSheets = workbook.sheets.length;

  return (
    <div className="bg-white rounded-2xl border border-slate-200/90 shadow-xs overflow-hidden mb-5 transition-all">
      {/* Top Banner: File Info & Controls */}
      <div className="p-4 sm:p-5 bg-slate-50/70 border-b border-slate-200/80 flex flex-wrap items-center justify-between gap-4">
        <div className="flex items-center gap-3.5 min-w-0">
          <div className="w-11 h-11 rounded-xl bg-emerald-600/10 border border-emerald-500/20 text-emerald-700 flex items-center justify-center shrink-0">
            <FileSpreadsheet className="w-6 h-6" />
          </div>
          <div className="min-w-0">
            <div className="flex items-center gap-2">
              <h3 className="text-base font-bold text-slate-900 truncate">
                {workbook.fileName}
              </h3>
              <span className="px-2 py-0.5 text-xs font-semibold rounded-md bg-slate-200 text-slate-700 shrink-0">
                {(workbook.fileSize / 1024).toFixed(1)} KB
              </span>
              <span className="px-2 py-0.5 text-xs font-semibold rounded-md bg-emerald-100 text-emerald-800 shrink-0">
                {workbook.isBinaryXls || workbook.fileName.toLowerCase().endsWith(".xls")
                  ? ".xls (Excel 97–2003)"
                  : ".xlsx"}
              </span>
            </div>
            <div className="flex items-center gap-2 text-xs text-slate-500 mt-0.5">
              <span>{totalSheets} sheets found</span>
              <span>•</span>
              <span className="text-emerald-700 font-medium">
                {selectedCount} selected for split
              </span>
            </div>
          </div>
        </div>

        {/* Action buttons */}
        <div className="flex items-center gap-2">
          <button
            type="button"
            onClick={handleDownloadFolderZip}
            disabled={selectedCount === 0 || isDownloadingFolder}
            className="px-3.5 py-2 bg-emerald-50 hover:bg-emerald-100/80 border border-emerald-200 text-emerald-800 rounded-xl text-xs font-bold transition-colors flex items-center gap-1.5 disabled:opacity-50 disabled:cursor-not-allowed shadow-xs"
            title="Download all sheets of this master file in its own ZIP folder"
          >
            <Download className="w-4 h-4 text-emerald-700" />
            <span>
              {isDownloadingFolder ? "Compressing..." : "Download Folder (ZIP)"}
            </span>
          </button>

          <button
            type="button"
            onClick={() => onRemoveWorkbook(workbook.id)}
            className="p-2 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded-xl transition-colors"
            title="Remove this Excel file"
          >
            <Trash2 className="w-4 h-4" />
          </button>

          <button
            type="button"
            onClick={() => setIsExpanded(!isExpanded)}
            className="p-2 text-slate-400 hover:text-slate-700 hover:bg-slate-200/60 rounded-xl transition-colors"
            title={isExpanded ? "Collapse" : "Expand"}
          >
            {isExpanded ? (
              <ChevronUp className="w-4 h-4" />
            ) : (
              <ChevronDown className="w-4 h-4" />
            )}
          </button>
        </div>
      </div>

      {/* Main Body */}
      {isExpanded && (
        <div className="p-4 sm:p-6 space-y-6">
          {/* Format Status */}
          <div className="px-3.5 py-2.5 rounded-lg bg-emerald-50/70 border border-emerald-200/80 flex items-center gap-2 text-xs text-emerald-800">
            <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
            <span>
              <strong>Ready to split:</strong> Original sheet content, fonts (including Khmer Unicode), and formatting will be preserved with 0 loss.
            </span>
          </div>

          {/* Master Folder Name Customization Box */}
          <div className="p-4 rounded-xl bg-amber-50/60 border border-amber-200/80 flex flex-col md:flex-row md:items-center justify-between gap-3">
            <div className="flex items-center gap-3">
              <div className="p-2 rounded-lg bg-amber-100 text-amber-800">
                <FolderOpen className="w-5 h-5" />
              </div>
              <div>
                <div className="flex items-center gap-2">
                  <span className="text-xs font-bold uppercase tracking-wider text-amber-900">
                    Master Folder Name:
                  </span>
                  <span className="text-[11px] text-amber-700/80">
                    (Separated sheets will be saved inside this folder)
                  </span>
                </div>
                <div className="flex items-center gap-2 mt-1">
                  <input
                    type="text"
                    value={workbook.folderName}
                    onChange={(e) => handleFolderNameChange(e.target.value)}
                    className="px-3 py-1.5 bg-white border border-amber-300 rounded-lg text-sm font-semibold text-slate-900 focus:outline-none focus:ring-2 focus:ring-amber-500/20 focus:border-amber-500 min-w-[260px] max-w-sm"
                    placeholder={globalConfig.masterFolderName?.trim() || defaultMasterFolderName}
                  />
                  {workbook.folderName !== defaultMasterFolderName && (
                    <button
                      type="button"
                      onClick={handleResetFolderName}
                      className="px-2.5 py-1.5 text-xs text-amber-800 hover:text-amber-950 hover:bg-amber-100 rounded-lg font-medium transition-colors flex items-center gap-1"
                      title="Reset to default master file name"
                    >
                      <RefreshCw className="w-3.5 h-3.5" />
                      Reset to Master Name
                    </button>
                  )}
                </div>
              </div>
            </div>

            <div className="text-right">
              <span className="text-xs text-slate-500 font-mono block">
                Target Folder Path:
              </span>
              <span className="text-xs font-mono font-bold text-slate-800 bg-white/80 px-2 py-1 rounded border border-amber-200 inline-block mt-0.5">
                📁 {sanitizeFileName(effectiveConfig.folderName || workbook.folderName, defaultMasterFolderName)}/
              </span>
            </div>
          </div>

          {/* Optional: Override global settings just for this workbook */}
          <div className="border border-slate-200 rounded-xl overflow-hidden bg-slate-50/40">
            <div
              onClick={() => toggleUseCustomConfig(!workbook.useCustomConfig)}
              className="px-4 py-3 cursor-pointer flex items-center justify-between bg-slate-100/60 hover:bg-slate-100 transition-colors"
            >
              <div className="flex items-center gap-2 text-xs font-bold text-slate-800">
                <Sliders className="w-4 h-4 text-slate-600" />
                <span>Custom Prefix & Naming Rules for This File</span>
                {workbook.useCustomConfig && (
                  <span className="px-2 py-0.5 rounded text-[10px] bg-emerald-100 text-emerald-800 font-semibold uppercase">
                    Custom Active
                  </span>
                )}
              </div>
              <span className="text-xs text-emerald-700 font-medium hover:underline">
                {workbook.useCustomConfig ? "Revert to Global" : "Configure Custom Rules"}
              </span>
            </div>

            {workbook.useCustomConfig && (
              <div className="p-4 bg-white border-t border-slate-200 space-y-4 animate-in fade-in duration-150">
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">
                      File Prefix
                    </label>
                    <input
                      type="text"
                      value={workbook.customPrefix ?? ""}
                      onChange={(e) =>
                        updateCustomSetting("customPrefix", e.target.value)
                      }
                      placeholder="e.g. SalesReport_"
                      className="w-full px-3 py-1.5 bg-slate-50 border border-slate-300 rounded-lg text-xs font-medium focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">
                      Behind Naming
                    </label>
                    <select
                      value={workbook.customNamingMode ?? globalConfig.namingMode}
                      onChange={(e) =>
                        updateCustomSetting(
                          "customNamingMode",
                          e.target.value as NamingMode
                        )
                      }
                      className="w-full px-3 py-1.5 bg-slate-50 border border-slate-300 rounded-lg text-xs font-medium focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500"
                    >
                      <option value="original">Old Sheet Name</option>
                      <option value="sequence">Auto Number (1, 2, 3...)</option>
                      <option value="sequence_original">Number + Sheet Name</option>
                    </select>
                  </div>

                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">
                      Separator
                    </label>
                    <select
                      value={workbook.customSeparator ?? globalConfig.separator}
                      onChange={(e) =>
                        updateCustomSetting(
                          "customSeparator",
                          e.target.value as Separator
                        )
                      }
                      className="w-full px-3 py-1.5 bg-slate-50 border border-slate-300 rounded-lg text-xs font-medium focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500"
                    >
                      <option value="_">Underscore (_)</option>
                      <option value="-">Hyphen (-)</option>
                      <option value=" ">Space ( )</option>
                      <option value="">None</option>
                    </select>
                  </div>
                </div>

                {workbook.customNamingMode !== "original" && (
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-2">
                    <div>
                      <label className="block text-xs font-semibold text-slate-700 mb-1">
                        Start Number
                      </label>
                      <input
                        type="number"
                        min={0}
                        value={workbook.customStartNumber ?? 1}
                        onChange={(e) =>
                          updateCustomSetting(
                            "customStartNumber",
                            parseInt(e.target.value) || 0
                          )
                        }
                        className="w-full px-3 py-1.5 bg-slate-50 border border-slate-300 rounded-lg text-xs font-medium"
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-semibold text-slate-700 mb-1">
                        Padding
                      </label>
                      <select
                        value={workbook.customZeroPadding ?? "2"}
                        onChange={(e) =>
                          updateCustomSetting(
                            "customZeroPadding",
                            e.target.value as ZeroPadding
                          )
                        }
                        className="w-full px-3 py-1.5 bg-slate-50 border border-slate-300 rounded-lg text-xs font-medium"
                      >
                        <option value="none">No Padding (1, 2...)</option>
                        <option value="2">2 Digits (01, 02...)</option>
                        <option value="3">3 Digits (001, 002...)</option>
                        <option value="auto">Auto</option>
                      </select>
                    </div>
                  </div>
                )}

                <div className="pt-2 border-t border-slate-100 space-y-2">
                  <label className="flex items-center gap-2 cursor-pointer text-xs font-semibold text-slate-700">
                    <input
                      type="checkbox"
                      checked={
                        workbook.customFirstSheetIsInfo ??
                        globalConfig.firstSheetIsInfo
                      }
                      onChange={(e) =>
                        updateCustomSetting(
                          "customFirstSheetIsInfo",
                          e.target.checked
                        )
                      }
                      className="w-4 h-4 text-emerald-600 rounded border-slate-300 focus:ring-emerald-500 cursor-pointer"
                    />
                    <span>
                      First sheet is Info / Menu (Behind name format starts from sheet #2)
                    </span>
                  </label>

                  {(workbook.customFirstSheetIsInfo ?? globalConfig.firstSheetIsInfo) && (
                    <div className="ml-6 space-y-1">
                      <label className="block text-[11px] font-semibold text-slate-600">
                        Custom Info File Name (optional)
                      </label>
                      <div className="flex items-center gap-2 max-w-sm">
                        <input
                          type="text"
                          value={workbook.customInfoFileName ?? ""}
                          onChange={(e) =>
                            updateCustomSetting("customInfoFileName", e.target.value)
                          }
                          placeholder={
                            effectiveConfig.prefix
                              ? `${effectiveConfig.prefix} (Default: Prefix)`
                              : "Info"
                          }
                          className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg text-xs font-medium focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500"
                        />
                        <span className="text-[11px] font-mono text-slate-500 shrink-0">
                          .{workbook.isBinaryXls || workbook.fileName.toLowerCase().endsWith(".xls") ? "xls" : "xlsx"}
                        </span>
                      </div>
                    </div>
                  )}
                </div>
              </div>
            )}
          </div>

          {/* Sheet List & Selection */}
          <div>
            <div className="flex flex-wrap items-center justify-between gap-2 mb-3">
              <div className="flex items-center gap-2.5 flex-wrap">
                <span className="text-xs font-bold uppercase tracking-wider text-slate-600">
                  Sheets to Split ({selectedCount} / {totalSheets})
                </span>
                <button
                  type="button"
                  onClick={() => {
                    const nextVal = !effectiveConfig.firstSheetIsInfo;
                    onUpdateWorkbook({
                      ...workbook,
                      useCustomConfig: true,
                      customFirstSheetIsInfo: nextVal,
                    });
                  }}
                  className={`text-xs px-2.5 py-1 rounded-lg border font-semibold transition-all flex items-center gap-1.5 ${
                    effectiveConfig.firstSheetIsInfo
                      ? "bg-amber-100/90 border-amber-300 text-amber-950 shadow-2xs font-bold"
                      : "bg-slate-100 hover:bg-slate-200/80 border-slate-200 text-slate-600"
                  }`}
                  title="Toggle whether the 1st sheet is treated as an Info/Menu sheet"
                >
                  <span>📌 1st Sheet is Info / Menu</span>
                  {effectiveConfig.firstSheetIsInfo && (
                    <span className="w-1.5 h-1.5 rounded-full bg-amber-600" />
                  )}
                </button>
              </div>
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => toggleSelectAll(true)}
                  className="text-xs font-medium text-emerald-700 hover:text-emerald-800 px-2 py-1 rounded hover:bg-emerald-50 transition-colors"
                >
                  Select All
                </button>
                <span className="text-slate-300">|</span>
                <button
                  type="button"
                  onClick={() => toggleSelectAll(false)}
                  className="text-xs font-medium text-slate-500 hover:text-slate-700 px-2 py-1 rounded hover:bg-slate-100 transition-colors"
                >
                  Deselect All
                </button>
              </div>
            </div>

            <div className="border border-slate-200 rounded-xl overflow-hidden shadow-2xs divide-y divide-slate-100">
              {workbook.sheets.map((sheet, idx) => {
                const isXls = workbook.isBinaryXls || workbook.fileName.toLowerCase().endsWith(".xls");
                const fileExt = isXls ? "xls" : "xlsx";
                const isInfo = effectiveConfig.firstSheetIsInfo && idx === 0;
                const effectiveIndex = effectiveConfig.firstSheetIsInfo && idx > 0 ? idx - 1 : idx;
                const totalSeqSheets = effectiveConfig.firstSheetIsInfo
                  ? Math.max(1, workbook.sheets.filter((s) => s.selected).length - 1)
                  : workbook.sheets.filter((s) => s.selected).length;

                const targetFileName = resolveSheetFileName({
                  prefix: effectiveConfig.prefix,
                  separator: effectiveConfig.separator,
                  namingMode: effectiveConfig.namingMode,
                  originalSheetName: sheet.originalName,
                  customName: sheet.customName,
                  index: effectiveIndex,
                  startNumber: effectiveConfig.startNumber,
                  zeroPadding: effectiveConfig.zeroPadding,
                  totalSheets: totalSeqSheets,
                  extension: fileExt,
                  isInfoSheet: isInfo,
                  infoFileName: effectiveConfig.infoFileName,
                });

                return (
                  <div
                    key={sheet.id}
                    className={`p-3 sm:px-4 sm:py-3.5 flex flex-wrap items-center justify-between gap-3 transition-colors ${
                      sheet.selected
                        ? "bg-white hover:bg-slate-50/70"
                        : "bg-slate-50/60 opacity-60 hover:opacity-80"
                    }`}
                  >
                    {/* Checkbox and Sheet Info */}
                    <div className="flex items-center gap-3 min-w-0">
                      <button
                        type="button"
                        onClick={() => toggleSheetSelected(sheet.id)}
                        className="text-emerald-600 hover:text-emerald-700 transition-colors shrink-0"
                      >
                        {sheet.selected ? (
                          <CheckSquare className="w-5 h-5 text-emerald-600" />
                        ) : (
                          <Square className="w-5 h-5 text-slate-400" />
                        )}
                      </button>

                      <div className="min-w-0">
                        <div className="flex items-center gap-2 flex-wrap">
                          <span className="text-xs font-mono font-bold text-slate-400">
                            #{idx + 1}
                          </span>
                          {isInfo && (
                            <span className="text-[10px] font-bold uppercase px-2 py-0.5 rounded-md bg-amber-100 text-amber-900 border border-amber-300 shrink-0 shadow-2xs">
                              Info / Menu
                            </span>
                          )}
                          <span className="text-sm font-semibold text-slate-900 truncate">
                            {sheet.originalName}
                          </span>
                          <span className="text-[11px] px-2 py-0.5 rounded-full bg-slate-100 text-slate-600 font-medium">
                            {sheet.rowCount} rows × {sheet.colCount} cols
                          </span>
                        </div>
                      </div>
                    </div>

                    {/* Output File Name Preview & Actions */}
                    <div className="flex items-center gap-2.5 sm:ml-auto">
                      {editingSheetId === sheet.id ? (
                        <div className="flex items-center gap-1.5 bg-slate-100 p-1 rounded-lg">
                          <input
                            type="text"
                            value={sheet.customName ?? ""}
                            onChange={(e) =>
                              updateSheetCustomName(sheet.id, e.target.value)
                            }
                            placeholder={targetFileName}
                            className="px-2 py-1 bg-white border border-slate-300 rounded text-xs font-mono w-44 focus:outline-none focus:ring-1 focus:ring-emerald-500"
                            autoFocus
                          />
                          <button
                            type="button"
                            onClick={() => setEditingSheetId(null)}
                            className="p-1 text-slate-500 hover:text-slate-800 rounded"
                            title="Done"
                          >
                            <X className="w-3.5 h-3.5" />
                          </button>
                        </div>
                      ) : (
                        <div className="flex items-center gap-1.5 text-xs font-mono bg-emerald-50 border border-emerald-200/90 text-emerald-900 px-2.5 py-1 rounded-lg group/name">
                          <ArrowRight className="w-3.5 h-3.5 text-emerald-600 shrink-0" />
                          <span className="font-semibold truncate max-w-[200px] sm:max-w-xs">
                            {targetFileName}
                          </span>
                          <button
                            type="button"
                            onClick={() => setEditingSheetId(sheet.id)}
                            className="ml-1 text-slate-400 hover:text-emerald-700 opacity-60 hover:opacity-100"
                            title="Rename this file manually"
                          >
                            <Edit3 className="w-3 h-3" />
                          </button>
                        </div>
                      )}

                      {/* Preview Sheet Data */}
                      <button
                        type="button"
                        onClick={() => onPreviewSheet(sheet, workbook.fileName)}
                        className="p-1.5 text-slate-500 hover:text-slate-800 hover:bg-slate-100 rounded-lg transition-colors"
                        title="Preview sheet data"
                      >
                        <Eye className="w-4 h-4" />
                      </button>

                      {/* Direct Single Download */}
                      <button
                        type="button"
                        onClick={() => handleDownloadSingleSheet(sheet, idx)}
                        className="p-1.5 text-emerald-700 hover:text-emerald-900 hover:bg-emerald-50 rounded-lg transition-colors"
                        title="Download this single sheet (.xlsx)"
                      >
                        <Download className="w-4 h-4" />
                      </button>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
