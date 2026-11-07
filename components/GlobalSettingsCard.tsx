"use client";

import React from "react";
import {
  Settings2,
  FolderTree,
  FolderOpen,
  FileText,
  Hash,
  Sparkles,
  Info,
} from "lucide-react";
import type { GlobalConfig, Separator, ZeroPadding } from "@/types/excel";
import { resolveSheetFileName } from "@/lib/excel-utils";

interface GlobalSettingsCardProps {
  config: GlobalConfig;
  onChange: (newConfig: GlobalConfig) => void;
  onReset: () => void;
}

export function GlobalSettingsCard({
  config,
  onChange,
  onReset,
}: GlobalSettingsCardProps) {
  const update = <K extends keyof GlobalConfig>(key: K, value: GlobalConfig[K]) => {
    onChange({
      ...config,
      [key]: value,
    });
  };

  // Compute live preview examples
  const hasInfo = config.firstSheetIsInfo;
  const sampleInfo = resolveSheetFileName({
    prefix: config.prefix,
    separator: config.separator,
    namingMode: config.namingMode,
    originalSheetName: "Menu_Info",
    index: 0,
    startNumber: config.startNumber,
    zeroPadding: config.zeroPadding,
    totalSheets: 5,
    isInfoSheet: true,
    infoFileName: config.infoFileName,
  });

  const sample1 = resolveSheetFileName({
    prefix: config.prefix,
    separator: config.separator,
    namingMode: config.namingMode,
    originalSheetName: "January_Sales",
    index: 0,
    startNumber: config.startNumber,
    zeroPadding: config.zeroPadding,
    totalSheets: hasInfo ? 4 : 5,
    isInfoSheet: false,
  });

  const sample2 = resolveSheetFileName({
    prefix: config.prefix,
    separator: config.separator,
    namingMode: config.namingMode,
    originalSheetName: "February_Sales",
    index: 1,
    startNumber: config.startNumber,
    zeroPadding: config.zeroPadding,
    totalSheets: hasInfo ? 4 : 5,
    isInfoSheet: false,
  });

  return (
    <div className="bg-white rounded-2xl border border-slate-200/90 shadow-sm overflow-hidden mb-6 transition-all">
      <div className="px-5 py-4 bg-slate-50/70 border-b border-slate-200/80 flex flex-wrap items-center justify-between gap-3">
        <div className="flex items-center gap-2.5">
          <div className="p-2 rounded-lg bg-emerald-100 text-emerald-800">
            <Settings2 className="w-5 h-5" />
          </div>
          <div>
            <h2 className="text-base font-semibold text-slate-900">
              Split & Naming Configuration
            </h2>
            <p className="text-xs text-slate-500">
              Default rules applied to all split sheets and master folders
            </p>
          </div>
        </div>

        <button
          onClick={onReset}
          className="text-xs font-medium text-slate-500 hover:text-slate-800 hover:bg-slate-200/60 px-3 py-1.5 rounded-lg transition-colors"
        >
          Reset Defaults
        </button>
      </div>

      <div className="p-5 sm:p-6 space-y-6">
        {/* Top grid: Master Folder & Prefix & Separator */}
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-5">
          {/* Master Folder Name */}
          <div className="lg:col-span-5">
            <label className="block text-xs font-bold uppercase tracking-wider text-slate-700 mb-1.5 flex items-center justify-between">
              <span className="flex items-center gap-1.5">
                <FolderOpen className="w-3.5 h-3.5 text-emerald-600" />
                Master Folder Name
              </span>
              <span className="text-[10px] text-emerald-700 font-semibold normal-case bg-emerald-50 px-2 py-0.5 rounded border border-emerald-200">
                Folder inside ZIP & ZIP Name
              </span>
            </label>
            <div className="relative">
              <input
                type="text"
                value={config.masterFolderName ?? ""}
                onChange={(e) => update("masterFolderName", e.target.value)}
                placeholder="Leave blank to use original Master Excel name"
                className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-300 rounded-xl text-sm font-medium text-slate-900 placeholder:text-slate-400 focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 transition-all"
              />
              {config.masterFolderName && (
                <button
                  type="button"
                  onClick={() => update("masterFolderName", "")}
                  className="absolute right-3 top-1/2 -translate-y-1/2 text-xs font-semibold text-slate-400 hover:text-slate-600"
                >
                  Clear
                </button>
              )}
            </div>
            <p className="text-xs text-slate-500 mt-1 flex items-center gap-1">
              <Info className="w-3.5 h-3.5 text-slate-400 shrink-0" />
              Freely assign the folder name created when unzipping (and the ZIP file name).
            </p>
          </div>

          {/* Prefix Name */}
          <div className="lg:col-span-4">
            <label className="block text-xs font-bold uppercase tracking-wider text-slate-600 mb-1.5">
              File Prefix Name
            </label>
            <div className="relative">
              <input
                type="text"
                value={config.prefix}
                onChange={(e) => update("prefix", e.target.value)}
                placeholder="e.g. Export, Sheet, or leave empty"
                className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-300 rounded-xl text-sm font-medium text-slate-900 placeholder:text-slate-400 focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 transition-all"
              />
              {config.prefix && (
                <button
                  type="button"
                  onClick={() => update("prefix", "")}
                  className="absolute right-3 top-1/2 -translate-y-1/2 text-xs font-semibold text-slate-400 hover:text-slate-600"
                >
                  Clear
                </button>
              )}
            </div>
            <p className="text-xs text-slate-500 mt-1 flex items-center gap-1">
              <Info className="w-3.5 h-3.5 text-slate-400 shrink-0" />
              Prepended to each split Excel sheet file.
            </p>
          </div>

          {/* Separator */}
          <div className="lg:col-span-3">
            <label className="block text-xs font-bold uppercase tracking-wider text-slate-600 mb-1.5">
              Prefix Separator
            </label>
            <div className="grid grid-cols-4 gap-1 p-1 bg-slate-100 rounded-xl border border-slate-200">
              {(
                [
                  { label: "Underscore (_)", value: "_" },
                  { label: "Hyphen (-)", value: "-" },
                  { label: "Space ( )", value: " " },
                  { label: "None", value: "" },
                ] as { label: string; value: Separator }[]
              ).map((sep) => (
                <button
                  key={sep.value}
                  type="button"
                  onClick={() => update("separator", sep.value)}
                  className={`py-2 px-1 text-center text-xs font-semibold rounded-lg transition-all ${
                    config.separator === sep.value
                      ? "bg-white text-emerald-800 shadow-sm border border-slate-200"
                      : "text-slate-600 hover:text-slate-900"
                  }`}
                >
                  {sep.value === ""
                    ? "None"
                    : sep.value === " "
                    ? "Space"
                    : sep.value}
                </button>
              ))}
            </div>
            <p className="text-xs text-slate-500 mt-1">
              Symbol after prefix.
            </p>
          </div>
        </div>

        {/* Naming Mode Selector */}
        <div>
          <label className="block text-xs font-bold uppercase tracking-wider text-slate-600 mb-2">
            Behind Name Format (Suffix / Body)
          </label>
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
            {/* Mode 1: Follow Original Sheet Name */}
            <div
              onClick={() => update("namingMode", "original")}
              className={`cursor-pointer p-4 rounded-xl border transition-all ${
                config.namingMode === "original"
                  ? "bg-emerald-50/70 border-emerald-400 ring-2 ring-emerald-500/20 shadow-sm"
                  : "bg-white border-slate-200 hover:border-slate-300 hover:bg-slate-50/50"
              }`}
            >
              <div className="flex items-center gap-2 mb-1.5">
                <FileText
                  className={`w-4 h-4 ${
                    config.namingMode === "original"
                      ? "text-emerald-700"
                      : "text-slate-500"
                  }`}
                />
                <span className="text-sm font-bold text-slate-900">
                  Follow Old Sheet Name
                </span>
                <span className="ml-auto text-[10px] font-semibold uppercase px-1.5 py-0.5 rounded bg-emerald-100 text-emerald-800">
                  Default
                </span>
              </div>
              <p className="text-xs text-slate-500 leading-relaxed">
                Uses the exact original sheet name from the Excel workbook.
              </p>
              <div className="mt-2.5 font-mono text-[11px] px-2 py-1 bg-white/80 rounded border border-slate-200/80 text-slate-700 truncate">
                {config.prefix
                  ? `${config.prefix}${config.separator}SheetName.xlsx`
                  : "SheetName.xlsx"}
              </div>
            </div>

            {/* Mode 2: Sequential Numbers (Small to Big) */}
            <div
              onClick={() => update("namingMode", "sequence")}
              className={`cursor-pointer p-4 rounded-xl border transition-all ${
                config.namingMode === "sequence"
                  ? "bg-emerald-50/70 border-emerald-400 ring-2 ring-emerald-500/20 shadow-sm"
                  : "bg-white border-slate-200 hover:border-slate-300 hover:bg-slate-50/50"
              }`}
            >
              <div className="flex items-center gap-2 mb-1.5">
                <Hash
                  className={`w-4 h-4 ${
                    config.namingMode === "sequence"
                      ? "text-emerald-700"
                      : "text-slate-500"
                  }`}
                />
                <span className="text-sm font-bold text-slate-900">
                  Auto Number Sequence
                </span>
              </div>
              <p className="text-xs text-slate-500 leading-relaxed">
                Number increase in order from small to big (e.g. 1, 2, 3...).
              </p>
              <div className="mt-2.5 font-mono text-[11px] px-2 py-1 bg-white/80 rounded border border-slate-200/80 text-slate-700 truncate">
                {config.prefix
                  ? `${config.prefix}${config.separator}01.xlsx, 02.xlsx...`
                  : "01.xlsx, 02.xlsx..."}
              </div>
            </div>

            {/* Mode 3: Combined Number + Sheet Name */}
            <div
              onClick={() => update("namingMode", "sequence_original")}
              className={`cursor-pointer p-4 rounded-xl border transition-all ${
                config.namingMode === "sequence_original"
                  ? "bg-emerald-50/70 border-emerald-400 ring-2 ring-emerald-500/20 shadow-sm"
                  : "bg-white border-slate-200 hover:border-slate-300 hover:bg-slate-50/50"
              }`}
            >
              <div className="flex items-center gap-2 mb-1.5">
                <Sparkles
                  className={`w-4 h-4 ${
                    config.namingMode === "sequence_original"
                      ? "text-emerald-700"
                      : "text-slate-500"
                  }`}
                />
                <span className="text-sm font-bold text-slate-900">
                  Number + Sheet Name
                </span>
              </div>
              <p className="text-xs text-slate-500 leading-relaxed">
                Numbered index followed by original sheet name.
              </p>
              <div className="mt-2.5 font-mono text-[11px] px-2 py-1 bg-white/80 rounded border border-slate-200/80 text-slate-700 truncate">
                {config.prefix
                  ? `${config.prefix}${config.separator}01_SheetName.xlsx`
                  : "01_SheetName.xlsx"}
              </div>
            </div>
          </div>
        </div>

        {/* Number Sequence Options (Only visible when sequential numbering is active) */}
        {config.namingMode !== "original" && (
          <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 grid grid-cols-1 sm:grid-cols-2 gap-4 animate-in fade-in duration-200">
            <div>
              <label className="block text-xs font-semibold text-slate-700 mb-1.5">
                Starting Number
              </label>
              <input
                type="number"
                min={0}
                value={config.startNumber}
                onChange={(e) =>
                  update("startNumber", Math.max(0, parseInt(e.target.value) || 0))
                }
                className="w-full px-3 py-2 bg-white border border-slate-300 rounded-lg text-sm text-slate-900 focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500"
              />
              <span className="text-[11px] text-slate-500 mt-1 block">
                Sequence increments by +1 for each sheet in order.
              </span>
            </div>

            <div>
              <label className="block text-xs font-semibold text-slate-700 mb-1.5">
                Zero Padding
              </label>
              <div className="grid grid-cols-4 gap-1 p-1 bg-slate-200/60 rounded-lg">
                {(
                  [
                    { label: "1, 2...", value: "none" },
                    { label: "01, 02...", value: "2" },
                    { label: "001...", value: "3" },
                    { label: "Auto", value: "auto" },
                  ] as { label: string; value: ZeroPadding }[]
                ).map((pad) => (
                  <button
                    key={pad.value}
                    type="button"
                    onClick={() => update("zeroPadding", pad.value)}
                    className={`py-1.5 px-2 text-xs font-medium rounded-md transition-all ${
                      config.zeroPadding === pad.value
                        ? "bg-white text-emerald-800 shadow-xs font-bold"
                        : "text-slate-600 hover:text-slate-900"
                    }`}
                  >
                    {pad.label}
                  </button>
                ))}
              </div>
              <span className="text-[11px] text-slate-500 mt-1 block">
                Pads sequence numbers with leading zeros.
              </span>
            </div>
          </div>
        )}

        {/* First Sheet as Info / Menu Option */}
        <div className="p-4 rounded-xl bg-slate-50 border border-slate-200/90 space-y-3">
          <div className="flex items-start gap-3">
            <input
              type="checkbox"
              id="globalFirstSheetIsInfo"
              checked={config.firstSheetIsInfo}
              onChange={(e) => update("firstSheetIsInfo", e.target.checked)}
              className="mt-1 w-4 h-4 text-emerald-600 rounded border-slate-300 focus:ring-emerald-500 cursor-pointer shrink-0"
            />
            <label htmlFor="globalFirstSheetIsInfo" className="text-xs cursor-pointer select-none">
              <span className="font-bold text-slate-900 block text-sm">
                First Sheet is Info / Menu
              </span>
              <span className="text-slate-500 leading-relaxed block mt-0.5">
                When enabled, the 1st sheet is exported as a standalone Info/Menu file without suffixes or numbering. Sequential numbering begins on sheet #2.
              </span>
            </label>
          </div>

          {config.firstSheetIsInfo && (
            <div className="pl-7 pt-2 border-t border-slate-200/70 space-y-1.5 animate-in fade-in duration-150">
              <label className="block text-xs font-bold text-slate-700">
                Info Sheet File Name (Freely customizable)
              </label>
              <div className="flex items-center gap-2 max-w-md">
                <input
                  type="text"
                  value={config.infoFileName ?? ""}
                  onChange={(e) => update("infoFileName", e.target.value)}
                  placeholder={config.prefix.trim() ? `${config.prefix.trim()} (Default: Prefix)` : "Info"}
                  className="w-full px-3 py-2 bg-white border border-slate-300 rounded-lg text-xs font-semibold text-slate-900 placeholder:text-slate-400 focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500"
                />
                <span className="text-xs font-mono font-medium text-slate-500 bg-slate-200/60 px-2.5 py-2 rounded-lg border border-slate-200 shrink-0">
                  .xlsx / .xls
                </span>
                {config.infoFileName && (
                  <button
                    type="button"
                    onClick={() => update("infoFileName", "")}
                    className="text-xs text-slate-400 hover:text-slate-600 font-medium px-1.5 py-1"
                  >
                    Clear
                  </button>
                )}
              </div>
              <p className="text-[11px] text-slate-500">
                Type any custom name for the info sheet (e.g. <span className="font-mono text-emerald-700 font-semibold">Menu</span>, <span className="font-mono text-emerald-700 font-semibold">Index</span>, <span className="font-mono text-emerald-700 font-semibold">Overview</span>). If left blank, it defaults to the Prefix ({config.prefix.trim() || "Info"}).
              </p>
            </div>
          )}
        </div>

        {/* Live Structure Preview Box */}
        <div className="p-4 rounded-xl bg-gradient-to-r from-emerald-950 via-slate-900 to-slate-950 text-white shadow-inner">
          <div className="flex items-center justify-between mb-2">
            <div className="flex items-center gap-2">
              <FolderTree className="w-4 h-4 text-emerald-400" />
              <span className="text-xs font-bold uppercase tracking-wider text-emerald-300">
                Live Output Structure Preview
              </span>
            </div>
            <span className="text-[11px] text-slate-400">
              *Master folder name is customizable per file
            </span>
          </div>

          <div className="font-mono text-xs sm:text-sm bg-black/40 rounded-lg p-3 border border-white/10 space-y-1.5 overflow-x-auto">
            <div className="flex items-center gap-2 text-emerald-300 font-semibold">
              <span>📁 {config.masterFolderName?.trim() || "[Master_Excel_Name]"}/</span>
              <span className="text-[10px] text-slate-400 font-sans font-normal">
                {config.masterFolderName?.trim()
                  ? "(Custom master folder & ZIP archive name)"
                  : "(Folder named after original master Excel file)"}
              </span>
            </div>
            {hasInfo && (
              <div className="pl-6 flex items-center gap-2 text-slate-200">
                <span className="text-slate-500">├──</span>
                <span className="text-amber-400 font-bold">📄 {sampleInfo}</span>
                <span className="text-[10px] text-amber-300/80 font-sans font-normal">
                  (Sheet 1: Info / Menu File)
                </span>
              </div>
            )}
            <div className="pl-6 flex items-center gap-2 text-slate-200">
              <span className="text-slate-500">├──</span>
              <span className="text-emerald-400">📄 {sample1}</span>
              <span className="text-[10px] text-slate-400 font-sans font-normal">
                ({hasInfo ? "Sheet 2" : "Sheet 1"})
              </span>
            </div>
            <div className="pl-6 flex items-center gap-2 text-slate-200">
              <span className="text-slate-500">├──</span>
              <span className="text-emerald-400">📄 {sample2}</span>
              <span className="text-[10px] text-slate-400 font-sans font-normal">
                ({hasInfo ? "Sheet 3" : "Sheet 2"})
              </span>
            </div>
            <div className="pl-6 text-slate-500 text-xs">
              └── ... (each remaining sheet split into its own file)
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
