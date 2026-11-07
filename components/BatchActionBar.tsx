"use client";

import React from "react";
import { Download, Trash2, Plus, FolderTree, FileSpreadsheet } from "lucide-react";

interface BatchActionBarProps {
  fileName?: string;
  totalSheetsSelected: number;
  onDownload: () => void;
  onClear: () => void;
  onChangeFile: () => void;
  isProcessing: boolean;
}

export function BatchActionBar({
  fileName,
  totalSheetsSelected,
  onDownload,
  onClear,
  onChangeFile,
  isProcessing,
}: BatchActionBarProps) {
  if (!fileName) return null;

  return (
    <aside aria-label="File Actions" className="fixed bottom-0 inset-x-0 z-40 p-4 bg-white/95 backdrop-blur-md border-t border-slate-200/90 shadow-2xl">
      <div className="max-w-7xl mx-auto flex flex-col sm:flex-row items-center justify-between gap-4">
        {/* Left Stats */}
        <div className="flex items-center gap-4 text-xs sm:text-sm">
          <div className="flex items-center gap-2">
            <div className="p-1.5 rounded-lg bg-emerald-100 text-emerald-800">
              <FileSpreadsheet className="w-4 h-4" />
            </div>
            <div>
              <span className="font-bold text-slate-900 truncate max-w-[200px] sm:max-w-xs inline-block align-bottom">
                {fileName}
              </span>
            </div>
          </div>

          <span className="text-slate-300">•</span>

          <div className="flex items-center gap-2">
            <div className="p-1.5 rounded-lg bg-emerald-100 text-emerald-800">
              <FolderTree className="w-4 h-4" />
            </div>
            <div>
              <span className="font-bold text-slate-900">{totalSheetsSelected}</span>
              <span className="text-slate-500 ml-1">
                {totalSheetsSelected === 1 ? "Sheet Selected" : "Sheets Selected"}
              </span>
            </div>
          </div>
        </div>

        {/* Right Action Buttons */}
        <div className="flex items-center gap-3 w-full sm:w-auto justify-end">
          <button
            type="button"
            onClick={onClear}
            disabled={isProcessing}
            className="px-3 py-2 text-slate-500 hover:text-rose-600 hover:bg-rose-50 text-xs font-semibold rounded-xl transition-colors flex items-center gap-1.5 disabled:opacity-50"
          >
            <Trash2 className="w-4 h-4" />
            Remove File
          </button>

          <button
            type="button"
            onClick={onChangeFile}
            disabled={isProcessing}
            className="px-3.5 py-2 bg-slate-100 hover:bg-slate-200/80 text-slate-700 text-xs font-semibold rounded-xl transition-colors flex items-center gap-1.5 disabled:opacity-50"
          >
            <Plus className="w-4 h-4" />
            Change File
          </button>

          <button
            type="button"
            onClick={onDownload}
            disabled={isProcessing || totalSheetsSelected === 0}
            className="px-5 py-2.5 bg-emerald-600 hover:bg-emerald-700 active:scale-[0.99] text-white text-xs sm:text-sm font-bold rounded-xl shadow-md shadow-emerald-600/25 transition-all flex items-center gap-2 disabled:opacity-50 disabled:cursor-not-allowed"
          >
            <Download className="w-4 h-4" />
            <span>Split & Download Sheets (ZIP)</span>
          </button>
        </div>
      </div>
    </aside>
  );
}

