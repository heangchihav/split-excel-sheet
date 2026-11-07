"use client";

import React, { useState, useRef, DragEvent, ChangeEvent } from "react";
import {
  UploadCloud,
  FileSpreadsheet,
  PlusCircle,
  Sparkles,
  Loader2,
} from "lucide-react";
import { generateSampleExcelFile } from "@/lib/excel-utils";

interface FileDropzoneProps {
  onFileSelected: (file: File) => void;
  isLoading?: boolean;
}

export function FileDropzone({ onFileSelected, isLoading = false }: FileDropzoneProps) {
  const [isDragOver, setIsDragOver] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleDragOver = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragOver(true);
  };

  const handleDragLeave = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragOver(false);
  };

  const handleDrop = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragOver(false);

    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      const file = e.dataTransfer.files[0];
      if (isValidExcelFile(file.name)) {
        onFileSelected(file);
      }
    }
  };

  const handleFileInputChange = (e: ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      const file = e.target.files[0];
      if (isValidExcelFile(file.name)) {
        onFileSelected(file);
      }
      // Reset input value so re-uploading the same file still triggers onChange
      e.target.value = "";
    }
  };

  const handleLoadDemo = () => {
    const demoFile = generateSampleExcelFile();
    onFileSelected(demoFile);
  };

  function isValidExcelFile(name: string): boolean {
    const lower = name.toLowerCase();
    return (
      lower.endsWith(".xlsx") ||
      lower.endsWith(".xls") ||
      lower.endsWith(".xlsm") ||
      lower.endsWith(".xlsb") ||
      lower.endsWith(".csv")
    );
  }

  return (
    <div className="mb-6">
      <input
        ref={fileInputRef}
        type="file"
        accept=".xlsx,.xls,.xlsm,.xlsb,.csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel"
        className="hidden"
        onChange={handleFileInputChange}
      />

      <div
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onDrop={handleDrop}
        onClick={() => fileInputRef.current?.click()}
        className={`relative group cursor-pointer rounded-2xl border-2 border-dashed transition-all duration-200 p-8 sm:p-10 text-center ${
          isDragOver
            ? "border-emerald-500 bg-emerald-50/70 scale-[1.005] shadow-lg shadow-emerald-500/10"
            : "border-slate-300 hover:border-emerald-400 bg-white hover:bg-slate-50/60 shadow-xs"
        }`}
      >
        <div className="max-w-md mx-auto flex flex-col items-center">
          {isLoading ? (
            <div className="p-4 rounded-2xl bg-emerald-100 text-emerald-700 animate-spin mb-4">
              <Loader2 className="w-8 h-8" />
            </div>
          ) : (
            <div
              className={`p-4 rounded-2xl mb-4 transition-all duration-200 ${
                isDragOver
                  ? "bg-emerald-600 text-white scale-110 shadow-lg shadow-emerald-600/30"
                  : "bg-emerald-50 text-emerald-600 group-hover:bg-emerald-100 group-hover:scale-105"
              }`}
            >
              <UploadCloud className="w-8 h-8" />
            </div>
          )}

          <h3 className="text-lg font-bold text-slate-900 mb-1">
            {isLoading ? "Reading Excel File..." : "Upload Master Excel File"}
          </h3>
          <p className="text-sm text-slate-500 mb-4 leading-relaxed">
            Drag & drop your Excel file here, or click to browse from your computer.
          </p>

          <div className="flex flex-wrap items-center justify-center gap-2 mb-4">
            <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-md text-xs font-bold bg-emerald-100 text-emerald-800 border border-emerald-200">
              <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-700" />
              .xlsx
            </span>
            <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-md text-xs font-semibold bg-emerald-100/70 text-emerald-800 border border-emerald-200/60">
              .xls (Excel 97–2003)
            </span>
            <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-md text-xs font-semibold bg-slate-100 text-slate-700">
              .xlsm
            </span>
            <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-md text-xs font-semibold bg-slate-100 text-slate-700">
              .xlsb
            </span>
            <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-md text-xs font-semibold bg-slate-100 text-slate-700">
              .csv
            </span>
          </div>

          <div className="flex flex-wrap items-center justify-center gap-3">
            <button
              type="button"
              className="px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold rounded-xl shadow-sm transition-colors flex items-center gap-1.5"
            >
              <PlusCircle className="w-4 h-4" />
              Select Master Excel File
            </button>

            <button
              type="button"
              onClick={(e) => {
                e.stopPropagation();
                handleLoadDemo();
              }}
              className="px-3.5 py-2 bg-slate-100 hover:bg-slate-200/80 text-slate-700 text-xs font-semibold rounded-xl transition-colors flex items-center gap-1.5"
            >
              <Sparkles className="w-3.5 h-3.5 text-amber-500" />
              Load Demo Multi-Sheet File
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}
