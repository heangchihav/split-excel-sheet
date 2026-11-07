"use client";

import React, { useState, useRef } from "react";
import { Navbar } from "@/components/Navbar";
import { GlobalSettingsCard } from "@/components/GlobalSettingsCard";
import { FileDropzone } from "@/components/FileDropzone";
import { WorkbookItemCard } from "@/components/WorkbookItemCard";
import { SheetPreviewModal } from "@/components/SheetPreviewModal";
import { ProcessingModal } from "@/components/ProcessingModal";
import { BatchActionBar } from "@/components/BatchActionBar";
import { GlobalConfig, SheetItem, UploadedWorkbook } from "@/types/excel";
import {
  extractWorkbookData,
  createWorkbookZip,
  triggerBlobDownload,
  sanitizeFileName,
} from "@/lib/excel-utils";
import {
  CheckCircle2,
  HelpCircle,
  FileSpreadsheet,
} from "lucide-react";

const DEFAULT_GLOBAL_CONFIG: GlobalConfig = {
  prefix: "",
  namingMode: "original",
  startNumber: 1,
  zeroPadding: "2",
  separator: "_",
  firstSheetIsInfo: false,
  infoFileName: "",
  masterFolderName: "",
};

export default function Home() {
  const [globalConfig, setGlobalConfig] = useState<GlobalConfig>(() => {
    if (typeof window !== "undefined") {
      try {
        const saved = localStorage.getItem("excel_splitter_config");
        if (saved) {
          return { ...DEFAULT_GLOBAL_CONFIG, ...JSON.parse(saved) };
        }
      } catch {
        // Ignore storage errors
      }
    }
    return DEFAULT_GLOBAL_CONFIG;
  });

  const [workbook, setWorkbook] = useState<UploadedWorkbook | null>(null);
  const [isLoadingFile, setIsLoadingFile] = useState(false);

  const handleGlobalConfigChange = (newConfig: GlobalConfig) => {
    setGlobalConfig(newConfig);
    try {
      localStorage.setItem("excel_splitter_config", JSON.stringify(newConfig));
    } catch {
      // Ignore storage errors
    }
  };

  const handleResetGlobalConfig = () => {
    setGlobalConfig(DEFAULT_GLOBAL_CONFIG);
    try {
      localStorage.removeItem("excel_splitter_config");
    } catch {
      // Ignore storage errors
    }
  };

  // Preview modal state
  const [previewState, setPreviewState] = useState<{
    sheet: SheetItem | null;
    masterName: string;
  }>({
    sheet: null,
    masterName: "",
  });

  // Download progress state
  const [isProcessing, setIsProcessing] = useState(false);
  const [processProgress, setProcessProgress] = useState(0);
  const [currentTask, setCurrentTask] = useState("");

  const replaceFileInputRef = useRef<HTMLInputElement>(null);

  // Single file loading handler
  const handleFileSelected = async (file: File) => {
    setIsLoadingFile(true);
    try {
      const parsed = await extractWorkbookData(file);
      setWorkbook(parsed);
    } catch (err) {
      console.error(`Error parsing ${file.name}:`, err);
      alert(
        `Could not read "${file.name}". Please ensure it is a valid Excel spreadsheet.`
      );
    } finally {
      setIsLoadingFile(false);
    }
  };

  const handleUpdateWorkbook = (updated: UploadedWorkbook) => {
    setWorkbook(updated);
  };

  const handleRemoveWorkbook = () => {
    setWorkbook(null);
  };

  const handleClear = () => {
    if (confirm("Are you sure you want to remove the uploaded Excel file?")) {
      setWorkbook(null);
    }
  };

  // Split & Download
  const handleDownload = async () => {
    if (!workbook) return;

    try {
      setIsProcessing(true);
      setProcessProgress(20);
      setCurrentTask("Preparing workbook sheets...");

      const zipBlob = await createWorkbookZip(
        workbook,
        globalConfig
      );

      setProcessProgress(100);
      setCurrentTask("Download ready!");

      setTimeout(() => {
        const customGlobalFolder = globalConfig.masterFolderName?.trim();
        const targetName =
          customGlobalFolder ||
          workbook.folderName?.trim() ||
          workbook.fileName.replace(/\.[^/.]+$/, "");
        const zipName = `${sanitizeFileName(targetName, "Split_Excel_Sheets")}.zip`;

        triggerBlobDownload(zipBlob, zipName);
        setIsProcessing(false);
      }, 500);
    } catch (err) {
      console.error("Split error:", err);
      alert("An error occurred while splitting the workbook.");
      setIsProcessing(false);
    }
  };

  const totalSheetsSelected = workbook
    ? workbook.sheets.filter((s) => s.selected).length
    : 0;

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col">
      <Navbar />

      <main className="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-8 pb-32">
        {/* Hero Section */}
        <div className="text-center max-w-3xl mx-auto mb-8">
          <div className="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-emerald-100/80 border border-emerald-200 text-emerald-800 text-xs font-semibold mb-3">
            <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600" />
            Auto-Split Multi-Sheet Workbook into Independent Files
          </div>
          <h1 className="text-3xl sm:text-4xl font-extrabold text-slate-900 tracking-tight mb-3">
            Split Excel Sheets into Dedicated Folders
          </h1>
          <p className="text-sm sm:text-base text-slate-600 leading-relaxed">
            Upload your master Excel file. Each sheet is cleanly extracted into its
            own workbook inside a dedicated folder named after the master file (or custom named).
            Customize prefixes, sequential numbering, and Info/Menu sheet rules.
          </p>

          {/* 3 Step Workflow */}
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 mt-6 text-left">
            <div className="p-3 bg-white rounded-xl border border-slate-200 flex items-start gap-2.5 shadow-2xs">
              <span className="w-6 h-6 rounded-lg bg-emerald-100 text-emerald-800 font-bold text-xs flex items-center justify-center shrink-0">
                1
              </span>
              <div>
                <h2 className="text-xs font-bold text-slate-900">Upload Master File</h2>
                <p className="text-[11px] text-slate-500">
                  Select or drag your .xlsx, .xls, or .csv file.
                </p>
              </div>
            </div>

            <div className="p-3 bg-white rounded-xl border border-slate-200 flex items-start gap-2.5 shadow-2xs">
              <span className="w-6 h-6 rounded-lg bg-emerald-100 text-emerald-800 font-bold text-xs flex items-center justify-center shrink-0">
                2
              </span>
              <div>
                <h2 className="text-xs font-bold text-slate-900">Set Prefix & Naming</h2>
                <p className="text-[11px] text-slate-500">
                  Choose prefix, sequence numbering (1, 2, 3...) or sheet names.
                </p>
              </div>
            </div>

            <div className="p-3 bg-white rounded-xl border border-slate-200 flex items-start gap-2.5 shadow-2xs">
              <span className="w-6 h-6 rounded-lg bg-emerald-100 text-emerald-800 font-bold text-xs flex items-center justify-center shrink-0">
                3
              </span>
              <div>
                <h2 className="text-xs font-bold text-slate-900">Download Split ZIP</h2>
                <p className="text-[11px] text-slate-500">
                  All sheets extracted into separate files inside your master folder.
                </p>
              </div>
            </div>
          </div>
        </div>

        {/* Global Settings Panel */}
        <GlobalSettingsCard
          config={globalConfig}
          onChange={handleGlobalConfigChange}
          onReset={handleResetGlobalConfig}
        />

        {/* File Dropzone - visible when no file is uploaded */}
        {!workbook && (
          <FileDropzone
            onFileSelected={handleFileSelected}
            isLoading={isLoadingFile}
          />
        )}

        {/* Hidden input to support "Change File" / "Replace File" */}
        <input
          ref={replaceFileInputRef}
          type="file"
          accept=".xlsx,.xls,.xlsm,.xlsb,.csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel"
          className="hidden"
          onChange={(e) => {
            if (e.target.files && e.target.files.length > 0) {
              handleFileSelected(e.target.files[0]);
              e.target.value = "";
            }
          }}
        />

        {/* Uploaded File Section */}
        {workbook && (
          <div className="mt-8">
            <div className="flex items-center justify-between mb-4">
              <div className="flex items-center gap-2">
                <FileSpreadsheet className="w-5 h-5 text-emerald-600" />
                <h2 className="text-base font-bold text-slate-900">
                  Uploaded Master File
                </h2>
              </div>
              <button
                type="button"
                onClick={() => replaceFileInputRef.current?.click()}
                className="text-xs font-semibold text-emerald-700 hover:text-emerald-800 hover:bg-emerald-50 px-3 py-1.5 rounded-lg border border-emerald-200 transition-colors flex items-center gap-1.5"
              >
                Change / Replace File
              </button>
            </div>

            <div className="space-y-4">
              <WorkbookItemCard
                workbook={workbook}
                globalConfig={globalConfig}
                onUpdateWorkbook={handleUpdateWorkbook}
                onRemoveWorkbook={handleRemoveWorkbook}
                onPreviewSheet={(sheet, masterName) =>
                  setPreviewState({ sheet, masterName })
                }
              />
            </div>
          </div>
        )}

        {/* Feature Highlights / FAQ Banner */}
        {!workbook && (
          <div className="mt-12 p-6 rounded-2xl bg-white border border-slate-200">
            <div className="flex items-center gap-2 mb-4">
              <HelpCircle className="w-5 h-5 text-emerald-600" />
              <h2 className="text-sm font-bold text-slate-900">
                Key Features & How It Works
              </h2>
            </div>
            <div className="grid grid-cols-1 md:grid-cols-3 gap-6 text-xs text-slate-600 leading-relaxed">
              <div>
                <strong className="text-slate-800 block mb-1">
                  1. Dedicated Master Folder Packaging
                </strong>
                Your master Excel file sheets are saved inside a dedicated folder
                inside the ZIP archive. You can freely name this folder or use the original master file name.
              </div>
              <div>
                <strong className="text-slate-800 block mb-1">
                  2. Flexible File Naming & Numbering
                </strong>
                Add custom prefixes (e.g. <code>Dept_</code>) and choose whether
                the suffix follows the original sheet name or increments
                sequentially (e.g. <code>01, 02, 03...</code> from small to big).
              </div>
              <div>
                <strong className="text-slate-800 block mb-1">
                  3. High-Fidelity Font & Style Preservation
                </strong>
                Apache POI isolation preserves 100% of your Khmer fonts,
                formulas, borders, column widths, and cell styling with zero
                repair warnings in Microsoft Excel.
              </div>
            </div>
          </div>
        )}
      </main>

      {/* Floating Action Toolbar */}
      <BatchActionBar
        fileName={workbook?.fileName}
        totalSheetsSelected={totalSheetsSelected}
        onDownload={handleDownload}
        onClear={handleClear}
        onChangeFile={() => replaceFileInputRef.current?.click()}
        isProcessing={isProcessing}
      />

      {/* Sheet Data Preview Modal */}
      <SheetPreviewModal
        sheet={previewState.sheet}
        masterFileName={previewState.masterName}
        onClose={() => setPreviewState({ sheet: null, masterName: "" })}
      />

      {/* Processing & Progress Modal */}
      <ProcessingModal
        isOpen={isProcessing}
        progress={processProgress}
        currentTask={currentTask}
      />
    </div>
  );
}
