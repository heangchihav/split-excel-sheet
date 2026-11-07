"use client";

import React from "react";
import { FileSpreadsheet, ShieldCheck, Layers } from "lucide-react";

export function Navbar() {
  return (
    <header className="border-b border-slate-200 bg-white/80 backdrop-blur-md sticky top-0 z-30">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
        <div className="flex items-center gap-3">
          <div className="h-10 w-10 rounded-xl bg-emerald-600 flex items-center justify-center text-white shadow-md shadow-emerald-500/20">
            <FileSpreadsheet className="h-5 w-5" />
          </div>
          <div>
            <div className="flex items-center gap-2">
              <h1 className="text-lg font-bold text-slate-900 tracking-tight">
                ExcelSheet Splitter
              </h1>
              <span className="px-2 py-0.5 text-xs font-semibold rounded-full bg-emerald-100 text-emerald-800 border border-emerald-200">
                v2.0
              </span>
            </div>
            <p className="text-xs text-slate-500">
              Batch split Excel sheets into separate workbooks inside organized folders
            </p>
          </div>
        </div>

        <div className="hidden sm:flex items-center gap-3">
          <div className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-emerald-50 border border-emerald-200 text-emerald-800 text-xs font-medium">
            <ShieldCheck className="h-4 w-4 text-emerald-600" />
            <span>High-Fidelity Excel Engine</span>
          </div>
          <div className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-slate-100 text-slate-600 text-xs font-medium">
            <Layers className="h-4 w-4 text-slate-500" />
            <span>Multi-File Batch Supported</span>
          </div>
        </div>
      </div>
    </header>
  );
}

