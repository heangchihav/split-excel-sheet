"use client";

import React from "react";
import { Loader2, FileArchive, CheckCircle2 } from "lucide-react";

interface ProcessingModalProps {
  isOpen: boolean;
  progress: number;
  currentTask: string;
}

export function ProcessingModal({
  isOpen,
  progress,
  currentTask,
}: ProcessingModalProps) {
  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-xs animate-in fade-in duration-150">
      <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 max-w-md w-full p-6 text-center animate-in zoom-in-95 duration-150">
        <div className="w-14 h-14 rounded-2xl bg-emerald-100 text-emerald-700 mx-auto flex items-center justify-center mb-4 shadow-sm">
          {progress >= 100 ? (
            <CheckCircle2 className="w-8 h-8 text-emerald-600 animate-in zoom-in duration-200" />
          ) : (
            <Loader2 className="w-8 h-8 animate-spin text-emerald-600" />
          )}
        </div>

        <h3 className="text-lg font-bold text-slate-900 mb-1">
          {progress >= 100 ? "Splitting Completed!" : "Splitting Excel Sheets..."}
        </h3>
        <p className="text-xs text-slate-500 mb-6 truncate max-w-xs mx-auto">
          {currentTask || "Organizing sheets into folders and compressing..."}
        </p>

        {/* Progress Bar */}
        <div className="w-full bg-slate-100 rounded-full h-3 overflow-hidden p-0.5 border border-slate-200 mb-3">
          <div
            className="bg-emerald-600 h-full rounded-full transition-all duration-200 shadow-sm"
            style={{ width: `${Math.min(100, Math.max(0, progress))}%` }}
          />
        </div>

        <div className="flex items-center justify-between text-xs font-semibold text-slate-600">
          <span className="flex items-center gap-1">
            <FileArchive className="w-3.5 h-3.5 text-slate-400" />
            Packaging separate workbooks
          </span>
          <span className="font-mono text-emerald-700">{progress}%</span>
        </div>
      </div>
    </div>
  );
}

