"use client";

import React from "react";
import { X, Table, Eye } from "lucide-react";
import { SheetItem } from "@/types/excel";

interface SheetPreviewModalProps {
  sheet: SheetItem | null;
  masterFileName: string;
  onClose: () => void;
}

export function SheetPreviewModal({
  sheet,
  masterFileName,
  onClose,
}: SheetPreviewModalProps) {
  if (!sheet) return null;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-xs animate-in fade-in duration-150">
      <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 max-w-4xl w-full max-h-[85vh] flex flex-col overflow-hidden animate-in zoom-in-95 duration-150">
        {/* Modal Header */}
        <div className="px-6 py-4 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="p-2 rounded-lg bg-emerald-100 text-emerald-800">
              <Table className="w-5 h-5" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h3 className="text-base font-bold text-slate-900">
                  Sheet Preview: {sheet.originalName}
                </h3>
                <span className="text-xs px-2 py-0.5 rounded-full bg-slate-200 text-slate-700 font-medium">
                  {masterFileName}
                </span>
              </div>
              <p className="text-xs text-slate-500 mt-0.5">
                Showing sample rows from this worksheet ({sheet.rowCount} rows,{" "}
                {sheet.colCount} cols)
              </p>
            </div>
          </div>

          <button
            onClick={onClose}
            className="p-1.5 rounded-lg text-slate-400 hover:text-slate-700 hover:bg-slate-200/60 transition-colors"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Modal Body / Table */}
        <div className="p-6 overflow-auto flex-1">
          {(!sheet.sampleHeaders || sheet.sampleHeaders.length === 0) &&
          (!sheet.sampleRows || sheet.sampleRows.length === 0) ? (
            <div className="py-12 text-center text-slate-400">
              <Eye className="w-8 h-8 mx-auto mb-2 opacity-40" />
              <p className="text-sm">This sheet appears to be empty or has no data range.</p>
            </div>
          ) : (
            <div className="border border-slate-200 rounded-xl overflow-hidden shadow-xs">
              <table className="w-full text-left text-xs text-slate-700 border-collapse">
                {sheet.sampleHeaders && sheet.sampleHeaders.length > 0 && (
                  <thead className="bg-slate-100/90 text-slate-800 font-bold border-b border-slate-200">
                    <tr>
                      <th className="px-3 py-2.5 text-center w-10 text-slate-400 font-mono text-[11px] border-r border-slate-200">
                        #
                      </th>
                      {sheet.sampleHeaders.map((col, idx) => (
                        <th
                          key={idx}
                          className="px-3 py-2.5 whitespace-nowrap border-r border-slate-200 last:border-r-0"
                        >
                          {col || `Col ${idx + 1}`}
                        </th>
                      ))}
                    </tr>
                  </thead>
                )}
                <tbody className="divide-y divide-slate-200">
                  {sheet.sampleRows?.map((row, rIdx) => (
                    <tr
                      key={rIdx}
                      className={rIdx % 2 === 0 ? "bg-white" : "bg-slate-50/50"}
                    >
                      <td className="px-3 py-2 text-center text-slate-400 font-mono text-[11px] bg-slate-100/50 border-r border-slate-200">
                        {rIdx + 1}
                      </td>
                      {row.map((val, cIdx) => (
                        <td
                          key={cIdx}
                          className="px-3 py-2 whitespace-nowrap font-mono text-slate-800 border-r border-slate-200 last:border-r-0"
                        >
                          {val !== "" ? val : <span className="text-slate-300 italic">null</span>}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>

        {/* Modal Footer */}
        <div className="px-6 py-3 bg-slate-50 border-t border-slate-200 flex justify-between items-center text-xs text-slate-500">
          <span>*Only first rows are rendered in this fast preview. Full sheet data is exported during split.</span>
          <button
            onClick={onClose}
            className="px-4 py-2 bg-slate-800 hover:bg-slate-900 text-white rounded-lg font-semibold transition-colors"
          >
            Close Preview
          </button>
        </div>
      </div>
    </div>
  );
}
