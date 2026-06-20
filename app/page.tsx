"use client";

import { useState, useCallback, useEffect, useRef } from "react";
import { useDropzone } from "react-dropzone";

type Status = "idle" | "extracting" | "preview" | "building" | "done" | "error";

interface IncomeStatement {
  revenue: (number | null)[];
  membership_fees: (number | null)[];
  cogs: (number | null)[];
  gross_profit: (number | null)[];
  sga: (number | null)[];
  ebitda: (number | null)[];
  depreciation_amortization: (number | null)[];
  ebit: (number | null)[];
  interest_expense: (number | null)[];
  ebt: (number | null)[];
  tax: (number | null)[];
  net_income: (number | null)[];
}

interface Suggestion {
  id: string;
  formula: string;
  description: string;
  preview_values: (number | null)[];
}

interface PendingDerivation {
  field: string;
  section: string;
  label: string;
  suggestions: Suggestion[];
}

interface Financials {
  company_name: string;
  ticker: string;
  unit_in_filing: string;
  fiscal_year_end: string;
  years: string[];
  income_statement: IncomeStatement;
  balance_sheet: Record<string, (number | null)[]>;
  cash_flow: Record<string, (number | null)[]>;
  operating_metrics: Record<string, (number | null)[]>;
  shares_outstanding: (number | null)[];
  key_value_drivers: { driver: string; description: string; current_value: number; unit: string }[];
  _raw_strings?: Record<string, unknown>;
  _computed?: string[];
}

const IS_LABELS: Record<string, string> = {
  revenue: "Net Sales / Revenue",
  membership_fees: "Membership Fee Revenue",
  cogs: "Cost of Goods Sold",
  gross_profit: "Gross Profit",
  sga: "SG&A Expenses",
  ebitda: "EBITDA",
  depreciation_amortization: "Depreciation & Amortization",
  ebit: "EBIT",
  interest_expense: "Interest Expense",
  ebt: "Earnings Before Tax",
  tax: "Income Tax",
  net_income: "Net Income",
};

const BS_LABELS: Record<string, string> = {
  cash: "Cash & Equivalents",
  accounts_receivable: "Accounts Receivable",
  inventory: "Inventory",
  total_current_assets: "Total Current Assets",
  ppe_net: "PP&E, Net",
  total_assets: "Total Assets",
  accounts_payable: "Accounts Payable",
  total_current_liabilities: "Total Current Liabilities",
  total_debt: "Total Debt",
  total_equity: "Total Equity",
};

const CF_LABELS: Record<string, string> = {
  cfo: "Cash from Operations",
  capex: "Capital Expenditures",
  fcf: "Free Cash Flow",
  dividends_paid: "Dividends Paid",
};

const OP_LABELS: Record<string, string> = {
  store_count: "Store Count",
  comp_sales_growth: "Comp Sales Growth (%)",
  membership_count: "Membership Count (M)",
  membership_renewal_rate: "Renewal Rate (%)",
};

function fmt(v: number | null | undefined) {
  if (v === null || v === undefined) return "—";
  return v.toLocaleString("en-US", { maximumFractionDigits: 1 });
}

function PreviewTable({
  title,
  labels,
  data,
  years,
  computedKeys,
  onEdit,
}: {
  title: string;
  labels: Record<string, string>;
  data: Record<string, (number | null)[]>;
  years: string[];
  computedKeys: Set<string>;   // "key:yearIdx" pairs that were derived
  onEdit: (section: string, key: string, yearIdx: number, value: string) => void;
}) {
  return (
    <div className="mb-6">
      <h3 className="text-xs font-bold text-gray-500 uppercase tracking-widest mb-2">{title}</h3>
      <div className="bg-white border border-gray-100 rounded-xl overflow-hidden">
        <table className="w-full text-sm">
          <thead>
            <tr className="bg-gray-50">
              <th className="text-left px-4 py-2 text-gray-500 font-medium w-52">Line Item</th>
              {years.map((y) => (
                <th key={y} className="text-right px-4 py-2 text-gray-700 font-semibold">{y}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {Object.entries(labels).map(([key, label]) => (
              <tr key={key} className="border-t border-gray-50 hover:bg-gray-50">
                <td className="px-4 py-1.5 text-gray-600">{label}</td>
                {years.map((_, i) => {
                  const isComputed = computedKeys.has(`${key}:${i}`);
                  return (
                    <td key={i} className="px-2 py-1 text-right">
                      <div className="relative inline-flex items-center gap-1">
                        {isComputed && (
                          <span title="Derived — not directly stated in the filing" className="text-amber-400 text-xs cursor-help">∫</span>
                        )}
                        <input
                          key={`${key}-${i}-${data[key]?.[i]}`}
                          type="text"
                          defaultValue={fmt(data[key]?.[i])}
                          onBlur={(e) => onEdit(title, key, i, e.target.value)}
                          className={`w-28 text-right border focus:outline-none rounded px-2 py-0.5 font-mono text-xs ${
                            isComputed
                              ? "bg-amber-50 border-amber-200 text-amber-800 hover:border-amber-400 focus:border-amber-500"
                              : "border-transparent text-gray-800 hover:border-blue-200 focus:border-blue-400"
                          }`}
                        />
                      </div>
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

export default function Home() {
  const [files, setFiles] = useState<File[]>([]);
  const [status, setStatus] = useState<Status>("idle");
  const [errorMsg, setErrorMsg] = useState("");
  const [financials, setFinancials] = useState<Financials | null>(null);
  const [pendingDerivations, setPendingDerivations] = useState<PendingDerivation[]>([]);
  const [selectedFormulas, setSelectedFormulas] = useState<Record<string, string>>({}); // field -> suggestion id
  const [downloadUrl, setDownloadUrl] = useState("");
  const [fileName, setFileName] = useState("");
  const [showRaw, setShowRaw] = useState(false);
  const [elapsed, setElapsed] = useState(0);
  const timerRef = useRef<ReturnType<typeof setInterval> | null>(null);

  useEffect(() => {
    if (status === "extracting") {
      setElapsed(0);
      timerRef.current = setInterval(() => setElapsed((s) => s + 1), 1000);
    } else {
      if (timerRef.current) clearInterval(timerRef.current);
    }
    return () => { if (timerRef.current) clearInterval(timerRef.current); };
  }, [status]);

  const onDrop = useCallback((accepted: File[]) => {
    setFiles((prev) => [...prev, ...accepted]);
    setStatus("idle");
    setErrorMsg("");
  }, []);

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    accept: {
      "application/pdf": [".pdf"],
      "image/png": [".png"],
      "image/jpeg": [".jpg", ".jpeg"],
      "image/webp": [".webp"],
    },
    multiple: true,
  });

  const removeFile = (index: number) => {
    setFiles((prev) => prev.filter((_, i) => i !== index));
  };

  // Step 1: extract and show preview
  const handleExtract = async () => {
    if (!files.length) return;
    setStatus("extracting");
    setErrorMsg("");

    const formData = new FormData();
    files.forEach((f) => formData.append("files", f));

    try {
      const res = await fetch(`${process.env.NEXT_PUBLIC_API_URL}/api/extract-preview`, {
        method: "POST",
        body: formData,
      });
      if (!res.ok) {
        const err = await res.json();
        throw new Error(err.detail || "Extraction failed");
      }
      const data = await res.json();
      setFinancials(data.financials);
      setPendingDerivations(data.pending_derivations || []);
      // Pre-select first suggestion for each pending item
      const defaults: Record<string, string> = {};
      (data.pending_derivations || []).forEach((d: PendingDerivation) => {
        if (d.suggestions.length > 0) defaults[d.field] = d.suggestions[0].id;
      });
      setSelectedFormulas(defaults);
      setStatus("preview");
    } catch (err: unknown) {
      setErrorMsg(err instanceof Error ? err.message : "Unknown error");
      setStatus("error");
    }
  };

  // Apply approved derivations into financials state
  const applyDerivations = () => {
    if (!financials) return;
    setFinancials((prev) => {
      if (!prev) return prev;
      const updated = JSON.parse(JSON.stringify(prev)) as Financials;
      const newComputed = [...(updated._computed || [])];

      pendingDerivations.forEach((d) => {
        const chosenId = selectedFormulas[d.field];
        if (!chosenId) return;
        const sug = d.suggestions.find((s) => s.id === chosenId);
        if (!sug) return;
        const section = updated[d.section as keyof Financials] as Record<string, (number | null)[]>;
        sug.preview_values.forEach((v, i) => {
          if (section[d.field][i] === null && v !== null) {
            section[d.field][i] = v;
            newComputed.push(`${d.label} = ${sug.formula} [year ${i + 1}]`);
          }
        });
      });

      updated._computed = newComputed;
      return updated;
    });
    setPendingDerivations([]);
  };

  // Inline edit handler
  const handleEdit = (_section: string, key: string, yearIdx: number, value: string) => {
    if (!financials) return;
    const num = parseFloat(value.replace(/,/g, ""));
    if (isNaN(num)) return;

    setFinancials((prev) => {
      if (!prev) return prev;
      const updated = { ...prev };
      // Find which section this key belongs to
      for (const section of ["income_statement", "balance_sheet", "cash_flow", "operating_metrics"] as const) {
        if (key in (updated[section] as Record<string, unknown>)) {
          const arr = [...((updated[section] as Record<string, (number | null)[]>)[key])];
          arr[yearIdx] = num;
          (updated[section] as Record<string, (number | null)[]>)[key] = arr;
          break;
        }
      }
      return updated;
    });
  };

  // Step 2: build model from confirmed financials
  const handleBuild = async () => {
    if (!financials) return;
    setStatus("building");
    setErrorMsg("");

    try {
      const res = await fetch(`${process.env.NEXT_PUBLIC_API_URL}/api/build-model`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ financials }),
      });
      if (!res.ok) {
        const err = await res.json();
        throw new Error(err.detail || "Build failed");
      }
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const cd = res.headers.get("content-disposition") || "";
      const match = cd.match(/filename="?(.+)"?/);
      const name = match ? match[1] : "Financial_Model.xlsx";
      setDownloadUrl(url);
      setFileName(name);
      setStatus("done");
    } catch (err: unknown) {
      setErrorMsg(err instanceof Error ? err.message : "Unknown error");
      setStatus("error");
    }
  };

  const reset = () => {
    setFiles([]);
    setStatus("idle");
    setErrorMsg("");
    setFinancials(null);
    setPendingDerivations([]);
    setSelectedFormulas({});
    if (downloadUrl) URL.revokeObjectURL(downloadUrl);
    setDownloadUrl("");
    setFileName("");
  };

  return (
    <main className="min-h-screen bg-gray-50 flex flex-col items-center pb-16 px-4">
      {/* Nav */}
      <nav className="fixed top-0 left-0 right-0 bg-blue-900 text-white px-8 py-4 flex items-center gap-6 z-10">
        <span className="font-bold text-sm tracking-widest uppercase">Equity Research Platform</span>
        <div className="flex gap-4 text-sm">
          <a href="/" className="text-white font-semibold border-b border-white pb-0.5">Model Generator</a>
          <a href="/analyze" className="text-blue-300 hover:text-white transition-colors">Model Chat</a>
        </div>
      </nav>
      <div className="h-20" />

      {/* ── IDLE: upload ── */}
      {(status === "idle" || status === "error") && (
        <>
          <div className="w-full max-w-2xl mb-8 text-center">
            <h1 className="text-4xl font-bold text-gray-900 mb-3">Financial Model Generator</h1>
            <p className="text-gray-500 text-base">
              Upload 10-K / annual report PDFs. The AI extracts the numbers for you to verify, then builds the Excel model.
            </p>
          </div>

          <div className="w-full max-w-2xl bg-white rounded-2xl shadow-sm border border-gray-100 overflow-hidden">
            <div className="p-8">
              <div
                {...getRootProps()}
                className={`border-2 border-dashed rounded-xl p-10 text-center cursor-pointer transition-colors ${
                  isDragActive ? "border-blue-500 bg-blue-50" : "border-gray-200 hover:border-blue-400 hover:bg-gray-50"
                }`}
              >
                <input {...getInputProps()} />
                <div className="flex flex-col items-center gap-3">
                  <svg className="w-12 h-12 text-blue-300" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5}
                      d="M3 16.5v2.25A2.25 2.25 0 005.25 21h13.5A2.25 2.25 0 0021 18.75V16.5m-13.5-9L12 3m0 0l4.5 4.5M12 3v13.5" />
                  </svg>
                  {isDragActive ? (
                    <p className="text-blue-600 font-medium">Drop PDFs here</p>
                  ) : (
                    <>
                      <p className="text-gray-700 font-medium">Drag & drop financial statements here</p>
                      <p className="text-gray-400 text-sm">or click to browse — PDF or images (PNG, JPG)</p>
                    </>
                  )}
                </div>
              </div>

              {files.length > 0 && (
                <ul className="mt-5 space-y-2">
                  {files.map((f, i) => (
                    <li key={i} className="flex items-center justify-between bg-gray-50 rounded-lg px-4 py-3 text-sm">
                      <div className="flex items-center gap-3 min-w-0">
                        <svg className="w-5 h-5 text-red-400 flex-shrink-0" fill="currentColor" viewBox="0 0 24 24">
                          <path d="M7 3a2 2 0 00-2 2v14a2 2 0 002 2h10a2 2 0 002-2V8l-5-5H7zm5 1l4 4h-4V4z" />
                        </svg>
                        <span className="truncate text-gray-700 font-medium">{f.name}</span>
                        <span className="text-gray-400 flex-shrink-0">{(f.size / 1024).toFixed(0)} KB</span>
                      </div>
                      <button onClick={() => removeFile(i)} className="text-gray-300 hover:text-red-400 ml-3">&times;</button>
                    </li>
                  ))}
                </ul>
              )}
            </div>

            <div className="border-t border-gray-100 px-8 py-5 bg-gray-50 flex items-center justify-between gap-4">
              <p className="text-sm text-gray-400">
                {files.length === 0 ? "No files selected" : `${files.length} file${files.length > 1 ? "s" : ""} ready`}
              </p>
              <div className="flex gap-3">
                {files.length > 0 && (
                  <button onClick={reset} className="text-sm text-gray-400 hover:text-gray-600 transition-colors">Clear</button>
                )}
                <button
                  onClick={handleExtract}
                  disabled={!files.length}
                  className="bg-blue-900 hover:bg-blue-800 disabled:bg-gray-200 disabled:cursor-not-allowed text-white text-sm font-semibold px-6 py-2.5 rounded-lg transition-colors"
                >
                  Extract Data
                </button>
              </div>
            </div>
          </div>

          {status === "error" && (
            <div className="mt-6 w-full max-w-2xl bg-red-50 border border-red-200 rounded-xl p-5">
              <p className="font-semibold text-red-900">Something went wrong</p>
              <p className="text-red-600 text-sm mt-1">{errorMsg}</p>
            </div>
          )}

          <div className="mt-12 w-full max-w-2xl grid grid-cols-4 gap-4 text-center">
            {[
              { step: "01", title: "Upload PDFs", desc: "Annual reports or 10-Ks" },
              { step: "02", title: "AI Extracts", desc: "Tables parsed exactly as reported" },
              { step: "03", title: "You Verify", desc: "Review & correct any number inline" },
              { step: "04", title: "Build Model", desc: "Full Excel with formulas & DCF" },
            ].map((item) => (
              <div key={item.step} className="text-center">
                <div className="text-2xl font-bold text-blue-200 mb-2">{item.step}</div>
                <p className="text-sm font-semibold text-gray-700 mb-1">{item.title}</p>
                <p className="text-xs text-gray-400">{item.desc}</p>
              </div>
            ))}
          </div>
        </>
      )}

      {/* ── EXTRACTING ── */}
      {status === "extracting" && (
        <div className="w-full max-w-2xl mt-8">
          <div className="bg-blue-50 border border-blue-100 rounded-xl p-8 text-center">
            <div className="w-8 h-8 border-2 border-blue-500 border-t-transparent rounded-full animate-spin mx-auto mb-4" />
            <p className="font-semibold text-blue-900 text-lg">Extracting financial data</p>
            <p className="text-blue-600 text-sm mt-2">Reading tables and calling AI — this takes 30–60 seconds</p>
            <p className="text-blue-400 text-xs mt-1 font-mono">{elapsed}s elapsed — still working, please wait…</p>
            <div className="mt-5 space-y-2 text-left max-w-sm mx-auto">
              {["Parsing PDF tables page by page", "Identifying financial statement sections", "Extracting exact numbers as reported", "Detecting units and scale"].map((s) => (
                <div key={s} className="flex items-center gap-2 text-sm text-blue-700">
                  <div className="w-1.5 h-1.5 bg-blue-400 rounded-full animate-pulse" />
                  {s}
                </div>
              ))}
            </div>
          </div>
        </div>
      )}

      {/* ── PREVIEW: verify numbers ── */}
      {(status === "preview" || status === "building" || status === "error") && financials && (
        <div className="w-full max-w-4xl">
          {/* Header */}
          <div className="mb-6 flex items-start justify-between">
            <div>
              <h2 className="text-2xl font-bold text-gray-900">
                {financials.company_name} ({financials.ticker})
              </h2>
              <p className="text-gray-500 text-sm mt-1">
                Fiscal year ends: {financials.fiscal_year_end} · Values in USD millions · Filed in {financials.unit_in_filing}
              </p>
              <p className="text-amber-700 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2 text-sm mt-3 inline-block">
                Review the extracted numbers below. Click any value to edit it, then click <strong>Build Model</strong>.
              </p>
              <div className="mt-2">
                <button
                  onClick={() => setShowRaw((v) => !v)}
                  className="text-xs text-blue-500 hover:text-blue-700 underline"
                >
                  {showRaw ? "Hide raw strings" : "View raw strings extracted from PDF"}
                </button>
              </div>
            </div>
            <div className="flex gap-3 flex-shrink-0 ml-6">
              <button onClick={reset} className="text-sm text-gray-400 hover:text-gray-600 transition-colors px-4 py-2">
                Start over
              </button>
              <button
                onClick={handleBuild}
                disabled={status === "building"}
                className="bg-blue-900 hover:bg-blue-800 disabled:bg-blue-400 text-white text-sm font-semibold px-6 py-2.5 rounded-lg transition-colors"
              >
                {status === "building" ? "Building..." : "Build Model"}
              </button>
            </div>
          </div>

          {showRaw && financials._raw_strings && (
            <div className="mb-6">
              <h3 className="text-xs font-bold text-gray-500 uppercase tracking-widest mb-2">
                Raw strings extracted from PDF (exactly as Claude read them — before Python converts to numbers)
              </h3>
              <pre className="bg-gray-900 text-green-300 text-xs rounded-xl p-5 overflow-auto max-h-96 whitespace-pre-wrap">
                {JSON.stringify(financials._raw_strings, null, 2)}
              </pre>
            </div>
          )}

          {/* ── Pending derivations approval panel ── */}
          {pendingDerivations.length > 0 && (
            <div className="mb-6 bg-white border border-blue-100 rounded-2xl overflow-hidden shadow-sm">
              <div className="bg-blue-900 px-5 py-4 flex items-center justify-between">
                <div>
                  <p className="text-white font-semibold">Missing line items — how should we calculate these?</p>
                  <p className="text-blue-300 text-xs mt-0.5">
                    {pendingDerivations.length} field{pendingDerivations.length > 1 ? "s" : ""} not directly stated in the filing. Select a formula for each and approve.
                  </p>
                </div>
                <button
                  onClick={applyDerivations}
                  className="bg-white text-blue-900 font-semibold text-sm px-5 py-2 rounded-lg hover:bg-blue-50 transition-colors flex-shrink-0"
                >
                  Apply approved formulas
                </button>
              </div>

              <div className="divide-y divide-gray-50">
                {pendingDerivations.map((d) => (
                  <div key={d.field} className="px-5 py-4">
                    <p className="font-semibold text-gray-900 text-sm mb-3">{d.label}</p>
                    <div className="space-y-2">
                      {d.suggestions.map((sug) => {
                        const selected = selectedFormulas[d.field] === sug.id;
                        return (
                          <label
                            key={sug.id}
                            className={`flex items-start gap-3 p-3 rounded-xl border cursor-pointer transition-colors ${
                              selected
                                ? "border-blue-400 bg-blue-50"
                                : "border-gray-100 hover:border-blue-200 hover:bg-gray-50"
                            }`}
                          >
                            <input
                              type="radio"
                              name={d.field}
                              value={sug.id}
                              checked={selected}
                              onChange={() =>
                                setSelectedFormulas((prev) => ({ ...prev, [d.field]: sug.id }))
                              }
                              className="mt-0.5 accent-blue-700 flex-shrink-0"
                            />
                            <div className="flex-1 min-w-0">
                              <p className="text-sm font-semibold text-gray-800">{sug.formula}</p>
                              <p className="text-xs text-gray-500 mt-0.5">{sug.description}</p>
                              {sug.preview_values.some((v) => v !== null) && (
                                <div className="flex gap-4 mt-2">
                                  {financials.years.map((y, i) => (
                                    <div key={y} className="text-center">
                                      <p className="text-xs text-gray-400">{y}</p>
                                      <p className={`text-sm font-mono font-semibold ${selected ? "text-blue-700" : "text-gray-600"}`}>
                                        {sug.preview_values[i] != null
                                          ? sug.preview_values[i]!.toLocaleString("en-US", { maximumFractionDigits: 1 })
                                          : "—"}
                                      </p>
                                    </div>
                                  ))}
                                </div>
                              )}
                            </div>
                          </label>
                        );
                      })}
                      <label
                        className={`flex items-center gap-3 p-3 rounded-xl border cursor-pointer transition-colors ${
                          selectedFormulas[d.field] === "skip"
                            ? "border-gray-300 bg-gray-50"
                            : "border-gray-100 hover:border-gray-300"
                        }`}
                      >
                        <input
                          type="radio"
                          name={d.field}
                          value="skip"
                          checked={selectedFormulas[d.field] === "skip"}
                          onChange={() =>
                            setSelectedFormulas((prev) => ({ ...prev, [d.field]: "skip" }))
                          }
                          className="accent-gray-500 flex-shrink-0"
                        />
                        <p className="text-sm text-gray-400">Leave blank</p>
                      </label>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* Legend */}
          {financials._computed && financials._computed.length > 0 && (
            <div className="mb-4 flex items-start gap-3 bg-amber-50 border border-amber-100 rounded-xl px-4 py-3">
              <span className="text-amber-400 text-lg leading-none mt-0.5">∫</span>
              <div>
                <p className="text-sm font-semibold text-amber-800">
                  {financials._computed.length} value{financials._computed.length > 1 ? "s" : ""} derived — not directly stated in the filing
                </p>
                <p className="text-xs text-amber-600 mt-0.5">
                  {financials._computed.slice(0, 5).join(" · ")}
                  {financials._computed.length > 5 ? ` · +${financials._computed.length - 5} more` : ""}
                </p>
              </div>
            </div>
          )}

          {(() => {
            // Build a Set of "key:yearIdx" for every computed value
            const computedKeys = new Set<string>();
            (financials._computed || []).forEach((label) => {
              const match = label.match(/^(.+?) [=≈]/);
              const yearMatch = label.match(/\[year (\d+)\]/);
              if (match && yearMatch) {
                const fieldName = Object.entries({
                  ...IS_LABELS, ...BS_LABELS, ...CF_LABELS, ...OP_LABELS
                }).find(([, v]) => label.startsWith(v))?.[0]
                  ?? Object.keys({...IS_LABELS,...BS_LABELS,...CF_LABELS,...OP_LABELS}).find(
                    k => label.toLowerCase().includes(k.replace(/_/g," "))
                  );
                if (fieldName) computedKeys.add(`${fieldName}:${parseInt(yearMatch[1]) - 1}`);
              }
            });
            return (
              <>
                <PreviewTable title="Income Statement" labels={IS_LABELS} data={financials.income_statement as unknown as Record<string, (number | null)[]>} years={financials.years} computedKeys={computedKeys} onEdit={handleEdit} />
                <PreviewTable title="Balance Sheet" labels={BS_LABELS} data={financials.balance_sheet} years={financials.years} computedKeys={computedKeys} onEdit={handleEdit} />
                <PreviewTable title="Cash Flow" labels={CF_LABELS} data={financials.cash_flow} years={financials.years} computedKeys={computedKeys} onEdit={handleEdit} />
                <PreviewTable title="Operating Metrics" labels={OP_LABELS} data={financials.operating_metrics} years={financials.years} computedKeys={computedKeys} onEdit={handleEdit} />
              </>
            );
          })()}

          {/* Key value drivers */}
          <div className="mb-6">
            <h3 className="text-xs font-bold text-gray-500 uppercase tracking-widest mb-2">Key Value Drivers (AI-identified)</h3>
            <div className="grid grid-cols-3 gap-4">
              {financials.key_value_drivers.map((d, i) => (
                <div key={i} className="bg-white border border-gray-100 rounded-xl p-4">
                  <p className="font-semibold text-gray-900 text-sm">{d.driver}</p>
                  <p className="text-gray-500 text-xs mt-1">{d.description}</p>
                  <p className="text-blue-700 font-mono text-sm mt-2">{d.current_value} {d.unit}</p>
                </div>
              ))}
            </div>
          </div>

          {status === "building" && (
            <div className="bg-blue-50 border border-blue-100 rounded-xl p-6 text-center">
              <div className="w-6 h-6 border-2 border-blue-500 border-t-transparent rounded-full animate-spin mx-auto mb-3" />
              <p className="font-semibold text-blue-900">Building your Excel model...</p>
              <p className="text-blue-600 text-sm mt-1">Generating all sheets with live formulas</p>
            </div>
          )}

          {status === "error" && (
            <div className="bg-red-50 border border-red-200 rounded-xl p-5 mt-4">
              <p className="font-semibold text-red-900">Build failed</p>
              <p className="text-red-600 text-sm mt-1">{errorMsg}</p>
            </div>
          )}
        </div>
      )}

      {/* ── DONE ── */}
      {status === "done" && downloadUrl && (
        <div className="w-full max-w-2xl mt-8">
          <div className="bg-green-50 border border-green-200 rounded-xl p-8 flex items-center justify-between">
            <div className="flex items-center gap-4">
              <div className="w-12 h-12 bg-green-100 rounded-full flex items-center justify-center">
                <svg className="w-6 h-6 text-green-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 13l4 4L19 7" />
                </svg>
              </div>
              <div>
                <p className="font-semibold text-green-900 text-lg">Model ready</p>
                <p className="text-green-700 text-sm">{fileName}</p>
              </div>
            </div>
            <div className="flex gap-3">
              <a
                href={downloadUrl}
                download={fileName}
                className="bg-green-700 hover:bg-green-600 text-white text-sm font-semibold px-5 py-2.5 rounded-lg transition-colors"
              >
                Download Excel
              </a>
              <button onClick={reset} className="text-sm text-gray-400 hover:text-gray-600 transition-colors">Start over</button>
            </div>
          </div>
        </div>
      )}
    </main>
  );
}
