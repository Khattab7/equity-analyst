"use client";

import { useState, useCallback, useRef, useEffect } from "react";
import { useDropzone } from "react-dropzone";

interface Message {
  role: "user" | "assistant";
  content: string;
}

interface ExcelMeta {
  filename: string;
  sheet_names: string[];
  row_counts: Record<string, number>;
  context: string;
}

export default function AnalyzePage() {
  const [excelMeta, setExcelMeta] = useState<ExcelMeta | null>(null);
  const [uploading, setUploading] = useState(false);
  const [uploadError, setUploadError] = useState("");
  const [messages, setMessages] = useState<Message[]>([]);
  const [input, setInput] = useState("");
  const [thinking, setThinking] = useState(false);
  const bottomRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    bottomRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages, thinking]);

  const onDrop = useCallback(async (accepted: File[]) => {
    const file = accepted[0];
    if (!file) return;
    setUploading(true);
    setUploadError("");
    setMessages([]);

    const formData = new FormData();
    formData.append("file", file);

    try {
      const res = await fetch(`${process.env.NEXT_PUBLIC_API_URL}/api/parse-excel`, {
        method: "POST",
        body: formData,
      });
      if (!res.ok) {
        const err = await res.json();
        throw new Error(err.detail || "Upload failed");
      }
      const data = await res.json();
      setExcelMeta(data);
      setMessages([
        {
          role: "assistant",
          content: `I've read **${data.filename}** — ${data.sheet_names.length} sheet${data.sheet_names.length > 1 ? "s" : ""}: ${data.sheet_names.join(", ")}. Ask me anything about the model.`,
        },
      ]);
    } catch (err: unknown) {
      setUploadError(err instanceof Error ? err.message : "Unknown error");
    } finally {
      setUploading(false);
    }
  }, []);

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    accept: {
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"],
      "application/vnd.ms-excel": [".xls"],
    },
    multiple: false,
  });

  const sendMessage = async () => {
    if (!input.trim() || !excelMeta || thinking) return;

    const userMsg: Message = { role: "user", content: input.trim() };
    const updated = [...messages, userMsg];
    setMessages(updated);
    setInput("");
    setThinking(true);

    try {
      const res = await fetch(`${process.env.NEXT_PUBLIC_API_URL}/api/chat`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          excel_context: excelMeta.context,
          messages: updated,
        }),
      });
      if (!res.ok) {
        const err = await res.json();
        throw new Error(err.detail || "Chat failed");
      }
      const data = await res.json();
      setMessages((prev) => [...prev, { role: "assistant", content: data.reply }]);
    } catch (err: unknown) {
      setMessages((prev) => [
        ...prev,
        {
          role: "assistant",
          content: `Error: ${err instanceof Error ? err.message : "Something went wrong"}`,
        },
      ]);
    } finally {
      setThinking(false);
    }
  };

  const handleKey = (e: React.KeyboardEvent) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      sendMessage();
    }
  };

  const reset = () => {
    setExcelMeta(null);
    setMessages([]);
    setInput("");
    setUploadError("");
  };

  return (
    <main className="min-h-screen bg-gray-50 flex flex-col">
      {/* Top nav */}
      <nav className="bg-blue-900 text-white px-8 py-4 flex items-center justify-between">
        <div className="flex items-center gap-6">
          <span className="font-bold text-sm tracking-widest uppercase">
            Equity Research Platform
          </span>
          <div className="flex gap-4 text-sm">
            <a href="/" className="text-blue-300 hover:text-white transition-colors">
              Model Generator
            </a>
            <a href="/analyze" className="text-white font-semibold border-b border-white pb-0.5">
              Model Chat
            </a>
          </div>
        </div>
        {excelMeta && (
          <button
            onClick={reset}
            className="text-blue-300 hover:text-white text-sm transition-colors"
          >
            Load different model
          </button>
        )}
      </nav>

      {!excelMeta ? (
        /* ── Upload state ── */
        <div className="flex-1 flex flex-col items-center justify-center px-4 py-16">
          <div className="w-full max-w-xl text-center mb-8">
            <h1 className="text-3xl font-bold text-gray-900 mb-3">
              Chat with your financial model
            </h1>
            <p className="text-gray-500">
              Upload any Excel model — your own or the one generated here. Ask
              questions in plain English and get analyst-level answers grounded
              in the actual numbers.
            </p>
          </div>

          <div className="w-full max-w-xl">
            <div
              {...getRootProps()}
              className={`border-2 border-dashed rounded-2xl p-14 text-center cursor-pointer transition-colors ${
                isDragActive
                  ? "border-blue-500 bg-blue-50"
                  : "border-gray-200 hover:border-blue-400 hover:bg-gray-50"
              }`}
            >
              <input {...getInputProps()} />
              {uploading ? (
                <div className="flex flex-col items-center gap-3">
                  <div className="w-8 h-8 border-2 border-blue-500 border-t-transparent rounded-full animate-spin" />
                  <p className="text-blue-700 font-medium">Reading model...</p>
                </div>
              ) : (
                <div className="flex flex-col items-center gap-3">
                  <svg className="w-14 h-14 text-green-300" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5}
                      d="M9 17V7m0 10a2 2 0 01-2 2H5a2 2 0 01-2-2V7a2 2 0 012-2h2a2 2 0 012 2m0 10a2 2 0 002 2h2a2 2 0 002-2M9 7a2 2 0 012-2h2a2 2 0 012 2m0 10V7" />
                  </svg>
                  <p className="text-gray-700 font-medium">
                    Drop your Excel model here
                  </p>
                  <p className="text-gray-400 text-sm">.xlsx or .xls files</p>
                </div>
              )}
            </div>
            {uploadError && (
              <p className="mt-3 text-red-500 text-sm text-center">{uploadError}</p>
            )}

            <div className="mt-8 grid grid-cols-3 gap-4 text-center">
              {[
                { step: "01", title: "Upload model", desc: "Any Excel financial model you've built" },
                { step: "02", title: "AI reads it", desc: "Every sheet, every number, every assumption" },
                { step: "03", title: "Ask anything", desc: "Target price, risks, sensitivities, comparisons" },
              ].map((s) => (
                <div key={s.step}>
                  <div className="text-2xl font-bold text-blue-200 mb-1">{s.step}</div>
                  <p className="text-sm font-semibold text-gray-700 mb-1">{s.title}</p>
                  <p className="text-xs text-gray-400">{s.desc}</p>
                </div>
              ))}
            </div>
          </div>
        </div>
      ) : (
        /* ── Chat state ── */
        <div className="flex-1 flex overflow-hidden" style={{ height: "calc(100vh - 57px)" }}>
          {/* Sidebar */}
          <aside className="w-64 bg-white border-r border-gray-100 flex flex-col p-5 flex-shrink-0">
            <p className="text-xs text-gray-400 uppercase tracking-widest mb-3 font-semibold">
              Model loaded
            </p>
            <div className="bg-green-50 border border-green-100 rounded-lg px-3 py-2 mb-5">
              <p className="text-sm font-semibold text-green-900 truncate">
                {excelMeta.filename}
              </p>
            </div>

            <p className="text-xs text-gray-400 uppercase tracking-widest mb-2 font-semibold">
              Sheets
            </p>
            <ul className="space-y-1 mb-6">
              {excelMeta.sheet_names.map((name) => (
                <li key={name} className="flex items-center justify-between text-sm">
                  <span className="text-gray-700 truncate">{name}</span>
                  <span className="text-gray-300 text-xs ml-2 flex-shrink-0">
                    {excelMeta.row_counts[name]}r
                  </span>
                </li>
              ))}
            </ul>

            <p className="text-xs text-gray-400 uppercase tracking-widest mb-2 font-semibold">
              Suggested questions
            </p>
            <ul className="space-y-2">
              {[
                "What is the target price and upside?",
                "What are the key assumptions driving the valuation?",
                "What's the EBITDA margin trend?",
                "Walk me through the DCF logic",
                "What are the biggest risks to the bull case?",
              ].map((q) => (
                <li key={q}>
                  <button
                    onClick={() => setInput(q)}
                    className="text-xs text-left text-blue-600 hover:text-blue-800 hover:underline transition-colors"
                  >
                    {q}
                  </button>
                </li>
              ))}
            </ul>
          </aside>

          {/* Chat panel */}
          <div className="flex-1 flex flex-col min-w-0">
            {/* Messages */}
            <div className="flex-1 overflow-y-auto px-6 py-6 space-y-5">
              {messages.map((msg, i) => (
                <div
                  key={i}
                  className={`flex ${msg.role === "user" ? "justify-end" : "justify-start"}`}
                >
                  {msg.role === "assistant" && (
                    <div className="w-7 h-7 rounded-full bg-blue-900 flex items-center justify-center text-white text-xs font-bold mr-3 flex-shrink-0 mt-0.5">
                      AI
                    </div>
                  )}
                  <div
                    className={`max-w-2xl rounded-2xl px-5 py-3 text-sm leading-relaxed whitespace-pre-wrap ${
                      msg.role === "user"
                        ? "bg-blue-900 text-white rounded-br-sm"
                        : "bg-white border border-gray-100 text-gray-800 rounded-bl-sm shadow-sm"
                    }`}
                  >
                    {msg.content}
                  </div>
                </div>
              ))}

              {thinking && (
                <div className="flex justify-start">
                  <div className="w-7 h-7 rounded-full bg-blue-900 flex items-center justify-center text-white text-xs font-bold mr-3 flex-shrink-0">
                    AI
                  </div>
                  <div className="bg-white border border-gray-100 rounded-2xl rounded-bl-sm px-5 py-3 shadow-sm">
                    <div className="flex gap-1.5 items-center h-5">
                      <div className="w-2 h-2 bg-gray-300 rounded-full animate-bounce" style={{ animationDelay: "0ms" }} />
                      <div className="w-2 h-2 bg-gray-300 rounded-full animate-bounce" style={{ animationDelay: "150ms" }} />
                      <div className="w-2 h-2 bg-gray-300 rounded-full animate-bounce" style={{ animationDelay: "300ms" }} />
                    </div>
                  </div>
                </div>
              )}
              <div ref={bottomRef} />
            </div>

            {/* Input */}
            <div className="border-t border-gray-100 bg-white px-6 py-4">
              <div className="flex gap-3 items-end">
                <textarea
                  value={input}
                  onChange={(e) => setInput(e.target.value)}
                  onKeyDown={handleKey}
                  placeholder="Ask about the model — target price, assumptions, risks, DCF logic..."
                  rows={2}
                  className="flex-1 resize-none border border-gray-200 rounded-xl px-4 py-3 text-sm text-gray-800 placeholder-gray-400 focus:outline-none focus:border-blue-400 transition-colors"
                />
                <button
                  onClick={sendMessage}
                  disabled={!input.trim() || thinking}
                  className="bg-blue-900 hover:bg-blue-800 disabled:bg-gray-200 disabled:cursor-not-allowed text-white px-5 py-3 rounded-xl text-sm font-semibold transition-colors flex-shrink-0"
                >
                  Send
                </button>
              </div>
              <p className="text-xs text-gray-300 mt-2">Enter to send · Shift+Enter for new line</p>
            </div>
          </div>
        </div>
      )}
    </main>
  );
}
