"use client";

import { useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx-js-style";
import {
  parseTextHeaders,
  parseTextRows,
  type TextExcelRow,
} from "./text-processing";

const DEFAULT_HEADERS = "手机号 运营商 充值金额 姓名 余额";

type TextResult = {
  headers: string[];
  rows: TextExcelRow[];
  incompleteCount: number;
};

function tableToText(headers: string[], rows: TextExcelRow[]) {
  return [
    headers.join("\t"),
    ...rows.map((row) => headers.map((header) => row[header] ?? "").join("\t")),
  ].join("\n");
}

function textToTable(text: string): TextResult | null {
  const lines = text.split(/\r?\n/).filter((line) => line.trim().length > 0);
  if (!lines.length) return null;
  const headers = lines[0].split("\t").map((header) => header.trim()).filter(Boolean);
  if (!headers.length) return null;
  const rows = lines.slice(1).map((line) => {
    const values = line.split("\t");
    return Object.fromEntries(
      headers.map((header, index) => [header, values[index]?.trim() ?? ""]),
    );
  });
  return {
    headers,
    rows,
    incompleteCount: rows.filter((row) =>
      headers.some((header) => !row[header]),
    ).length,
  };
}

function downloadWorkbook(headers: string[], rows: TextExcelRow[]) {
  const values = [
    headers,
    ...rows.map((row) => headers.map((header) => row[header] ?? "")),
  ];
  const sheet = XLSX.utils.aoa_to_sheet(values);
  headers.forEach((_, column) => {
    const address = XLSX.utils.encode_cell({ r: 0, c: column });
    const cell = sheet[address] as XLSX.CellObject & {
      s?: Record<string, unknown>;
    };
    cell.s = {
      font: { bold: true, color: { rgb: "16362D" } },
      fill: {
        patternType: "solid",
        fgColor: { rgb: "E7FF4F" },
        bgColor: { rgb: "E7FF4F" },
      },
      alignment: { horizontal: "center", vertical: "center" },
    };
  });
  sheet["!cols"] = headers.map((header) => ({
    wch: Math.min(
      32,
      Math.max(
        header.length + 4,
        ...rows.map((row) => String(row[header] ?? "").length + 2),
      ),
    ),
  }));
  if (sheet["!ref"]) sheet["!autofilter"] = { ref: sheet["!ref"] };

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheet, "文本处理结果");
  const bytes = XLSX.write(workbook, {
    type: "array",
    bookType: "xlsx",
    cellStyles: true,
    compression: true,
  }) as ArrayBuffer;
  const blob = new Blob([bytes], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "文本处理结果.xlsx";
  document.body.appendChild(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

export default function TextExcelTool() {
  const [headerInput, setHeaderInput] = useState(DEFAULT_HEADERS);
  const [textInput, setTextInput] = useState("");
  const [result, setResult] = useState<TextResult | null>(null);
  const [outputText, setOutputText] = useState("");
  const [copied, setCopied] = useState(false);
  const [error, setError] = useState("");
  const outputRef = useRef<HTMLTextAreaElement>(null);

  const detectedHeaders = useMemo(
    () => parseTextHeaders(headerInput),
    [headerInput],
  );
  const currentResult = useMemo(
    () => (result ? textToTable(outputText) ?? result : null),
    [outputText, result],
  );

  function processText() {
    setError("");
    if (!detectedHeaders.length) {
      setResult(null);
      setError("请先输入表头，表头之间用空格分开。");
      return;
    }
    const rows = parseTextRows(detectedHeaders, textInput);
    if (!rows.length) {
      setResult(null);
      setError("请粘贴至少一行需要处理的文本。");
      return;
    }
    const incompleteCount = rows.filter((row) =>
      detectedHeaders.some((header) => !row[header]),
    ).length;
    setResult({ headers: detectedHeaders, rows, incompleteCount });
    setOutputText(tableToText(detectedHeaders, rows));
    setCopied(false);
  }

  function resetTextTool() {
    setHeaderInput(DEFAULT_HEADERS);
    setTextInput("");
    setResult(null);
    setOutputText("");
    setCopied(false);
    setError("");
  }

  async function copyOutput() {
    const field = outputRef.current;
    if (!field) return;
    try {
      await navigator.clipboard.writeText(field.value);
    } catch {
      field.focus();
      field.select();
      document.execCommand("copy");
    }
    setCopied(true);
    window.setTimeout(() => setCopied(false), 1800);
  }

  return (
    <>
      <section className="workspace-card text-tool-card" aria-label="文本整理成 Excel">
        <div className="simple-field-heading">
          <label htmlFor="text-excel-headers">表头</label>
          <span>用空格分隔，例如：手机号 运营商 充值金额 姓名 余额</span>
        </div>
        <input
          id="text-excel-headers"
          className="header-input"
          type="text"
          value={headerInput}
          placeholder="例如：手机号 运营商 充值金额 姓名 余额"
          onChange={(event) => {
            setHeaderInput(event.target.value);
            setResult(null);
          }}
        />

        <div className="simple-field-heading text-source-heading">
          <label htmlFor="text-excel-source">待处理文本</label>
        </div>
        <textarea
          id="text-excel-source"
          className="source-textarea"
          value={textInput}
          placeholder="请粘贴需要处理的文本，每行一条"
          aria-label="需要处理的文本，每行一条"
          spellCheck={false}
          onChange={(event) => {
            setTextInput(event.target.value);
            setResult(null);
          }}
        />

        {error && <div className="error-banner" role="alert">{error}</div>}
        <button
          className="compare-button"
          type="button"
          onClick={processText}
        >
          开始整理
        </button>
      </section>

      {result && currentResult && (
        <section className="result-card text-result-card" aria-live="polite">
          <div className="compact-result-heading">
            <div>
              <strong>处理结果</strong>
              <span>
                {currentResult.rows.length.toLocaleString()} 行 · {currentResult.headers.length} 列
                {currentResult.incompleteCount
                  ? ` · ${currentResult.incompleteCount} 行有空白字段`
                  : ""}
              </span>
            </div>
            <button type="button" onClick={copyOutput}>
              {copied ? "已复制" : "复制全部"}
            </button>
          </div>

          <textarea
            ref={outputRef}
            className="result-editor"
            value={outputText}
            aria-label="整理后的表格内容，可编辑和复制"
            spellCheck={false}
            onChange={(event) => {
              setOutputText(event.target.value);
              setCopied(false);
            }}
          />

          <div className="text-result-actions">
            <button
              className="download-button"
              type="button"
              onClick={() =>
                downloadWorkbook(currentResult.headers, currentResult.rows)
              }
            >
              下载 Excel
            </button>
            <button className="secondary-button" type="button" onClick={resetTextTool}>
              处理其他文本
            </button>
          </div>
        </section>
      )}
    </>
  );
}
