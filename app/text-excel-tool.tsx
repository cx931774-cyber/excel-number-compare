"use client";

import { useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx-js-style";
import {
  parseTextHeaders,
  parseTextRows,
  type TextExcelRow,
} from "./text-processing";

const DEFAULT_HEADERS = "手机号 运营商 充值金额 姓名 余额";
const EXAMPLE_TEXT = `13450438325 冯林林 300 移动 余额：192.49
18021306062 俞淑钧 江苏电信 300 余额：112.03
18112613136 金宝 江苏电信 300 余额：100.16`;

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

  function fillExample() {
    setHeaderInput(DEFAULT_HEADERS);
    setTextInput(EXAMPLE_TEXT);
    setResult(null);
    setOutputText("");
    setCopied(false);
    setError("");
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
        <div className="step-row">
          <span>第 1 步</span>
          <strong>输入 Excel 表头</strong>
          <small>空格分隔，每个词作为一列表头</small>
        </div>

        <label className="text-field-label" htmlFor="text-excel-headers">
          表头顺序
        </label>
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
        <div className="header-chip-row" aria-label="识别到的表头">
          {detectedHeaders.length ? (
            detectedHeaders.map((header, index) => (
              <span key={header}>
                <b>{index + 1}</b>
                {header}
              </span>
            ))
          ) : (
            <em>尚未识别到表头</em>
          )}
        </div>

        <div className="text-step-heading">
          <div className="step-row compact">
            <span>第 2 步</span>
            <strong>粘贴需要处理的文本</strong>
            <small>每行生成 Excel 中的一行</small>
          </div>
          <button type="button" onClick={fillExample}>填入示例</button>
        </div>
        <textarea
          className="source-textarea"
          value={textInput}
          placeholder={EXAMPLE_TEXT}
          aria-label="需要处理的文本，每行一条"
          spellCheck={false}
          onChange={(event) => {
            setTextInput(event.target.value);
            setResult(null);
          }}
        />

        <div className="recognition-note">
          <strong>自动识别：</strong>
          手机号、姓名、移动/联通/电信/广电运营商、充值金额，以及“余额：192.49”这样的带标签内容；字段顺序可以不同。
        </div>

        {error && <div className="error-banner" role="alert">{error}</div>}
        <button
          className="compare-button"
          type="button"
          onClick={processText}
        >
          处理文本并生成表格
        </button>
      </section>

      {result && currentResult && (
        <section className="result-card text-result-card" aria-live="polite">
          <div className="result-heading">
            <div>
              <p className="eyebrow">处理完成</p>
              <h2>已整理 {currentResult.rows.length.toLocaleString()} 行文本</h2>
              <p>内容已经按照你输入的表头顺序放入对应列。</p>
            </div>
            <div className="success-badge">可导出</div>
          </div>

          <div className="text-result-summary">
            <div><span>表头列数</span><strong>{currentResult.headers.length}</strong></div>
            <div><span>数据行数</span><strong>{currentResult.rows.length}</strong></div>
            <div className={currentResult.incompleteCount ? "has-warning" : ""}>
              <span>存在空白字段的行</span>
              <strong>{currentResult.incompleteCount}</strong>
            </div>
          </div>

          <div className="editable-output-panel">
            <div className="editable-output-heading">
              <div>
                <strong>结果编辑框</strong>
                <span>第一行为表头，列之间用 Tab 分隔；可修改、全选或复制</span>
              </div>
              <button type="button" onClick={copyOutput}>
                {copied ? "已复制" : "复制全部"}
              </button>
            </div>
            <textarea
              ref={outputRef}
              value={outputText}
              aria-label="整理后的表格内容，可编辑和复制"
              spellCheck={false}
              onChange={(event) => {
                setOutputText(event.target.value);
                setCopied(false);
              }}
            />
          </div>

          <div className="text-preview-wrap">
            <table className="text-preview-table">
              <thead>
                <tr>
                  <th>行号</th>
                  {currentResult.headers.map((header) => <th key={header}>{header}</th>)}
                </tr>
              </thead>
              <tbody>
                {currentResult.rows.slice(0, 20).map((row, rowIndex) => (
                  <tr key={rowIndex}>
                    <td>{rowIndex + 1}</td>
                    {currentResult.headers.map((header) => (
                      <td key={header}>{row[header] || <span className="empty-cell">空</span>}</td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {currentResult.rows.length > 20 && (
            <p className="preview-limit">预览前 20 行，导出的 Excel 包含全部 {currentResult.rows.length.toLocaleString()} 行。</p>
          )}

          <div className="text-result-actions">
            <button
              className="download-button"
              type="button"
              onClick={() =>
                downloadWorkbook(currentResult.headers, currentResult.rows)
              }
            >
              按编辑框内容下载 Excel
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
