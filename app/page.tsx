"use client";

import { useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx-js-style";

type WorkbookFile = {
  file: File;
  buffer: ArrayBuffer;
  workbook: XLSX.WorkBook;
};

type ColumnOption = {
  index: number;
  letter: string;
  label: string;
  dataStartRow: number;
};

type MatchResult = {
  workbook: XLSX.WorkBook;
  filteredWorkbook: XLSX.WorkBook;
  sourceCount: number;
  targetCount: number;
  sourceMatchRowCount: number;
  matchCount: number;
  uniqueMatchCount: number;
  uniqueUnmatchedCount: number;
  protectedLongIds: number;
  matchedNumbers: string[];
  unmatchedNumbers: string[];
  samples: Array<{ row: number; value: string }>;
  outputName: string;
  filteredOutputName: string;
};

type StyledCell = XLSX.CellObject & {
  s?: Record<string, unknown>;
};

const MAX_FILE_SIZE = 30 * 1024 * 1024;

function cellText(cell?: XLSX.CellObject) {
  if (!cell || cell.v === null || cell.v === undefined) return "";
  if (typeof cell.v === "string") return cell.v.trim();
  if (typeof cell.v === "number") {
    if (Number.isSafeInteger(cell.v)) return String(cell.v);
    const displayed = cell.w || XLSX.utils.format_cell(cell);
    return String(displayed || cell.v).trim();
  }
  return String(cell.v).trim();
}

function normalize(value: string, relaxed: boolean) {
  let normalized = value
    .trim()
    .replace(/^'/, "")
    .replace(/[０-９]/g, (digit) =>
      String.fromCharCode(digit.charCodeAt(0) - 0xfee0),
    );
  if (relaxed) normalized = normalized.replace(/[\s\-—–()（）]/g, "");
  return normalized;
}

function isNumberLike(value: string) {
  return /^\+?\d+$/.test(normalize(value, true));
}

function worksheetColumns(
  workbook: XLSX.WorkBook | null,
  sheetName: string,
): ColumnOption[] {
  if (!workbook || !sheetName) return [];
  const sheet = workbook.Sheets[sheetName];
  if (!sheet || !sheet["!ref"]) return [];
  const used = XLSX.utils.decode_range(sheet["!ref"]);
  const columns: ColumnOption[] = [];
  for (let column = used.s.c; column <= Math.min(used.e.c, 199); column += 1) {
    let firstRow = used.s.r;
    let firstValue = "";
    for (
      let row = used.s.r;
      row <= Math.min(used.e.r, used.s.r + 20);
      row += 1
    ) {
      const value = cellText(sheet[XLSX.utils.encode_cell({ r: row, c: column })]);
      if (value) {
        firstRow = row;
        firstValue = value;
        break;
      }
    }
    if (!firstValue) continue;
    const letter = XLSX.utils.encode_col(column);
    const hasHeader = !isNumberLike(firstValue);
    const preview =
      firstValue.length > 24 ? `${firstValue.slice(0, 24)}…` : firstValue;
    columns.push({
      index: column,
      letter,
      label: hasHeader
        ? `${letter} · ${preview}`
        : `${letter} 列 · 首项 ${preview}`,
      dataStartRow: hasHeader ? firstRow + 1 : firstRow,
    });
  }
  return columns;
}

function preferredColumn(options: ColumnOption[], target: "source" | "target") {
  if (!options.length) return -1;
  if (target === "target") {
    const likely = options.find((option) =>
      /玩家账号|账号|号码|手机|电话|phone|account/i.test(option.label),
    );
    if (likely) return likely.index;
  }
  return options[0].index;
}

function protectLongIdentifiers(workbook: XLSX.WorkBook) {
  let protectedCount = 0;
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet?.["!ref"]) continue;
    const used = XLSX.utils.decode_range(sheet["!ref"]);
    for (let row = used.s.r; row <= used.e.r; row += 1) {
      for (let column = used.s.c; column <= used.e.c; column += 1) {
        const address = XLSX.utils.encode_cell({ r: row, c: column });
        const cell = sheet[address] as StyledCell | undefined;
        if (!cell || cell.f) continue;
        const value = cellText(cell).replace(/\s/g, "");
        if (/^\d{15,}$/.test(value)) {
          cell.t = "s";
          cell.v = value;
          cell.w = value;
          cell.z = "@";
          protectedCount += 1;
        }
      }
    }
  }
  return protectedCount;
}

function extractRows(
  sheet: XLSX.WorkSheet,
  dataStartRow: number,
  selectedRows: number[],
) {
  const extracted: XLSX.WorkSheet = {};
  if (!sheet["!ref"]) return extracted;

  const used = XLSX.utils.decode_range(sheet["!ref"]);
  const headerRows = Array.from(
    { length: Math.max(0, dataStartRow - used.s.r) },
    (_, index) => used.s.r + index,
  );
  const dataRows = [...new Set(selectedRows)]
    .filter((row) => row >= dataStartRow && row <= used.e.r)
    .sort((left, right) => left - right);
  const rowsToCopy = [...headerRows, ...dataRows];
  if (!rowsToCopy.length) return extracted;

  rowsToCopy.forEach((sourceRow, outputRow) => {
    for (let column = used.s.c; column <= used.e.c; column += 1) {
      const sourceAddress = XLSX.utils.encode_cell({ r: sourceRow, c: column });
      const sourceCell = sheet[sourceAddress];
      if (!sourceCell) continue;
      const outputAddress = XLSX.utils.encode_cell({ r: outputRow, c: column });
      extracted[outputAddress] = JSON.parse(JSON.stringify(sourceCell));
    }
  });

  extracted["!ref"] = XLSX.utils.encode_range({
    s: { r: 0, c: used.s.c },
    e: { r: rowsToCopy.length - 1, c: used.e.c },
  });
  if (sheet["!cols"]) {
    extracted["!cols"] = sheet["!cols"].map((column) =>
      column ? { ...column } : column,
    );
  }
  if (sheet["!rows"]) {
    extracted["!rows"] = rowsToCopy.map((row) => {
      const sourceRow = sheet["!rows"]?.[row];
      return sourceRow ? { ...sourceRow } : sourceRow;
    });
  }
  return extracted;
}

async function parseWorkbook(file: File): Promise<WorkbookFile> {
  if (!/\.xlsx?$/i.test(file.name)) {
    throw new Error("请选择 .xlsx 或 .xls 格式的 Excel 文件。");
  }
  if (file.size > MAX_FILE_SIZE) {
    throw new Error("单个文件请控制在 30 MB 以内。");
  }
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, {
    type: "array",
    cellStyles: true,
    cellNF: true,
    cellDates: false,
  });
  if (!workbook.SheetNames.length) throw new Error("表格中没有可读取的工作表。");
  return { file, buffer, workbook };
}

function UploadCard({
  id,
  number,
  title,
  description,
  value,
  onFile,
}: {
  id: string;
  number: string;
  title: string;
  description: string;
  value: WorkbookFile | null;
  onFile: (file: File) => void;
}) {
  const [dragging, setDragging] = useState(false);
  const size = value ? `${(value.file.size / 1024).toFixed(0)} KB` : "";

  return (
    <label
      className={`upload-zone ${value ? "has-file" : ""} ${dragging ? "is-dragging" : ""}`}
      htmlFor={id}
      onDragEnter={(event) => {
        event.preventDefault();
        setDragging(true);
      }}
      onDragOver={(event) => event.preventDefault()}
      onDragLeave={() => setDragging(false)}
      onDrop={(event) => {
        event.preventDefault();
        setDragging(false);
        const file = event.dataTransfer.files[0];
        if (file) onFile(file);
      }}
    >
      <span className="file-number">{number}</span>
      <strong>{title}</strong>
      <small>{description}</small>
      <input
        id={id}
        type="file"
        accept=".xlsx,.xls"
        onChange={(event) => {
          const file = event.target.files?.[0];
          if (file) onFile(file);
        }}
      />
      {value ? (
        <span className="file-chip">
          <span className="check-mark">✓</span>
          <span className="file-meta">
            <b>{value.file.name}</b>
            <em>
              {size} · {value.workbook.SheetNames.length} 个工作表
            </em>
          </span>
          <span className="replace-label">更换</span>
        </span>
      ) : (
        <span className="upload-button">选择文件或拖到这里</span>
      )}
    </label>
  );
}

export default function Home() {
  const [first, setFirst] = useState<WorkbookFile | null>(null);
  const [second, setSecond] = useState<WorkbookFile | null>(null);
  const [firstSheet, setFirstSheet] = useState("");
  const [secondSheet, setSecondSheet] = useState("");
  const [firstColumn, setFirstColumn] = useState(-1);
  const [secondColumn, setSecondColumn] = useState(-1);
  const [relaxed, setRelaxed] = useState(false);
  const [protectIds, setProtectIds] = useState(true);
  const [showMatchBox, setShowMatchBox] = useState(true);
  const [numberView, setNumberView] = useState<"matched" | "unmatched">(
    "matched",
  );
  const [copied, setCopied] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [result, setResult] = useState<MatchResult | null>(null);
  const [resetKey, setResetKey] = useState(0);
  const resultRef = useRef<HTMLElement>(null);
  const matchBoxRef = useRef<HTMLTextAreaElement>(null);

  const firstColumns = useMemo(
    () => worksheetColumns(first?.workbook ?? null, firstSheet),
    [first, firstSheet],
  );
  const secondColumns = useMemo(
    () => worksheetColumns(second?.workbook ?? null, secondSheet),
    [second, secondSheet],
  );

  async function loadFile(position: "first" | "second", file: File) {
    setError("");
    setResult(null);
    try {
      const loaded = await parseWorkbook(file);
      const sheet = loaded.workbook.SheetNames[0];
      const options = worksheetColumns(loaded.workbook, sheet);
      if (!options.length) throw new Error("找不到可用于比对的数据列。");
      if (position === "first") {
        setFirst(loaded);
        setFirstSheet(sheet);
        setFirstColumn(preferredColumn(options, "source"));
      } else {
        setSecond(loaded);
        setSecondSheet(sheet);
        setSecondColumn(preferredColumn(options, "target"));
      }
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "无法读取这个文件。");
    }
  }

  function changeSheet(position: "first" | "second", sheetName: string) {
    setResult(null);
    if (position === "first" && first) {
      setFirstSheet(sheetName);
      setFirstColumn(
        preferredColumn(worksheetColumns(first.workbook, sheetName), "source"),
      );
    }
    if (position === "second" && second) {
      setSecondSheet(sheetName);
      setSecondColumn(
        preferredColumn(worksheetColumns(second.workbook, sheetName), "target"),
      );
    }
  }

  async function compare() {
    if (!first || !second || firstColumn < 0 || secondColumn < 0) return;
    setBusy(true);
    setError("");
    setResult(null);
    await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));

    try {
      const sourceSheet = first.workbook.Sheets[firstSheet];
      const sourceOption = firstColumns.find(
        (option) => option.index === firstColumn,
      );
      const targetOption = secondColumns.find(
        (option) => option.index === secondColumn,
      );
      if (!sourceSheet?.["!ref"] || !sourceOption || !targetOption) {
        throw new Error("请选择有效的工作表和比对列。");
      }

      const outputWorkbook = XLSX.read(second.buffer.slice(0), {
        type: "array",
        cellStyles: true,
        cellNF: true,
        cellDates: false,
      });
      const targetSheet = outputWorkbook.Sheets[secondSheet];
      if (!targetSheet?.["!ref"]) throw new Error("第二个工作表没有可读取的数据。");

      const sourceRange = XLSX.utils.decode_range(sourceSheet["!ref"]);
      const sourceValues = new Set<string>();
      const sourceRowsByValue = new Map<string, number[]>();
      let sourceCount = 0;
      for (let row = sourceOption.dataStartRow; row <= sourceRange.e.r; row += 1) {
        const cell = sourceSheet[
          XLSX.utils.encode_cell({ r: row, c: firstColumn })
        ];
        const value = normalize(cellText(cell), relaxed);
        if (value) {
          sourceCount += 1;
          sourceValues.add(value);
          const rows = sourceRowsByValue.get(value) ?? [];
          rows.push(row);
          sourceRowsByValue.set(value, rows);
        }
      }

      const targetRange = XLSX.utils.decode_range(targetSheet["!ref"]);
      const matches: Array<{
        row: number;
        sourceRow: number;
        value: string;
        address: string;
      }> = [];
      const unmatchedNumbers = new Map<string, string>();
      const unmatchedRows: number[] = [];
      let targetCount = 0;
      for (let row = targetOption.dataStartRow; row <= targetRange.e.r; row += 1) {
        const address = XLSX.utils.encode_cell({ r: row, c: secondColumn });
        const cell = targetSheet[address];
        const displayValue = cellText(cell);
        const value = normalize(displayValue, relaxed);
        if (!value) continue;
        targetCount += 1;
        if (sourceValues.has(value)) {
          matches.push({
            row: row + 1,
            sourceRow: row,
            value: displayValue,
            address,
          });
        } else if (!unmatchedNumbers.has(value)) {
          unmatchedNumbers.set(value, displayValue);
          unmatchedRows.push(row);
        } else {
          unmatchedRows.push(row);
        }
      }

      const protectedLongIds = protectIds
        ? protectLongIdentifiers(outputWorkbook)
        : 0;
      for (const match of matches) {
        const cell = targetSheet[match.address] as StyledCell;
        cell.s = {
          ...(cell.s ?? {}),
          fill: {
            patternType: "solid",
            fgColor: { rgb: "FFFF00" },
            bgColor: { rgb: "FFFF00" },
          },
        };
      }

      const matchedKeys = new Set(
        matches.map((match) => normalize(match.value, relaxed)),
      );
      const sourceMatchRows = [...matchedKeys].flatMap(
        (key) => sourceRowsByValue.get(key) ?? [],
      );
      const filteredWorkbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(
        filteredWorkbook,
        extractRows(sourceSheet, sourceOption.dataStartRow, sourceMatchRows),
        "相同行_第一表",
      );
      XLSX.utils.book_append_sheet(
        filteredWorkbook,
        extractRows(
          targetSheet,
          targetOption.dataStartRow,
          matches.map((match) => match.sourceRow),
        ),
        "相同行_第二表",
      );
      XLSX.utils.book_append_sheet(
        filteredWorkbook,
        extractRows(targetSheet, targetOption.dataStartRow, unmatchedRows),
        "不重复行_第二表",
      );
      if (protectIds) protectLongIdentifiers(filteredWorkbook);

      const uniqueMatches = new Map<string, string>();
      for (const match of matches) {
        const key = normalize(match.value, relaxed);
        if (!uniqueMatches.has(key)) uniqueMatches.set(key, match.value);
      }

      const baseName = second.file.name.replace(/\.(xlsx?|xls)$/i, "");
      const finished: MatchResult = {
        workbook: outputWorkbook,
        filteredWorkbook,
        sourceCount,
        targetCount,
        sourceMatchRowCount: sourceMatchRows.length,
        matchCount: matches.length,
        uniqueMatchCount: uniqueMatches.size,
        uniqueUnmatchedCount: unmatchedNumbers.size,
        protectedLongIds,
        matchedNumbers: [...uniqueMatches.values()],
        unmatchedNumbers: [...unmatchedNumbers.values()],
        samples: matches.slice(0, 12).map(({ row, value }) => ({ row, value })),
        outputName: `${baseName}_比对标黄.xlsx`,
        filteredOutputName: `${baseName}_筛选完整行.xlsx`,
      };
      setResult(finished);
      setNumberView("matched");
      setCopied(false);
      setTimeout(
        () => resultRef.current?.scrollIntoView({ behavior: "smooth", block: "start" }),
        80,
      );
    } catch (caught) {
      setError(
        caught instanceof Error
          ? caught.message
          : "比对失败，请确认文件没有加密或损坏。",
      );
    } finally {
      setBusy(false);
    }
  }

  function saveWorkbook(workbook: XLSX.WorkBook, fileName: string) {
    const bytes = XLSX.write(workbook, {
      type: "array",
      bookType: "xlsx",
      cellStyles: true,
      bookSST: true,
      compression: true,
    }) as ArrayBuffer;
    const blob = new Blob([bytes], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    link.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function downloadResult() {
    if (!result) return;
    saveWorkbook(result.workbook, result.outputName);
  }

  function downloadFilteredRows() {
    if (!result) return;
    saveWorkbook(result.filteredWorkbook, result.filteredOutputName);
  }

  function reset() {
    setFirst(null);
    setSecond(null);
    setFirstSheet("");
    setSecondSheet("");
    setFirstColumn(-1);
    setSecondColumn(-1);
    setResult(null);
    setError("");
    setNumberView("matched");
    setCopied(false);
    setResetKey((current) => current + 1);
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  async function copyNumbers() {
    const field = matchBoxRef.current;
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

  const ready =
    first && second && firstColumn >= 0 && secondColumn >= 0 && !busy;

  return (
    <main className="site-shell">
      <header className="hero">
        <div className="brand-mark">表</div>
        <div>
          <p className="eyebrow">Excel 本地比对工具</p>
          <h1>号码对比，一次完成</h1>
          <p className="hero-copy">
            上传两个表格，找出重复号码，在第二个表格中标黄并导出。文件只在你的浏览器中处理。
          </p>
        </div>
        {(first || second) && (
          <button className="reset-button" type="button" onClick={reset}>
            重新开始
          </button>
        )}
      </header>

      <section className="workspace-card" aria-label="上传并比对表格">
        <div className="step-row">
          <span>第 1 步</span>
          <strong>上传两个 Excel 表格</strong>
        </div>
        <div className="upload-grid">
          <UploadCard
            key={`first-${resetKey}`}
            id="first-file"
            number="01"
            title="第一个表格"
            description="号码来源清单"
            value={first}
            onFile={(file) => loadFile("first", file)}
          />
          <UploadCard
            key={`second-${resetKey}`}
            id="second-file"
            number="02"
            title="第二个表格"
            description="需要标记并导出的表格"
            value={second}
            onFile={(file) => loadFile("second", file)}
          />
        </div>

        {first && second && (
          <div className="configuration">
            <div className="step-row compact">
              <span>第 2 步</span>
              <strong>选择需要比对的列</strong>
              <small>系统会自动跳过文字表头</small>
            </div>
            <div className="selector-grid">
              <div className="selector-card">
                <label htmlFor="first-sheet">第一个表格 · 工作表</label>
                <select
                  id="first-sheet"
                  value={firstSheet}
                  onChange={(event) => changeSheet("first", event.target.value)}
                >
                  {first.workbook.SheetNames.map((name) => (
                    <option key={name} value={name}>{name}</option>
                  ))}
                </select>
                <label htmlFor="first-column">用于查找的号码列</label>
                <select
                  id="first-column"
                  value={firstColumn}
                  onChange={(event) => {
                    setFirstColumn(Number(event.target.value));
                    setResult(null);
                  }}
                >
                  {firstColumns.map((option) => (
                    <option key={option.index} value={option.index}>
                      {option.label}
                    </option>
                  ))}
                </select>
              </div>
              <div className="direction-arrow" aria-hidden="true">→</div>
              <div className="selector-card">
                <label htmlFor="second-sheet">第二个表格 · 工作表</label>
                <select
                  id="second-sheet"
                  value={secondSheet}
                  onChange={(event) => changeSheet("second", event.target.value)}
                >
                  {second.workbook.SheetNames.map((name) => (
                    <option key={name} value={name}>{name}</option>
                  ))}
                </select>
                <label htmlFor="second-column">需要标黄的号码列</label>
                <select
                  id="second-column"
                  value={secondColumn}
                  onChange={(event) => {
                    setSecondColumn(Number(event.target.value));
                    setResult(null);
                  }}
                >
                  {secondColumns.map((option) => (
                    <option key={option.index} value={option.index}>
                      {option.label}
                    </option>
                  ))}
                </select>
              </div>
            </div>

            <div className="options-row">
              <label className="check-option">
                <input
                  type="checkbox"
                  checked={protectIds}
                  onChange={(event) => setProtectIds(event.target.checked)}
                />
                <span>
                  <b>保护长编号</b>
                  <small>15 位以上数字导出时保存为文本，防止精度变化</small>
                </span>
              </label>
              <label className="check-option">
                <input
                  type="checkbox"
                  checked={relaxed}
                  onChange={(event) => setRelaxed(event.target.checked)}
                />
                <span>
                  <b>忽略常见分隔符</b>
                  <small>比对时忽略空格、短横线和括号</small>
                </span>
              </label>
              <label className="check-option">
                <input
                  type="checkbox"
                  checked={showMatchBox}
                  onChange={(event) => setShowMatchBox(event.target.checked)}
                />
                <span>
                  <b>显示号码文本框</b>
                  <small>可切换相同或不重复号码，并编辑、复制</small>
                </span>
              </label>
            </div>
          </div>
        )}

        {error && <div className="error-banner" role="alert">{error}</div>}

        <button
          className="compare-button"
          type="button"
          disabled={!ready}
          onClick={compare}
        >
          {busy
            ? "正在读取并比对，请稍候…"
            : ready
              ? "开始比对并标黄"
              : "上传两个文件后开始比对"}
        </button>
      </section>

      {result && (
        <section className="result-card" ref={resultRef} aria-live="polite">
          <div className="result-heading">
            <div>
              <p className="eyebrow">比对完成</p>
              <h2>已找到 {result.matchCount.toLocaleString()} 条相同号码</h2>
              <p>对应单元格已在第二个表格中填充为黄色。</p>
            </div>
            <div className="success-badge">完成</div>
          </div>

          <div className="stat-grid">
            <div><span>来源号码</span><strong>{result.sourceCount.toLocaleString()}</strong></div>
            <div><span>第二表号码</span><strong>{result.targetCount.toLocaleString()}</strong></div>
            <div><span>第一表相同行</span><strong>{result.sourceMatchRowCount.toLocaleString()}</strong></div>
            <div className="accent-stat"><span>匹配行数</span><strong>{result.matchCount.toLocaleString()}</strong></div>
            <div><span>唯一匹配号码</span><strong>{result.uniqueMatchCount.toLocaleString()}</strong></div>
            <div><span>不重复号码</span><strong>{result.uniqueUnmatchedCount.toLocaleString()}</strong></div>
          </div>

          {result.samples.length > 0 ? (
            <div className="sample-panel">
              <div className="sample-header">
                <strong>匹配预览</strong>
                <span>最多显示前 12 条</span>
              </div>
              <div className="sample-table" role="table">
                <div className="sample-row table-head" role="row">
                  <span>第二表行号</span>
                  <span>相同号码</span>
                  <span>状态</span>
                </div>
                {result.samples.map((sample) => (
                  <div className="sample-row" role="row" key={`${sample.row}-${sample.value}`}>
                    <span>第 {sample.row} 行</span>
                    <b>{sample.value}</b>
                    <em>已标黄</em>
                  </div>
                ))}
              </div>
            </div>
          ) : (
            <div className="empty-result">没有找到相同号码，仍可更换列或文件后再次比对。</div>
          )}

          {showMatchBox && (
            <div className="match-box-panel">
              <div className="match-box-heading">
                <div>
                  <strong>号码获取框</strong>
                  <span>
                    {numberView === "matched"
                      ? `每行一个，共 ${result.matchedNumbers.length.toLocaleString()} 个相同号码`
                      : `每行一个，共 ${result.unmatchedNumbers.length.toLocaleString()} 个不重复号码`}
                  </span>
                </div>
                <button type="button" onClick={copyNumbers}>
                  {copied ? "已复制" : "复制全部"}
                </button>
              </div>
              <div className="number-view-switch" aria-label="选择号码类型">
                <button
                  type="button"
                  className={numberView === "matched" ? "is-active" : ""}
                  aria-pressed={numberView === "matched"}
                  onClick={() => {
                    setNumberView("matched");
                    setCopied(false);
                  }}
                >
                  相同号码 · {result.matchedNumbers.length.toLocaleString()}
                </button>
                <button
                  type="button"
                  className={numberView === "unmatched" ? "is-active" : ""}
                  aria-pressed={numberView === "unmatched"}
                  onClick={() => {
                    setNumberView("unmatched");
                    setCopied(false);
                  }}
                >
                  一键获取不重复号码 · {result.unmatchedNumbers.length.toLocaleString()}
                </button>
              </div>
              <textarea
                key={`${result.outputName}-${numberView}`}
                ref={matchBoxRef}
                defaultValue={
                  numberView === "matched"
                    ? result.matchedNumbers.join("\n")
                    : result.unmatchedNumbers.join("\n")
                }
                placeholder={
                  numberView === "matched"
                    ? "没有相同号码"
                    : "没有不重复号码"
                }
                aria-label={
                  numberView === "matched"
                    ? "相同号码，可选择和编辑"
                    : "不重复号码，可选择和编辑"
                }
                spellCheck={false}
              />
            </div>
          )}

          {protectIds && (
            <p className="protection-note">
              已保护 {result.protectedLongIds.toLocaleString()} 个长编号，导出时不会被转成科学计数法。
            </p>
          )}
          <p className="filter-note">
            “筛选完整行”文件包含：第一表相同行、第二表相同行、第二表不重复行，原行的所有列都会保留。
          </p>
          <div className="result-actions">
            <button className="download-button" type="button" onClick={downloadResult}>
              下载标黄后的 Excel
            </button>
            <button
              className="filter-download-button"
              type="button"
              onClick={downloadFilteredRows}
            >
              下载筛选完整行
            </button>
            <button className="secondary-button" type="button" onClick={reset}>
              比对其他文件
            </button>
          </div>
        </section>
      )}

      <footer className="privacy-note">
        <span className="status-dot" />
        本地处理 · 不保存表格 · 支持 .xlsx / .xls · 导出为 .xlsx
      </footer>
    </main>
  );
}
