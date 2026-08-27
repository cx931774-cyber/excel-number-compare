"use client";

import { useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx-js-style";
import TextExcelTool from "./text-excel-tool";

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

type MatchMode = "all" | "any";

type MatchCondition = {
  id: number;
  firstColumn: number;
  secondColumn: number;
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
  const [activeTool, setActiveTool] = useState<"compare" | "text">("compare");
  const [first, setFirst] = useState<WorkbookFile | null>(null);
  const [second, setSecond] = useState<WorkbookFile | null>(null);
  const [firstSheet, setFirstSheet] = useState("");
  const [secondSheet, setSecondSheet] = useState("");
  const [conditions, setConditions] = useState<MatchCondition[]>([
    { id: 1, firstColumn: -1, secondColumn: -1 },
  ]);
  const [matchMode, setMatchMode] = useState<MatchMode>("all");
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
  const nextConditionId = useRef(2);
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
        setConditions((current) => [
          {
            id: nextConditionId.current++,
            firstColumn: preferredColumn(options, "source"),
            secondColumn: current[0]?.secondColumn ?? -1,
          },
        ]);
      } else {
        setSecond(loaded);
        setSecondSheet(sheet);
        setConditions((current) => [
          {
            id: nextConditionId.current++,
            firstColumn: current[0]?.firstColumn ?? -1,
            secondColumn: preferredColumn(options, "target"),
          },
        ]);
      }
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "无法读取这个文件。");
    }
  }

  function changeSheet(position: "first" | "second", sheetName: string) {
    setResult(null);
    if (position === "first" && first) {
      setFirstSheet(sheetName);
      setConditions((current) => [
        {
          id: nextConditionId.current++,
          firstColumn: preferredColumn(
            worksheetColumns(first.workbook, sheetName),
            "source",
          ),
          secondColumn: current[0]?.secondColumn ?? -1,
        },
      ]);
    }
    if (position === "second" && second) {
      setSecondSheet(sheetName);
      setConditions((current) => [
        {
          id: nextConditionId.current++,
          firstColumn: current[0]?.firstColumn ?? -1,
          secondColumn: preferredColumn(
            worksheetColumns(second.workbook, sheetName),
            "target",
          ),
        },
      ]);
    }
  }

  function updateCondition(
    id: number,
    field: "firstColumn" | "secondColumn",
    value: number,
  ) {
    setConditions((current) =>
      current.map((condition) =>
        condition.id === id ? { ...condition, [field]: value } : condition,
      ),
    );
    setResult(null);
  }

  function addCondition() {
    const usedFirst = new Set(conditions.map(({ firstColumn }) => firstColumn));
    const usedSecond = new Set(conditions.map(({ secondColumn }) => secondColumn));
    const firstOption =
      firstColumns.find(({ index }) => !usedFirst.has(index)) ?? firstColumns[0];
    const secondOption =
      secondColumns.find(({ index }) => !usedSecond.has(index)) ?? secondColumns[0];
    if (!firstOption || !secondOption) return;
    setConditions((current) => [
      ...current,
      {
        id: nextConditionId.current++,
        firstColumn: firstOption.index,
        secondColumn: secondOption.index,
      },
    ]);
    setResult(null);
  }

  function removeCondition(id: number) {
    setConditions((current) => current.filter((condition) => condition.id !== id));
    setResult(null);
  }

  async function compare() {
    if (!ready || !first || !second) return;
    setBusy(true);
    setError("");
    setResult(null);
    await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));

    try {
      const sourceSheet = first.workbook.Sheets[firstSheet];
      const activeConditions = conditions.map((condition) => ({
        ...condition,
        sourceOption: firstColumns.find(
          (option) => option.index === condition.firstColumn,
        ),
        targetOption: secondColumns.find(
          (option) => option.index === condition.secondColumn,
        ),
      }));
      if (
        !sourceSheet?.["!ref"] ||
        activeConditions.some(
          ({ sourceOption, targetOption }) => !sourceOption || !targetOption,
        )
      ) {
        throw new Error("请选择有效的工作表和比对列。");
      }

      const resolvedConditions = activeConditions as Array<
        MatchCondition & {
          sourceOption: ColumnOption;
          targetOption: ColumnOption;
        }
      >;

      const outputWorkbook = XLSX.read(second.buffer.slice(0), {
        type: "array",
        cellStyles: true,
        cellNF: true,
        cellDates: false,
      });
      const targetSheet = outputWorkbook.Sheets[secondSheet];
      if (!targetSheet?.["!ref"]) throw new Error("第二个工作表没有可读取的数据。");

      const sourceRange = XLSX.utils.decode_range(sourceSheet["!ref"]);
      const sourceStartRow = Math.max(
        ...resolvedConditions.map(({ sourceOption }) => sourceOption.dataStartRow),
      );
      const targetStartRow = Math.max(
        ...resolvedConditions.map(({ targetOption }) => targetOption.dataStartRow),
      );
      const sourceRowsByComposite = new Map<string, number[]>();
      const sourceRowsByCondition = resolvedConditions.map(
        () => new Map<string, number[]>(),
      );
      let sourceCount = 0;
      for (let row = sourceStartRow; row <= sourceRange.e.r; row += 1) {
        const values = resolvedConditions.map(({ firstColumn }) =>
          normalize(
            cellText(
              sourceSheet[XLSX.utils.encode_cell({ r: row, c: firstColumn })],
            ),
            relaxed,
          ),
        );
        const hasUsableValues =
          matchMode === "all" ? values.every(Boolean) : values.some(Boolean);
        if (!hasUsableValues) continue;
        sourceCount += 1;

        if (matchMode === "all") {
          const key = JSON.stringify(values);
          const rows = sourceRowsByComposite.get(key) ?? [];
          rows.push(row);
          sourceRowsByComposite.set(key, rows);
        } else {
          values.forEach((value, index) => {
            if (!value) return;
            const rows = sourceRowsByCondition[index].get(value) ?? [];
            rows.push(row);
            sourceRowsByCondition[index].set(value, rows);
          });
        }
      }

      const targetRange = XLSX.utils.decode_range(targetSheet["!ref"]);
      const matches: Array<{
        row: number;
        worksheetRow: number;
        value: string;
        addresses: string[];
        sourceRows: number[];
      }> = [];
      const unmatchedNumbers = new Map<string, string>();
      const unmatchedRows: number[] = [];
      let targetCount = 0;
      for (let row = targetStartRow; row <= targetRange.e.r; row += 1) {
        const displayValues = resolvedConditions.map(({ secondColumn }) =>
          cellText(targetSheet[XLSX.utils.encode_cell({ r: row, c: secondColumn })]),
        );
        const values = displayValues.map((value) => normalize(value, relaxed));
        if (!values.some(Boolean)) continue;
        targetCount += 1;
        const key = JSON.stringify(values);
        const summary =
          displayValues.length === 1
            ? displayValues[0]
            : displayValues
                .map(
                  (value, index) =>
                    `${resolvedConditions[index].targetOption.letter}=${value || "（空）"}`,
                )
                .join(" | ");

        let matchingIndexes: number[] = [];
        let matchingSourceRows: number[] = [];
        if (matchMode === "all" && values.every(Boolean)) {
          matchingSourceRows = sourceRowsByComposite.get(key) ?? [];
          if (matchingSourceRows.length) {
            matchingIndexes = values.map((_, index) => index);
          }
        } else if (matchMode === "any") {
          matchingIndexes = values.flatMap((value, index) =>
            value && sourceRowsByCondition[index].has(value) ? [index] : [],
          );
          matchingSourceRows = [
            ...new Set(
              matchingIndexes.flatMap(
                (index) => sourceRowsByCondition[index].get(values[index]) ?? [],
              ),
            ),
          ];
        }

        if (matchingIndexes.length) {
          matches.push({
            row: row + 1,
            worksheetRow: row,
            value: summary,
            addresses: matchingIndexes.map((index) =>
              XLSX.utils.encode_cell({
                r: row,
                c: resolvedConditions[index].secondColumn,
              }),
            ),
            sourceRows: matchingSourceRows,
          });
        } else {
          if (!unmatchedNumbers.has(key)) unmatchedNumbers.set(key, summary);
          unmatchedRows.push(row);
        }
      }

      const protectedLongIds = protectIds
        ? protectLongIdentifiers(outputWorkbook)
        : 0;
      for (const match of matches) {
        for (const address of match.addresses) {
          const cell = targetSheet[address] as StyledCell;
          cell.s = {
            ...(cell.s ?? {}),
            fill: {
              patternType: "solid",
              fgColor: { rgb: "FFFF00" },
              bgColor: { rgb: "FFFF00" },
            },
          };
        }
      }

      const sourceMatchRows = [
        ...new Set(matches.flatMap((match) => match.sourceRows)),
      ].sort((left, right) => left - right);
      const filteredWorkbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(
        filteredWorkbook,
        extractRows(sourceSheet, sourceStartRow, sourceMatchRows),
        "相同行_第一表",
      );
      XLSX.utils.book_append_sheet(
        filteredWorkbook,
        extractRows(
          targetSheet,
          targetStartRow,
          matches.map((match) => match.worksheetRow),
        ),
        "相同行_第二表",
      );
      XLSX.utils.book_append_sheet(
        filteredWorkbook,
        extractRows(targetSheet, targetStartRow, unmatchedRows),
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
    setConditions([
      { id: nextConditionId.current++, firstColumn: -1, secondColumn: -1 },
    ]);
    setMatchMode("all");
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

  const ready = Boolean(
    first &&
      second &&
      conditions.length > 0 &&
      conditions.every(
        ({ firstColumn, secondColumn }) => firstColumn >= 0 && secondColumn >= 0,
      ) &&
      !busy,
  );

  return (
    <main className="site-shell">
      <header className="hero">
        <div className="brand-mark">表</div>
        <div>
          <p className="eyebrow">Excel 本地处理工具</p>
          <h1>
            {activeTool === "compare" ? "号码对比，一次完成" : "杂乱文本，整理成表"}
          </h1>
          <p className="hero-copy">
            {activeTool === "compare"
              ? "上传两个表格，按一个或多个自定义条件找出相同行，在第二个表格中标黄并导出。"
              : "输入空格分隔的表头，粘贴文本，系统自动识别字段并生成可下载的 Excel。"}
            文件只在你的浏览器中处理。
          </p>
        </div>
        {activeTool === "compare" && (first || second) && (
          <button className="reset-button" type="button" onClick={reset}>
            重新开始
          </button>
        )}
      </header>

      <div className="tool-switch" role="tablist" aria-label="选择处理方式">
        <button
          type="button"
          role="tab"
          aria-selected={activeTool === "compare"}
          className={activeTool === "compare" ? "is-active" : ""}
          onClick={() => setActiveTool("compare")}
        >
          <span>01</span>
          <b>Excel 表格比对</b>
          <small>两个文件查相同与不重复</small>
        </button>
        <button
          type="button"
          role="tab"
          aria-selected={activeTool === "text"}
          className={activeTool === "text" ? "is-active" : ""}
          onClick={() => setActiveTool("text")}
        >
          <span>02</span>
          <b>文本整理成 Excel</b>
          <small>粘贴文本自动分列导出</small>
        </button>
      </div>

      <div hidden={activeTool !== "compare"}>

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
              <strong>设置自定义比对条件</strong>
              <small>可添加多组列，并选择全部满足或任一满足</small>
            </div>
            <div className="sheet-selector-grid">
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
              </div>
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
              </div>
            </div>

            <div className="condition-builder">
              <div className="condition-toolbar">
                <div>
                  <strong>自定义条件</strong>
                  <span>每组条件把第一表的一列对应到第二表的一列</span>
                </div>
                <div className="match-mode-switch" role="group" aria-label="条件组合方式">
                  <button
                    type="button"
                    className={matchMode === "all" ? "is-active" : ""}
                    aria-pressed={matchMode === "all"}
                    onClick={() => {
                      setMatchMode("all");
                      setResult(null);
                    }}
                  >
                    全部条件相同
                  </button>
                  <button
                    type="button"
                    className={matchMode === "any" ? "is-active" : ""}
                    aria-pressed={matchMode === "any"}
                    onClick={() => {
                      setMatchMode("any");
                      setResult(null);
                    }}
                  >
                    任一条件相同
                  </button>
                </div>
              </div>

              <div className="condition-list">
                {conditions.map((condition, index) => (
                  <div className="condition-row" key={condition.id}>
                    <span className="condition-number">条件 {index + 1}</span>
                    <label htmlFor={`first-column-${condition.id}`}>
                      <span>第一表列</span>
                      <select
                        id={`first-column-${condition.id}`}
                        value={condition.firstColumn}
                        onChange={(event) =>
                          updateCondition(
                            condition.id,
                            "firstColumn",
                            Number(event.target.value),
                          )
                        }
                      >
                        {firstColumns.map((option) => (
                          <option key={option.index} value={option.index}>
                            {option.label}
                          </option>
                        ))}
                      </select>
                    </label>
                    <span className="condition-arrow" aria-hidden="true">→</span>
                    <label htmlFor={`second-column-${condition.id}`}>
                      <span>第二表列</span>
                      <select
                        id={`second-column-${condition.id}`}
                        value={condition.secondColumn}
                        onChange={(event) =>
                          updateCondition(
                            condition.id,
                            "secondColumn",
                            Number(event.target.value),
                          )
                        }
                      >
                        {secondColumns.map((option) => (
                          <option key={option.index} value={option.index}>
                            {option.label}
                          </option>
                        ))}
                      </select>
                    </label>
                    {conditions.length > 1 && (
                      <button
                        className="remove-condition"
                        type="button"
                        aria-label={`删除条件 ${index + 1}`}
                        onClick={() => removeCondition(condition.id)}
                      >
                        删除
                      </button>
                    )}
                  </div>
                ))}
              </div>

              <button
                className="add-condition"
                type="button"
                disabled={conditions.length >= 8}
                onClick={addCondition}
              >
                {conditions.length >= 8 ? "最多添加 8 个条件" : "+ 添加比对条件"}
              </button>
            </div>

            <div className="options-row">
              <label className="check-option" aria-label="保护长编号">
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
              <label className="check-option" aria-label="忽略常见分隔符">
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
              <label className="check-option" aria-label="显示号码文本框">
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
            ? "正在按自定义条件比对，请稍候…"
            : ready
              ? `按 ${conditions.length} 个条件开始比对并标黄`
              : "上传两个文件后开始比对"}
        </button>
      </section>

      {result && (
        <section className="result-card" ref={resultRef} aria-live="polite">
          <div className="result-heading">
            <div>
              <p className="eyebrow">比对完成</p>
              <h2>已找到 {result.matchCount.toLocaleString()} 条相同行</h2>
              <p>
                按“{matchMode === "all" ? "全部条件相同" : "任一条件相同"}”完成比对，对应单元格已标黄。
              </p>
            </div>
            <div className="success-badge">完成</div>
          </div>

          <div className="stat-grid">
            <div><span>第一表有效行</span><strong>{result.sourceCount.toLocaleString()}</strong></div>
            <div><span>第二表有效行</span><strong>{result.targetCount.toLocaleString()}</strong></div>
            <div><span>第一表相同行</span><strong>{result.sourceMatchRowCount.toLocaleString()}</strong></div>
            <div className="accent-stat"><span>匹配行数</span><strong>{result.matchCount.toLocaleString()}</strong></div>
            <div><span>唯一匹配内容</span><strong>{result.uniqueMatchCount.toLocaleString()}</strong></div>
            <div><span>不重复内容</span><strong>{result.uniqueUnmatchedCount.toLocaleString()}</strong></div>
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
                  <span>匹配的条件内容</span>
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
            <div className="empty-result">没有找到符合条件的相同行，可更换条件、列或文件后再次比对。</div>
          )}

          {showMatchBox && (
            <div className="match-box-panel">
              <div className="match-box-heading">
                <div>
                  <strong>匹配内容获取框</strong>
                  <span>
                    {numberView === "matched"
                      ? `每行一条，共 ${result.matchedNumbers.length.toLocaleString()} 条匹配内容`
                      : `每行一条，共 ${result.unmatchedNumbers.length.toLocaleString()} 条不重复内容`}
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
                  相同内容 · {result.matchedNumbers.length.toLocaleString()}
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
                  一键获取不重复号码/内容 · {result.unmatchedNumbers.length.toLocaleString()}
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
                    ? "没有相同内容"
                    : "没有不重复内容"
                }
                aria-label={
                  numberView === "matched"
                    ? "相同内容，可选择和编辑"
                    : "不重复内容，可选择和编辑"
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
      </div>

      <div hidden={activeTool !== "text"}>
        <TextExcelTool />
      </div>

      <footer className="privacy-note">
        <span className="status-dot" />
        本地处理 · 不保存表格 · 支持 .xlsx / .xls · 导出为 .xlsx
      </footer>
    </main>
  );
}
