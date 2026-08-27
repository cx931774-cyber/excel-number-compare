import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import {
  parseTextHeaders,
  parseTextRows,
} from "../app/text-processing.ts";

async function render() {
  const workerUrl = new URL("../dist/server/index.js", import.meta.url);
  workerUrl.searchParams.set("test", `${process.pid}-${Date.now()}`);
  const { default: worker } = await import(workerUrl.href);

  return worker.fetch(
    new Request("http://localhost/", {
      headers: { accept: "text/html" },
    }),
    {
      ASSETS: {
        fetch: async () => new Response("Not found", { status: 404 }),
      },
    },
    {
      waitUntil() {},
      passThroughOnException() {},
    },
  );
}

test("server-renders the Excel comparison site", async () => {
  const response = await render();
  assert.equal(response.status, 200);
  assert.match(response.headers.get("content-type") ?? "", /^text\/html\b/i);

  const html = await response.text();
  assert.match(html, /<title>Excel 表格比对与文本整理工具<\/title>/i);
  assert.match(html, /表格比对/);
  assert.match(html, /文本转 Excel/);
  assert.match(html, /第一个表格/);
  assert.match(html, /第二个表格/);
  assert.match(html, /上传两个文件后开始比对/);
  assert.doesNotMatch(html, /codex-preview|Your site is taking shape/i);
});

test("includes unmatched-number and full-row export features", async () => {
  const page = await readFile(
    new URL("../app/page.tsx", import.meta.url),
    "utf8",
  );

  assert.match(page, /一键获取不重复号码/);
  assert.match(page, /unmatchedNumbers/);
  assert.match(page, /相同行_第一表/);
  assert.match(page, /相同行_第二表/);
  assert.match(page, /不重复行_第二表/);
  assert.match(page, /下载筛选完整行/);
  assert.match(page, /URL\.revokeObjectURL/);
});

test("supports multiple custom comparison conditions", async () => {
  const page = await readFile(
    new URL("../app/page.tsx", import.meta.url),
    "utf8",
  );

  assert.match(page, /添加比对条件/);
  assert.match(page, /全部条件相同/);
  assert.match(page, /任一条件相同/);
  assert.match(page, /sourceRowsByComposite/);
  assert.match(page, /sourceRowsByCondition/);
  assert.match(page, /conditions\.length >= 8/);
});

test("classifies free-form text into the requested Excel columns", () => {
  const headers = parseTextHeaders("手机号 运营商 充值金额 姓名 余额");
  const rows = parseTextRows(
    headers,
    `13450438325 冯林林 300 移动 余额：192.49
18021306062  俞淑钧  江苏电信 300 余额：112.03
18112613136  金宝 江苏电信 300 余额：100.16`,
  );

  assert.deepEqual(rows, [
    {
      手机号: "13450438325",
      运营商: "移动",
      充值金额: "300",
      姓名: "冯林林",
      余额: "192.49",
    },
    {
      手机号: "18021306062",
      运营商: "江苏电信",
      充值金额: "300",
      姓名: "俞淑钧",
      余额: "112.03",
    },
    {
      手机号: "18112613136",
      运营商: "江苏电信",
      充值金额: "300",
      姓名: "金宝",
      余额: "100.16",
    },
  ]);
});

test("text results stay editable, copyable, and downloadable", async () => {
  const tool = await readFile(
    new URL("../app/text-excel-tool.tsx", import.meta.url),
    "utf8",
  );

  assert.match(tool, /处理结果/);
  assert.match(tool, /复制全部/);
  assert.match(tool, /result-editor/);
  assert.match(tool, /textToTable\(outputText\)/);
  assert.match(tool, /下载 Excel/);
  assert.match(tool, /URL\.revokeObjectURL/);
});
