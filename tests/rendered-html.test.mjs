import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

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
  assert.match(html, /<title>表格号码比对<\/title>/i);
  assert.match(html, /号码对比，一次完成/);
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
