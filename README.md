# Excel 表格号码比对

在浏览器本地上传两个 Excel 表格，可完成：

- 查找两个表格中的相同号码。
- 将第二个表格的相同号码标黄并导出。
- 一键获取第二表中未出现在第一表的不重复号码。
- 导出第一表相同行、第二表相同行和第二表不重复行，并保留整行数据。
- 保护 15 位以上长编号，避免导出时变成科学计数法。

所有表格都只在浏览器中处理，不会上传到服务器。

## 在线使用

[GitHub Pages 网站](https://cx931774-cyber.github.io/excel-number-compare/)

## 本地运行

需要 Node.js 22 或更高版本。

```bash
npm install
npm run dev
```

首页代码位于 `app/page.tsx`，页面样式位于 `app/globals.css`。

## GitHub Pages

推送到 `main` 分支后，`.github/workflows/pages.yml` 会自动构建并发布网站。

```bash
npm run build:github
```
