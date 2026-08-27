export type TextExcelRow = Record<string, string>;

type HeaderKind = "phone" | "operator" | "amount" | "name" | "balance" | "generic";

function headerKind(header: string): HeaderKind {
  if (/余额|结余|剩余/i.test(header)) return "balance";
  if (/运营商|运营|网络|归属/i.test(header)) return "operator";
  if (/手机号|手机|电话号码|电话|号码|phone/i.test(header)) return "phone";
  if (/充值|金额|缴费|money/i.test(header)) return "amount";
  if (/姓名|名字|联系人|用户|name/i.test(header)) return "name";
  return "generic";
}

function cleanToken(token: string) {
  return token.trim().replace(/^[,，;；]+|[,，;；]+$/g, "");
}

function cleanNumber(value: string) {
  const match = value.replace(/,/g, "").match(/[+-]?\d+(?:\.\d+)?/);
  return match?.[0] ?? cleanToken(value);
}

function findHeader(
  headers: string[],
  kind: HeaderKind,
  preferredLabel?: string,
) {
  const direct = preferredLabel
    ? headers.find((header) => header === preferredLabel)
    : undefined;
  return direct ?? headers.find((header) => headerKind(header) === kind);
}

export function parseTextHeaders(input: string) {
  return [...new Set(input.trim().split(/\s+/).filter(Boolean))];
}

export function parseTextLine(headers: string[], line: string): TextExcelRow {
  const row = Object.fromEntries(headers.map((header) => [header, ""]));
  const tokens = line.trim().split(/\s+/).map(cleanToken).filter(Boolean);
  const unused = new Set(tokens.map((_, index) => index));

  const assign = (header: string | undefined, value: string, tokenIndex: number) => {
    if (!header || row[header]) return false;
    row[header] = value;
    unused.delete(tokenIndex);
    return true;
  };

  tokens.forEach((token, index) => {
    const labeled = token.match(/^([^:：]+)[:：](.+)$/);
    if (!labeled) return;
    const label = labeled[1].trim();
    const rawValue = labeled[2].trim();
    const kind = headerKind(label);
    const header = findHeader(headers, kind, label);
    const value = ["phone", "amount", "balance"].includes(kind)
      ? cleanNumber(rawValue)
      : rawValue;
    assign(header, value, index);
  });

  tokens.forEach((token, index) => {
    if (!unused.has(index)) return;
    const normalizedPhone = token.replace(/\D/g, "");
    if (/^1[3-9]\d{9}$/.test(normalizedPhone)) {
      assign(findHeader(headers, "phone"), normalizedPhone, index);
    }
  });

  tokens.forEach((token, index) => {
    if (!unused.has(index)) return;
    if (/(?:中国)?(?:移动|联通|电信|广电)/.test(token)) {
      assign(findHeader(headers, "operator"), token, index);
    }
  });

  tokens.forEach((token, index) => {
    if (!unused.has(index)) return;
    if (/^(?:[¥￥])?[+-]?\d+(?:\.\d+)?(?:元)?$/.test(token.replace(/,/g, ""))) {
      assign(findHeader(headers, "amount"), cleanNumber(token), index);
    }
  });

  tokens.forEach((token, index) => {
    if (!unused.has(index)) return;
    if (/^[\u3400-\u9fff·]{2,8}$/.test(token) && !/(移动|联通|电信|广电)/.test(token)) {
      assign(findHeader(headers, "name"), token, index);
    }
  });

  const remainingHeaders = headers.filter((header) => !row[header]);
  const remainingTokens = [...unused].map((index) => tokens[index]);
  remainingHeaders.forEach((header, index) => {
    if (remainingTokens[index]) row[header] = remainingTokens[index];
  });

  return row;
}

export function parseTextRows(headers: string[], text: string) {
  return text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => parseTextLine(headers, line));
}
