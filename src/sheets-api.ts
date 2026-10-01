import { ErrorCode, McpError } from "@modelcontextprotocol/sdk/types.js";
import { Env } from "./types";
import { getAccessToken } from "./auth";

// Google Sheets API V4 Helper functions

export const MAX_WRITE_ROWS = 1_000;
export const MAX_WRITE_CELLS = 10_000;
export const MAX_WRITE_BODY_BYTES = 1_000_000;
const MAX_RANGE_LENGTH = 200;
const MAX_TITLE_LENGTH = 255;

// ── Validation ────────────────────────────────────────────────────────────────

const SPREADSHEET_ID_RE = /^[A-Za-z0-9_-]{20,128}$/;
const SPREADSHEET_URL_RE = /^https:\/\/docs\.google\.com\/spreadsheets\/d\/([A-Za-z0-9_-]+)(?:[/?#].*)?$/;

/** Accepts a raw spreadsheet ID or a Google Sheets URL and returns the validated ID. */
export function extractSpreadsheetId(input: unknown): string {
  if (typeof input !== "string") {
    throw new McpError(ErrorCode.InvalidParams, "Invalid spreadsheet ID format");
  }
  const trimmed = input.trim();
  const id = SPREADSHEET_URL_RE.exec(trimmed)?.[1] ?? trimmed;
  if (!SPREADSHEET_ID_RE.test(id)) {
    throw new McpError(ErrorCode.InvalidParams, "Invalid spreadsheet ID format");
  }
  return id;
}

// A1 notation: optional sheet prefix, then CELL, CELL:CELL, COL:COL or ROW:ROW.
// A bare sheet name (no "!") is allowed so append can target a whole sheet.
const CELL = "[A-Za-z]{1,3}\\d{0,7}";
const SHEET = "(?:'(?:[^']|'')+'|[A-Za-z0-9_ .-]+)";
const A1_RE = new RegExp(`^(?:${SHEET}!)?(?:${CELL}(?::${CELL})?|\\d{1,7}:\\d{1,7})$|^${SHEET}$`);

export function validateRange(range: unknown): string {
  if (typeof range !== "string") {
    throw new McpError(ErrorCode.InvalidParams, "Invalid A1 range");
  }
  const trimmed = range.trim();
  if (trimmed === "" || trimmed.length > MAX_RANGE_LENGTH || !A1_RE.test(trimmed)) {
    throw new McpError(ErrorCode.InvalidParams, "Invalid A1 range");
  }
  return trimmed;
}

export function validateValues(values: unknown): (string | number | boolean)[][] {
  if (!Array.isArray(values) || values.length === 0) {
    throw new McpError(ErrorCode.InvalidParams, '"values" must be a non-empty 2D array');
  }
  if (values.length > MAX_WRITE_ROWS) {
    throw new McpError(ErrorCode.InvalidParams, `"values" exceeds maximum of ${MAX_WRITE_ROWS} rows`);
  }
  let cells = 0;
  for (const row of values) {
    if (!Array.isArray(row)) {
      throw new McpError(ErrorCode.InvalidParams, '"values" must be a 2D array');
    }
    cells += row.length;
    if (cells > MAX_WRITE_CELLS) {
      throw new McpError(ErrorCode.InvalidParams, `"values" exceeds maximum of ${MAX_WRITE_CELLS} cells`);
    }
    for (const cell of row) {
      if (typeof cell !== "string" && typeof cell !== "number" && typeof cell !== "boolean") {
        throw new McpError(ErrorCode.InvalidParams, '"values" cells must be string, number or boolean');
      }
    }
  }
  return values as (string | number | boolean)[][];
}

function buildWriteBody(values: unknown): string {
  const body = JSON.stringify({ values: validateValues(values) });
  if (new TextEncoder().encode(body).length > MAX_WRITE_BODY_BYTES) {
    throw new McpError(ErrorCode.InvalidParams, `"values" exceeds maximum payload size of ${MAX_WRITE_BODY_BYTES} bytes`);
  }
  return body;
}

// ── Transport ─────────────────────────────────────────────────────────────────

async function authorizedAccessToken(userToken: string, env: Env): Promise<string> {
  const accessToken = await getAccessToken(userToken, env);
  if (!accessToken) throw new McpError(ErrorCode.InvalidRequest, "Unauthorized");
  return accessToken;
}

/** Fetches a Google API URL and returns parsed JSON; non-OK responses throw a minimal error (no body). */
async function googleJson(url: string, init: RequestInit): Promise<unknown> {
  const response = await fetch(url, init);
  if (!response.ok) {
    throw new McpError(ErrorCode.InternalError, `Google API request failed (HTTP ${response.status})`);
  }
  return await response.json();
}

// ── API ───────────────────────────────────────────────────────────────────────

export async function listSpreadsheets(userToken: string, env: Env) {
  const accessToken = await authorizedAccessToken(userToken, env);
  const q = encodeURIComponent("mimeType='application/vnd.google-apps.spreadsheet'");
  return googleJson(
    `https://www.googleapis.com/drive/v3/files?q=${q}&fields=${encodeURIComponent("files(id,name)")}&pageSize=20`,
    { headers: { Authorization: `Bearer ${accessToken}` } }
  );
}

export async function readSheetRange(userToken: string, env: Env, spreadsheetId: string, range: string) {
  const id = extractSpreadsheetId(spreadsheetId);
  const a1 = validateRange(range);
  const accessToken = await authorizedAccessToken(userToken, env);
  return googleJson(
    `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(id)}/values/${encodeURIComponent(a1)}`,
    { headers: { Authorization: `Bearer ${accessToken}` } }
  );
}

export async function createSpreadsheet(userToken: string, env: Env, title: string) {
  if (typeof title !== "string" || title.trim() === "" || title.length > MAX_TITLE_LENGTH) {
    throw new McpError(ErrorCode.InvalidParams, `"title" must be a non-empty string up to ${MAX_TITLE_LENGTH} characters`);
  }
  const accessToken = await authorizedAccessToken(userToken, env);
  return googleJson("https://sheets.googleapis.com/v4/spreadsheets", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ properties: { title } }),
  });
}

export async function writeSheetRange(userToken: string, env: Env, spreadsheetId: string, range: string, values: unknown) {
  const id = extractSpreadsheetId(spreadsheetId);
  const a1 = validateRange(range);
  const body = buildWriteBody(values);
  const accessToken = await authorizedAccessToken(userToken, env);
  return googleJson(
    `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(id)}/values/${encodeURIComponent(a1)}?valueInputOption=RAW`,
    {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body,
    }
  );
}

export async function appendSheetRows(userToken: string, env: Env, spreadsheetId: string, range: string, values: unknown) {
  const id = extractSpreadsheetId(spreadsheetId);
  const a1 = validateRange(range);
  const body = buildWriteBody(values);
  const accessToken = await authorizedAccessToken(userToken, env);
  return googleJson(
    `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(id)}/values/${encodeURIComponent(a1)}:append?valueInputOption=RAW`,
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body,
    }
  );
}
