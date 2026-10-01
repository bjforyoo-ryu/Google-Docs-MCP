import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";

vi.mock("../auth", () => ({
  getAccessToken: vi.fn(async () => "test-access-token"),
}));

import {
  extractSpreadsheetId,
  validateRange,
  validateValues,
  readSheetRange,
  writeSheetRange,
  appendSheetRows,
  MAX_WRITE_ROWS,
  MAX_WRITE_CELLS,
  MAX_WRITE_BODY_BYTES,
} from "../sheets-api";
import type { Env } from "../types";

const VALID_ID = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms";
const env = {} as Env;

describe("extractSpreadsheetId", () => {
  it("returns a raw valid ID", () => {
    expect(extractSpreadsheetId(VALID_ID)).toBe(VALID_ID);
  });

  it("extracts the ID from a Google Sheets URL", () => {
    expect(extractSpreadsheetId(`https://docs.google.com/spreadsheets/d/${VALID_ID}/edit#gid=0`)).toBe(VALID_ID);
  });

  it.each([
    ["empty string", ""],
    ["too short", "abc123"],
    ["path traversal", "../../drive/v3/files"],
    ["query injection", `${VALID_ID}?x=1`],
    ["slash in ID", `${VALID_ID}/values`],
    ["non-Sheets URL", `https://evil.example.com/spreadsheets/d/${VALID_ID}`],
  ])("rejects invalid spreadsheetId: %s", (_label, input) => {
    expect(() => extractSpreadsheetId(input)).toThrow("Invalid spreadsheet ID format");
  });

  it("rejects non-string input", () => {
    expect(() => extractSpreadsheetId(123 as unknown)).toThrow("Invalid spreadsheet ID format");
  });
});

describe("validateRange", () => {
  it.each(["A1", "A1:B10", "Sheet1!A1:B10", "'My Sheet'!A1:C3", "A:C", "1:3", "Sheet1"])(
    "accepts %s",
    (range) => {
      expect(validateRange(range)).toBe(range);
    }
  );

  it.each([
    ["empty", ""],
    ["query injection", "A1?valueRenderOption=FORMULA"],
    ["path segment", "A1/../x"],
    ["bad cell", "Sheet1!A1:B$"],
    ["unbalanced quote", "'Sheet!A1"],
    ["too long", "A".repeat(201)],
  ])("rejects invalid A1 range: %s", (_label, range) => {
    expect(() => validateRange(range)).toThrow("Invalid A1 range");
  });

  it("rejects non-string input", () => {
    expect(() => validateRange(null as unknown)).toThrow("Invalid A1 range");
  });
});

describe("validateValues", () => {
  it("accepts a 2D array of strings, numbers and booleans", () => {
    expect(validateValues([["a", 1, true], ["b"]])).toEqual([["a", 1, true], ["b"]]);
  });

  it.each([
    ["not an array", "abc"],
    ["empty array", []],
    ["1D array", ["a", "b"]],
    ["object cell", [[{ a: 1 }]]],
    ["null cell", [[null]]],
    ["nested array cell", [[["x"]]]],
  ])("rejects invalid values shape: %s", (_label, values) => {
    expect(() => validateValues(values as unknown)).toThrow(/"values"/);
  });
});

describe("write size limits", () => {
  it("rejects too many rows", () => {
    const rows = Array.from({ length: MAX_WRITE_ROWS + 1 }, () => ["x"]);
    expect(() => validateValues(rows)).toThrow("maximum of");
  });

  it("rejects too many cells", () => {
    const row = Array.from({ length: 101 }, () => "x");
    const rows = Array.from({ length: Math.floor(MAX_WRITE_CELLS / 101) + 1 }, () => row);
    expect(() => validateValues(rows)).toThrow("cells");
  });

  it("rejects an oversized payload before calling Google", async () => {
    const fetchMock = vi.fn();
    vi.stubGlobal("fetch", fetchMock);
    const big = "x".repeat(MAX_WRITE_BODY_BYTES / 2 + 1);
    await expect(writeSheetRange("tok", env, VALID_ID, "A1:A2", [[big], [big]])).rejects.toThrow("payload size");
    expect(fetchMock).not.toHaveBeenCalled();
  });
});

describe("Google response handling", () => {
  beforeEach(() => {
    vi.stubGlobal("fetch", vi.fn());
  });
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("throws a minimal error on non-OK response without leaking the body", async () => {
    (fetch as any).mockResolvedValue(
      new Response(JSON.stringify({ error: { message: "SECRET-DETAIL" } }), { status: 403 })
    );
    const err = (await readSheetRange("tok", env, VALID_ID, "A1:B2").catch((e: Error) => e)) as Error;
    expect(err.message).toContain("HTTP 403");
    expect(err.message).not.toContain("SECRET-DETAIL");
  });

  it("returns parsed JSON on OK response and encodes the URL safely", async () => {
    (fetch as any).mockResolvedValue(new Response(JSON.stringify({ values: [["1"]] }), { status: 200 }));
    const result = await readSheetRange("tok", env, VALID_ID, "Sheet1!A1:B2");
    expect(result).toEqual({ values: [["1"]] });
    const url = (fetch as any).mock.calls[0][0] as string;
    expect(url).toBe(`https://sheets.googleapis.com/v4/spreadsheets/${VALID_ID}/values/Sheet1!A1%3AB2`);
  });

  it("append propagates non-OK as an error", async () => {
    (fetch as any).mockResolvedValue(new Response("nope", { status: 500 }));
    await expect(appendSheetRows("tok", env, VALID_ID, "Sheet1", [["a"]])).rejects.toThrow("HTTP 500");
  });

  it("does not call Google when the spreadsheetId is invalid", async () => {
    await expect(readSheetRange("tok", env, "../x", "A1")).rejects.toThrow("Invalid spreadsheet ID format");
    expect(fetch).not.toHaveBeenCalled();
  });
});
