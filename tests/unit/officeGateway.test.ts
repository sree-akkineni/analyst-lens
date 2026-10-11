import { describe, expect, it } from "vitest";
import {
  normalizeOfficeError,
  OfficeGatewayError,
} from "../../src/platform/office/officeGateway";

describe("Office gateway errors", () => {
  it.each([
    ["ItemNotFound", "NOT_FOUND"],
    ["InvalidReference", "INVALID_REFERENCE"],
    ["InvalidObjectPath", "INVALID_REFERENCE"],
    ["AccessDenied", "ACCESS_DENIED"],
    ["UnsupportedOperation", "UNSUPPORTED_OPERATION"],
    ["ApiNotFound", "UNSUPPORTED_OPERATION"],
  ])("normalizes %s without exposing workbook content", (sourceCode, expected) => {
    const normalized = normalizeOfficeError({ code: sourceCode });

    expect(normalized).toBeInstanceOf(OfficeGatewayError);
    expect(normalized.code).toBe(expected);
    expect(normalized.message).not.toContain("formula");
  });

  it("preserves typed abort and stale request results", () => {
    const stale = new OfficeGatewayError("STALE_REQUEST", "Superseded request");

    expect(normalizeOfficeError(stale)).toBe(stale);
    expect(
      normalizeOfficeError(new DOMException("cancelled", "AbortError")).code,
    ).toBe("ABORTED");
  });
});
