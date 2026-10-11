import { describe, expect, it } from "vitest";
import {
  adjustNumberFormatDecimals,
  formatDescriptor,
} from "../../src/domain/commands/formatCommands";

describe("formatDescriptor", () => {
  it.each([
    ["format.input", { fontColor: "#0000FF" }],
    ["format.formula", { fontColor: "#000000" }],
    ["format.link", { fontColor: "#008000" }],
    ["format.warning", { fontColor: "#C00000" }],
    ["format.integer", { numberFormat: "#,##0;(#,##0);-" }],
    ["format.decimal.one", { numberFormat: "#,##0.0;(#,##0.0);-" }],
    ["format.decimal.two", { numberFormat: "#,##0.00;(#,##0.00);-" }],
    ["format.percent.oneDecimal", { numberFormat: "0.0%;(0.0%);-" }],
    ["format.currency.integer", { numberFormat: "$#,##0;($#,##0);-" }],
    ["format.multiple.oneDecimal", { numberFormat: "0.0x" }],
    ["format.date", { numberFormat: "m/d/yy" }],
    ["format.border.top", { border: { side: "top", style: "Continuous", weight: "Thin" } }],
    ["format.border.bottom", { border: { side: "bottom", style: "Continuous", weight: "Thin" } }],
    ["format.border.doubleBottom", { border: { side: "bottom", style: "Double", weight: "Medium" } }],
    ["format.border.clear", { clearBorders: true }],
    ["format.clear", { clearFormats: true }],
    ["format.decimals.increase", { decimalDelta: 1 }],
    ["format.decimals.decrease", { decimalDelta: -1 }],
  ] as const)("maps %s to an exact selected-range mutation", (id, expected) => {
    expect(formatDescriptor(id)).toEqual(expected);
  });

  it("rejects an unknown command instead of applying a fallback mutation", () => {
    expect(() => formatDescriptor("format.unknown")).toThrow(
      "Unknown format command: format.unknown",
    );
  });
});

describe("adjustNumberFormatDecimals", () => {
  it.each([
    ["#,##0;(#,##0);-", 1, "#,##0.0;(#,##0.0);-"],
    ["#,##0.00;(#,##0.00);-", -1, "#,##0.0;(#,##0.0);-"],
    ["0.0%;(0.0%);-", 1, "0.00%;(0.00%);-"],
    ["0.0x", -1, "0x"],
    ["#,##0;(#,##0);-", -1, "#,##0;(#,##0);-"],
    ["m/d/yy", 1, "m/d/yy"],
  ] as const)("adjusts %s by %s without changing other format sections", (format, delta, expected) => {
    expect(adjustNumberFormatDecimals(format, delta)).toBe(expected);
  });
});
