import type { CommandId, CommandResult } from "../../shared/types";
import type { WriteDescriptor } from "./revert";

export type FormatBorderMutation = {
  side: "bottom" | "top";
  style: "Continuous" | "Double";
  weight: "Medium" | "Thin";
};

export type FormatMutation =
  | { border: FormatBorderMutation }
  | { clearBorders: true }
  | { clearFormats: true }
  | { decimalDelta: -1 | 1 }
  | { fontColor: string }
  | { numberFormat: string };

const formatMutations = {
  "format.border.bottom": {
    border: { side: "bottom", style: "Continuous", weight: "Thin" },
  },
  "format.border.clear": { clearBorders: true },
  "format.border.doubleBottom": {
    border: { side: "bottom", style: "Double", weight: "Medium" },
  },
  "format.border.top": {
    border: { side: "top", style: "Continuous", weight: "Thin" },
  },
  "format.clear": { clearFormats: true },
  "format.currency.integer": { numberFormat: "$#,##0;($#,##0);-" },
  "format.date": { numberFormat: "m/d/yy" },
  "format.decimal.one": { numberFormat: "#,##0.0;(#,##0.0);-" },
  "format.decimal.two": { numberFormat: "#,##0.00;(#,##0.00);-" },
  "format.decimals.decrease": { decimalDelta: -1 },
  "format.decimals.increase": { decimalDelta: 1 },
  "format.formula": { fontColor: "#000000" },
  "format.input": { fontColor: "#0000FF" },
  "format.integer": { numberFormat: "#,##0;(#,##0);-" },
  "format.link": { fontColor: "#008000" },
  "format.multiple.oneDecimal": { numberFormat: "0.0x" },
  "format.percent.oneDecimal": { numberFormat: "0.0%;(0.0%);-" },
  "format.warning": { fontColor: "#C00000" },
} as const satisfies Readonly<Record<string, FormatMutation>>;

export const formatDescriptor = (id: CommandId): FormatMutation => {
  if (!(id in formatMutations)) {
    throw new Error(`Unknown format command: ${id}`);
  }
  return formatMutations[id as keyof typeof formatMutations];
};

export const adjustNumberFormatDecimals = (
  numberFormat: string,
  delta: -1 | 1,
): string =>
  numberFormat
    .split(";")
    .map((section) => section.replace(
      /([#0?](?:[#0?,]*[#0?])?)(?:\.([#0?]+))?/,
      (_match, integer: string, decimals: string | undefined) => {
        const current = decimals ?? "";
        const next = delta === 1 ? `${current}0` : current.slice(0, -1);
        return next.length === 0 ? integer : `${integer}.${next}`;
      },
    ))
    .join(";");

export type SelectionFormatSnapshot = {
  address: string;
  captured: "allFormats" | "borders" | "fontColor" | "numberFormat";
  cellProperties: unknown[][];
  numberFormat: string[][];
  sheetProtected: boolean;
  sheetId: string;
  workbookId: string;
};

export type SelectionFormatWritePort = {
  applySelectionFormat(
    mutation: FormatMutation,
    target: SelectionFormatSnapshot,
  ): Promise<void>;
  captureSelectionFormat(
    mutation: FormatMutation,
  ): Promise<SelectionFormatSnapshot>;
  preflightSelection(
    mutation: FormatMutation,
    target: SelectionFormatSnapshot,
  ): Promise<CommandResult>;
  restoreSelectionFormat(snapshot: SelectionFormatSnapshot): Promise<void>;
  verifySelectionFormat(
    mutation: FormatMutation,
    target: SelectionFormatSnapshot,
  ): Promise<boolean>;
  verifySelectionRestore(snapshot: SelectionFormatSnapshot): Promise<boolean>;
};

export const formatSelection = (
  id: CommandId,
): WriteDescriptor<SelectionFormatWritePort, SelectionFormatSnapshot> => {
  const mutation = formatDescriptor(id);
  return {
    apply: (port, snapshot) => port.applySelectionFormat(mutation, snapshot),
    capture: (port) => port.captureSelectionFormat(mutation),
    id,
    preflight: (port, snapshot) => port.preflightSelection(mutation, snapshot),
    revert: (port, snapshot) => port.restoreSelectionFormat(snapshot),
    verify: (port, snapshot) => port.verifySelectionFormat(mutation, snapshot),
    verifyRevert: (port, snapshot) => port.verifySelectionRestore(snapshot),
  };
};

export type NumberFormatSnapshot = {
  address: string;
  numberFormat: string;
  sheetId: string;
  workbookId: string;
};

export type SelectionWritePort = {
  preflightSelection(target?: NumberFormatSnapshot): Promise<CommandResult>;
  readSelectionNumberFormat(target?: NumberFormatSnapshot): Promise<NumberFormatSnapshot>;
  restoreSelectionNumberFormat(snapshot: NumberFormatSnapshot): Promise<void>;
  setSelectionNumberFormat(
    numberFormat: string,
    target?: NumberFormatSnapshot,
  ): Promise<void>;
};

const sameSnapshot = (
  left: NumberFormatSnapshot,
  right: NumberFormatSnapshot,
) =>
  left.address === right.address &&
  left.numberFormat === right.numberFormat &&
  left.sheetId === right.sheetId &&
  left.workbookId === right.workbookId;

export const formatNumber = (
  numberFormat: string,
): WriteDescriptor<SelectionWritePort, NumberFormatSnapshot> => ({
  apply: (port, snapshot) => port.setSelectionNumberFormat(numberFormat, snapshot),
  capture: (port) => port.readSelectionNumberFormat(),
  id: "format.number.custom",
  preflight: (port, snapshot) => port.preflightSelection(snapshot),
  revert: (port, snapshot) => port.restoreSelectionNumberFormat(snapshot),
  async verify(port, snapshot) {
    const readback = await port.readSelectionNumberFormat(snapshot);
    return readback.numberFormat === numberFormat;
  },
  async verifyRevert(port, snapshot) {
    return sameSnapshot(await port.readSelectionNumberFormat(snapshot), snapshot);
  },
});
