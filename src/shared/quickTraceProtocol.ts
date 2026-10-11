import { z } from "zod";
import type { QuickTraceRow } from "../domain/dependencies/quickTraceTree";
import type { CellRef, DependencyDirection } from "./types";

export const QUICK_TRACE_PROTOCOL_VERSION = 1 as const;
export type QuickTraceDirection = DependencyDirection;

const cellRefSchema = z.object({
  address: z.string().min(1),
  sheetId: z.string().min(1),
  sheetName: z.string().min(1),
  workbookId: z.string().min(1),
});
const dependencyNodeSchema = z.object({
  id: z.string().min(1),
  ref: cellRefSchema,
  unresolved: z.enum(["dynamic", "external"]).optional(),
  value: z.union([z.string(), z.number(), z.boolean(), z.null()]).optional(),
});
const envelope = z.object({
  nonce: z.string().min(1),
  protocol: z.literal(QUICK_TRACE_PROTOCOL_VERSION),
  sequence: z.number().int().nonnegative(),
  sessionId: z.string().min(1),
});
const rowSchema = z.object({
  depth: z.number().int().nonnegative(),
  disclosure: z.enum(["collapsed", "expanded", "leaf"]),
  displayValue: z.string(),
  issue: z.enum(["cycle", "dynamic", "external", "unavailable"]).nullable(),
  node: dependencyNodeSchema,
  parentRowId: z.string().min(1).nullable(),
  rowId: z.string().min(1),
});
const stateSchema = z.object({
  capped: z.boolean(),
  direction: z.enum(["precedent", "dependent"]),
  formula: z.string().nullable(),
  navigationBlocked: z.boolean(),
  nodeCount: z.number().int().nonnegative(),
  notice: z.string().nullable(),
  origin: cellRefSchema,
  pendingExcelSelection: cellRefSchema.nullable(),
  root: cellRefSchema,
  rows: z.array(rowSchema),
  selectedFormula: z.string().nullable(),
  selectedRowId: z.string().min(1),
  status: z.enum(["opening", "loading", "ready", "error"]),
});
const dialogIntentSchema = z.discriminatedUnion("type", [
  envelope.extend({ type: z.literal("quickTrace.ready") }),
  envelope.extend({
    payload: z.object({ movement: z.enum(["previous", "next"]) }),
    type: z.literal("quickTrace.move"),
  }),
  envelope.extend({ type: z.literal("quickTrace.left") }),
  envelope.extend({ type: z.literal("quickTrace.right") }),
  envelope.extend({
    payload: z.object({ direction: z.enum(["precedent", "dependent"]) }),
    type: z.literal("quickTrace.setDirection"),
  }),
  envelope.extend({
    payload: z.object({ rowId: z.string().min(1) }),
    type: z.literal("quickTrace.navigate"),
  }),
  envelope.extend({
    payload: z.object({ rowId: z.string().min(1) }),
    type: z.literal("quickTrace.toggleBranch"),
  }),
  envelope.extend({
    payload: z.object({ rowId: z.string().min(1) }),
    type: z.literal("quickTrace.reroot"),
  }),
  envelope.extend({ type: z.literal("quickTrace.returnToOrigin") }),
  envelope.extend({ type: z.literal("quickTrace.useExcelSelection") }),
  envelope.extend({ type: z.literal("quickTrace.close") }),
]);
const hostEventSchema = z.discriminatedUnion("type", [
  envelope.extend({
    payload: z.object({ state: stateSchema }),
    type: z.literal("quickTrace.initialize"),
  }),
  envelope.extend({
    payload: z.object({ state: stateSchema }),
    type: z.literal("quickTrace.state"),
  }),
  envelope.extend({
    payload: z.object({ code: z.string().min(1), message: z.string().min(1) }),
    type: z.literal("quickTrace.error"),
  }),
]);

export type QuickTraceSessionIdentity = Readonly<{
  nonce: string;
  sessionId: string;
}>;
export type QuickTraceEnvelope<Type extends string, Payload = never> =
  Readonly<QuickTraceSessionIdentity & {
    protocol: typeof QUICK_TRACE_PROTOCOL_VERSION;
    sequence: number;
    type: Type;
  }> & ([Payload] extends [never] ? object : Readonly<{ payload: Payload }>);

export type QuickTraceSnapshot = Readonly<{
  capped: boolean;
  direction: DependencyDirection;
  formula: string | null;
  navigationBlocked: boolean;
  nodeCount: number;
  notice: string | null;
  origin: CellRef;
  pendingExcelSelection: CellRef | null;
  root: CellRef;
  rows: readonly QuickTraceRow[];
  selectedFormula: string | null;
  selectedRowId: string;
  status: "opening" | "loading" | "ready" | "error";
}>;

export type QuickTraceDialogIntent = z.infer<typeof dialogIntentSchema>;
export type QuickTraceHostEvent =
  | QuickTraceEnvelope<"quickTrace.initialize", { state: QuickTraceSnapshot }>
  | QuickTraceEnvelope<"quickTrace.state", { state: QuickTraceSnapshot }>
  | QuickTraceEnvelope<"quickTrace.error", { code: string; message: string }>;

const assertSession = <Message extends QuickTraceSessionIdentity>(
  message: Message,
  expected: QuickTraceSessionIdentity,
): Message => {
  if (message.sessionId !== expected.sessionId || message.nonce !== expected.nonce) {
    throw new Error("Quick Trace rejected a stale or untrusted session message");
  }
  return message;
};

export const serializeQuickTraceMessage = (
  message: QuickTraceDialogIntent | QuickTraceHostEvent,
): string => JSON.stringify(message);

export const parseQuickTraceDialogIntent = (
  raw: string,
  expected: QuickTraceSessionIdentity,
): QuickTraceDialogIntent => assertSession(
  dialogIntentSchema.parse(JSON.parse(raw) as unknown),
  expected,
);

export const parseQuickTraceHostEvent = (
  raw: string,
  expected: QuickTraceSessionIdentity,
): QuickTraceHostEvent => assertSession(
  hostEventSchema.parse(JSON.parse(raw) as unknown) as QuickTraceHostEvent,
  expected,
);
