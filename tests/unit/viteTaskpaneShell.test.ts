import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import {
  analystLensAssetFileName,
  analystLensFixturesEnabled,
  quickTraceAssetFileName,
  quickTraceHtml,
  taskpaneAssetFileName,
  taskpaneHtml,
} from "../../vite.config";

describe("Vite task-pane shell", () => {
  it("emits a deterministic stylesheet and links it from taskpane.html", () => {
    expect(taskpaneAssetFileName({ names: ["taskpane.css"] })).toBe("assets/taskpane.css");
    expect(taskpaneAssetFileName({ names: ["logo.svg"] })).toBe("assets/[name]-[hash][extname]");
    expect(taskpaneHtml).toContain('<link rel="stylesheet" href="/assets/taskpane.css" />');
  });

  it("emits a deterministic Quick Trace dialog shell", () => {
    expect(quickTraceAssetFileName({ names: ["quickTrace.css"] }))
      .toBe("assets/quick-trace.css");
    expect(analystLensAssetFileName({ names: ["quickTrace.css"] }))
      .toBe("assets/quick-trace.css");
    expect(quickTraceHtml).toContain("/assets/quick-trace.css");
    expect(quickTraceHtml).toContain("/assets/quick-trace.js");
    expect(quickTraceHtml).toContain("quick-trace-root");
    expect(quickTraceHtml).toContain("office.js");
    expect(readFileSync("dist/quick-trace.html", "utf8"))
      .toContain("/assets/quick-trace.js");
  });

  it("keeps standalone fixture mode out of normal production builds", () => {
    expect(analystLensFixturesEnabled).toBe(false);
  });
});
