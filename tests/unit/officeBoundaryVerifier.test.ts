import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { spawnSync } from "node:child_process";
import { afterEach, describe, expect, it } from "vitest";

const verifier = resolve("scripts/verify-office-boundary.mjs");
const roots: string[] = [];

const createFixture = () => {
  const root = mkdtempSync(join(tmpdir(), "analyst-lens-office-boundary-"));
  roots.push(root);
  mkdirSync(join(root, "src", "platform", "office"), { recursive: true });
  mkdirSync(join(root, "src", "ui"), { recursive: true });
  return root;
};

afterEach(() => {
  for (const root of roots.splice(0)) {
    rmSync(root, { force: true, recursive: true });
  }
});

describe("Office boundary verifier", () => {
  it("permits Office.js only below platform/office", () => {
    const root = createFixture();
    writeFileSync(
      join(root, "src", "platform", "office", "gateway.ts"),
      "void Excel.run; void Office.context;",
    );
    writeFileSync(
      join(root, "src", "platform", "office", "quickTraceDialogHost.ts"),
      "void Office.context.ui;",
    );
    writeFileSync(
      join(root, "src", "platform", "office", "quickTraceDialogChild.ts"),
      "void Office.context.ui;",
    );

    const result = spawnSync(process.execPath, [verifier, root], {
      encoding: "utf8",
    });

    expect(result.status, result.stderr).toBe(0);
    expect(result.stdout.trim()).toBe(
      "Office boundary ok: src/platform/office is the only Office.js caller",
    );
  });

  it.each(["Excel.run", "Office.context", "Office.actions"])(
    "rejects %s outside the platform adapter",
    (forbiddenCall) => {
      const root = createFixture();
      writeFileSync(
        join(root, "src", "ui", "Inspector.tsx"),
        `void ${forbiddenCall};`,
      );

      const result = spawnSync(process.execPath, [verifier, root], {
        encoding: "utf8",
      });

      expect(result.status).not.toBe(0);
      expect(result.stderr).toContain("src/ui/Inspector.tsx");
    },
  );

  it("rejects Office.js copied into the Quick Trace UI", () => {
    const root = createFixture();
    const quickTraceUi = join(root, "src", "ui", "quick-trace");
    mkdirSync(quickTraceUi, { recursive: true });
    writeFileSync(
      join(quickTraceUi, "QuickTraceApp.tsx"),
      "void Office.context.ui;",
    );

    const result = spawnSync(process.execPath, [verifier, root], {
      encoding: "utf8",
    });

    expect(result.status).not.toBe(0);
    expect(result.stderr).toContain("src/ui/quick-trace/QuickTraceApp.tsx");
  });
});
