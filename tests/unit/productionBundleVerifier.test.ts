import {
  mkdirSync,
  mkdtempSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { spawnSync } from "node:child_process";
import { afterEach, describe, expect, it } from "vitest";

const verifier = resolve("scripts/verify-no-fixture-bundle.mjs");
const roots: string[] = [];

const createBundle = (source = "console.log('production');") => {
  const root = mkdtempSync(join(tmpdir(), "analyst-lens-production-"));
  roots.push(root);
  const assets = join(root, "assets");
  mkdirSync(assets, { recursive: true });
  writeFileSync(join(assets, "taskpane.js"), source);
  writeFileSync(join(assets, "quick-trace.js"), "console.log('quick trace');");
  writeFileSync(join(assets, "shared-AbCdEf12.js"), "console.log('shared');");
  return root;
};

afterEach(() => {
  roots.splice(0).forEach((root) => rmSync(root, { force: true, recursive: true }));
});

describe("production bundle fixture verifier", () => {
  it("accepts a fixture-free production bundle", () => {
    const result = spawnSync(process.execPath, [verifier, createBundle()], {
      encoding: "utf8",
    });

    expect(result.status, result.stderr).toBe(0);
  });

  it.each([
    ["fixture marker", "const fixture='sample-model';"],
    ["fixture module name", "import('./simulatedWorkbook-test.js');"],
    ["Quick Trace fixture module name", "import('./quickTraceFixture-test.js');"],
  ])("rejects a production bundle containing a %s", (_kind, source) => {
    const result = spawnSync(process.execPath, [verifier, createBundle(source)], {
      encoding: "utf8",
    });

    expect(result.status).not.toBe(0);
    expect(result.stderr).toContain("production bundle contains standalone fixture code");
  });

  it.each(["taskpane.js", "quick-trace.js", "shared-AbCdEf12.js"])(
    "scans every emitted JavaScript location including %s",
    (fileName) => {
      const root = createBundle();
      writeFileSync(join(root, "assets", fileName), "const fixture='quickTraceFixture';");
      const result = spawnSync(process.execPath, [verifier, root], {
        encoding: "utf8",
      });

      expect(result.status).not.toBe(0);
      expect(result.stderr).toContain(fileName);
    },
  );
});
