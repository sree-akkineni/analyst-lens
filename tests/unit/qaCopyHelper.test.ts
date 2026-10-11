import {
  chmodSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { spawnSync } from "node:child_process";
import { afterEach, describe, expect, it } from "vitest";

const scriptPath = resolve("scripts/qa-copy-fixture.sh");
const temporaryRoots: string[] = [];

const createFixture = () => {
  const root = mkdtempSync(join(tmpdir(), "analyst-lens-qa-copy-"));
  temporaryRoots.push(root);
  const sourceDir = join(root, "Source With Spaces");
  const qaRoot = join(root, "QA With Spaces");
  const sourceFile = join(sourceDir, "Sample_Model.xlsx");
  mkdirSync(sourceDir, { recursive: true });
  writeFileSync(sourceFile, "safe workbook fixture\n");
  chmodSync(sourceFile, 0o640);
  return { qaRoot, sourceFile };
};

afterEach(() => {
  for (const root of temporaryRoots.splice(0)) {
    rmSync(root, { force: true, recursive: true });
  }
});

describe("QA copy helper", () => {
  it("creates a timestamped byte-identical copy beneath the QA root", () => {
    const fixture = createFixture();

    const result = spawnSync(scriptPath, [], {
      encoding: "utf8",
      env: {
        ...process.env,
        ANALYST_LENS_QA_ROOT: fixture.qaRoot,
        ANALYST_LENS_QA_SOURCE_FILE: fixture.sourceFile,
        ANALYST_LENS_QA_TIMESTAMP: "20260715-220500",
      },
    });

    const target = result.stdout.trim();
    expect(result.status, result.stderr).toBe(0);
    expect(target).toBe(
      join(fixture.qaRoot, "Sample_Model_qa_20260715-220500.xlsx"),
    );
    expect(readFileSync(target)).toEqual(readFileSync(fixture.sourceFile));
  });

  it("refuses an in-place target and leaves the source unchanged", () => {
    const fixture = createFixture();
    const before = readFileSync(fixture.sourceFile);

    const result = spawnSync(scriptPath, [], {
      encoding: "utf8",
      env: {
        ...process.env,
        ANALYST_LENS_QA_ROOT: fixture.qaRoot,
        ANALYST_LENS_QA_SOURCE_FILE: fixture.sourceFile,
        ANALYST_LENS_QA_TARGET: fixture.sourceFile,
      },
    });

    expect(result.status).not.toBe(0);
    expect(result.stderr).toContain("refuses in-place QA");
    expect(readFileSync(fixture.sourceFile)).toEqual(before);
  });
});
