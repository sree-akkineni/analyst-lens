import { describe, expect, it } from "vitest";
import {
  capabilityRequirements,
  detectCapabilities,
} from "../../src/platform/office/capabilityMap";

describe("Office capability map", () => {
  it("uses the documented stable requirement set for every shipped adapter feature", () => {
    expect(capabilityRequirements).toEqual({
      directDependents: { set: "ExcelApi", version: "1.13" },
      directPrecedents: { set: "ExcelApi", version: "1.12" },
      formatCellProperties: { set: "ExcelApi", version: "1.9" },
      outlineGroups: { set: "ExcelApi", version: "1.10" },
      recalculationEvents: { set: "ExcelApi", version: "1.8" },
      selectionEvents: { set: "ExcelApi", version: "1.9" },
      sharedRuntime: { set: "SharedRuntime", version: "1.1" },
    });
  });

  it("checks each requirement without treating a newer unrelated feature as support", () => {
    const checked: string[] = [];
    const capabilities = detectCapabilities({
      isSetSupported(set, version) {
        checked.push(`${set}:${version}`);
        return version !== "1.13";
      },
    });

    expect(capabilities.directDependents.supported).toBe(false);
    expect(capabilities.directPrecedents.supported).toBe(true);
    expect(checked).toEqual([
      "ExcelApi:1.13",
      "ExcelApi:1.12",
      "ExcelApi:1.9",
      "ExcelApi:1.10",
      "ExcelApi:1.8",
      "ExcelApi:1.9",
      "SharedRuntime:1.1",
    ]);
  });
});
