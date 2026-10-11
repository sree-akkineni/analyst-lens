export const capabilityRequirements = {
  directDependents: { set: "ExcelApi", version: "1.13" },
  directPrecedents: { set: "ExcelApi", version: "1.12" },
  formatCellProperties: { set: "ExcelApi", version: "1.9" },
  outlineGroups: { set: "ExcelApi", version: "1.10" },
  recalculationEvents: { set: "ExcelApi", version: "1.8" },
  selectionEvents: { set: "ExcelApi", version: "1.9" },
  sharedRuntime: { set: "SharedRuntime", version: "1.1" },
} as const;

export type CapabilityName = keyof typeof capabilityRequirements;

export type RequirementSetSupport = {
  isSetSupported(set: string, version: string): boolean;
};

export type CapabilityState = {
  set: string;
  supported: boolean;
  version: string;
};

export type CapabilityMap = Record<CapabilityName, CapabilityState>;

export const detectCapabilities = (
  requirements: RequirementSetSupport,
): CapabilityMap =>
  Object.fromEntries(
    Object.entries(capabilityRequirements).map(([name, requirement]) => [
      name,
      {
        ...requirement,
        supported: requirements.isSetSupported(
          requirement.set,
          requirement.version,
        ),
      },
    ]),
  ) as CapabilityMap;
