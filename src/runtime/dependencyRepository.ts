import type { OfficeGateway } from "../platform/office/officeGateway";
import {
  canonicalCellId,
  type CellRef,
  type DependencyDirection,
  type DependencyRef,
} from "../shared/types";

export type DependencyRepository = Readonly<{
  clear(): void;
  get(
    ref: CellRef,
    direction: DependencyDirection,
    signal?: AbortSignal,
  ): Promise<readonly DependencyRef[]>;
}>;

const MAX_CACHE_ENTRIES = 500;
const repositories = new WeakMap<OfficeGateway, DependencyRepository>();

const cloneDependencies = (
  dependencies: readonly DependencyRef[],
): readonly DependencyRef[] => dependencies.map((dependency) => ({ ...dependency }));

export const dependencyRepositoryFor = (
  gateway: OfficeGateway,
): DependencyRepository => {
  const existing = repositories.get(gateway);
  if (existing !== undefined) return existing;

  const cache = new Map<string, readonly DependencyRef[]>();
  const repository: DependencyRepository = {
    clear: () => cache.clear(),
    async get(ref, direction, signal) {
      signal?.throwIfAborted();
      const formulaVersion = "formula" in ref ? ref.formula ?? "" : "";
      const key = `${canonicalCellId(ref)}::${formulaVersion}::${direction}`;
      const cached = cache.get(key);
      if (cached !== undefined) return cloneDependencies(cached);

      const dependencies = await gateway.getDirectDependencies(ref, direction, signal);
      signal?.throwIfAborted();
      if (cache.size >= MAX_CACHE_ENTRIES) {
        const oldest = cache.keys().next().value;
        if (oldest !== undefined) cache.delete(oldest);
      }
      cache.set(key, cloneDependencies(dependencies));
      return cloneDependencies(dependencies);
    },
  };
  repositories.set(gateway, repository);
  return repository;
};
