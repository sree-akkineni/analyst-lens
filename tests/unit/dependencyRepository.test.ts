import { describe, expect, it } from "vitest";
import type { OfficeGateway } from "../../src/platform/office/officeGateway";
import { dependencyRepositoryFor } from "../../src/runtime/dependencyRepository";
import {
  sampleModelScenario,
  createSimulatedWorkbookFixture,
} from "../fixtures/simulatedWorkbook";

const fixture = createSimulatedWorkbookFixture({ initialState: "formula" });
const root = sampleModelScenario.states.formula;
const input = sampleModelScenario.states.hardcode;
const gatewayWithDependencies = (
  read: OfficeGateway["getDirectDependencies"],
): OfficeGateway => ({ ...fixture.gateway, getDirectDependencies: read });

describe("dependency repository", () => {
  it("shares one cached dependency read per gateway and direction", async () => {
    let reads = 0;
    const gateway = gatewayWithDependencies(async () => {
      reads += 1;
      return [input];
    });
    const first = dependencyRepositoryFor(gateway);
    const second = dependencyRepositoryFor(gateway);
    await first.get(root, "precedent");
    await second.get(root, "precedent");

    expect(second).toBe(first);
    expect(reads).toBe(1);
  });

  it("keeps formula identity and direction in the cache key", async () => {
    let reads = 0;
    const changedFormula = { ...root, formula: "=1+1" };
    const repository = dependencyRepositoryFor(gatewayWithDependencies(async () => {
      reads += 1;
      return [input];
    }));
    await repository.get(root, "precedent");
    await repository.get(changedFormula, "precedent");
    await repository.get(root, "dependent");

    expect(reads).toBe(3);
  });

  it("returns defensive copies of cached dependencies", async () => {
    const repository = dependencyRepositoryFor(gatewayWithDependencies(async () => [input]));
    const first = await repository.get(root, "precedent");
    const second = await repository.get(root, "precedent");

    expect(second).toEqual(first);
    expect(second).not.toBe(first);
  });

  it("does not cache an aborted read", async () => {
    let reads = 0;
    const gateway = gatewayWithDependencies(async (_ref, _direction, signal) => {
      reads += 1;
      signal?.throwIfAborted();
      return [input];
    });
    const repository = dependencyRepositoryFor(gateway);
    const abort = new AbortController();
    abort.abort();

    await expect(repository.get(root, "precedent", abort.signal)).rejects.toThrow();
    await repository.get(root, "precedent");
    expect(reads).toBe(1);
  });

  it("does not cache a result aborted after the gateway read", async () => {
    let reads = 0;
    const abort = new AbortController();
    const repository = dependencyRepositoryFor(gatewayWithDependencies(async () => {
      reads += 1;
      abort.abort();
      return [input];
    }));

    await expect(repository.get(root, "precedent", abort.signal)).rejects.toThrow();
    await repository.get(root, "precedent");
    expect(reads).toBe(2);
  });

  it("clears cached edges for recalculation", async () => {
    let reads = 0;
    const repository = dependencyRepositoryFor(
      gatewayWithDependencies(async () => {
        reads += 1;
        return [input];
      }),
    );
    await repository.get(root, "precedent");
    repository.clear();
    await repository.get(root, "precedent");

    expect(reads).toBe(2);
  });

  it("evicts the oldest entry when the 500-entry cap is exceeded", async () => {
    let reads = 0;
    const repository = dependencyRepositoryFor(gatewayWithDependencies(async () => {
      reads += 1;
      return [input];
    }));
    for (let index = 0; index <= 500; index += 1) {
      await repository.get({ ...root, address: `$A$${index + 1}` }, "precedent");
    }
    await repository.get({ ...root, address: "$A$1" }, "precedent");

    expect(reads).toBe(502);
  });
});
