import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";

type ToolchainPackageJson = {
  engines: { node: string };
  scripts: {
    build: string;
    "build:e2e": string;
    "test:e2e": string;
    verify: string;
    "verify:manifest": string;
    "verify:network": string;
    "verify:office-boundary": string;
  };
  devDependencies: Record<string, string> & {
    "@types/node": string;
    "fs-extra": string;
    "office-addin-dev-certs": string;
    "office-addin-dev-settings"?: string;
    "office-addin-manifest": string;
  };
};

const packageJson = JSON.parse(
  readFileSync(resolve("package.json"), "utf8"),
) as ToolchainPackageJson;
const appCss = readFileSync(resolve("src/ui/app.css"), "utf8");
const quickTraceCss = readFileSync(resolve("src/ui/quick-trace/quickTrace.css"), "utf8");

describe("package toolchain contract", () => {
  it("uses Node 22 declarations with the Node 22 runtime", () => {
    expect(packageJson.devDependencies["@types/node"]).toMatch(/^22\.\d+\.\d+$/);
  });

  it("advertises the locked tooling's true Node 22 floor", () => {
    expect(packageJson.engines.node).toBe(">=22.13 <23");
  });

  it("prevents conflicting color variables in Playwright subprocesses", () => {
    expect(packageJson.scripts["test:e2e"]).toBe(
      "env -u NO_COLOR playwright test --pass-with-no-tests",
    );
  });

  it("builds both runtimes and exposes the architecture and network gates", () => {
    expect(packageJson.scripts.build).toBe(
      "env -u ANALYST_LENS_FIXTURES vite build && tsc -p local-host/tsconfig.json && node scripts/verify-no-fixture-bundle.mjs",
    );
    expect(packageJson.scripts["build:e2e"]).toBe(
      "env ANALYST_LENS_FIXTURES=1 vite build && tsc -p local-host/tsconfig.json",
    );
    expect(packageJson.scripts["verify:network"]).toBe(
      "node scripts/verify-no-outbound.mjs",
    );
    expect(packageJson.scripts["verify:manifest"]).toBe(
      "office-addin-manifest validate manifest.xml && office-addin-manifest validate manifest.json",
    );
    expect(packageJson.scripts["verify:office-boundary"]).toBe(
      "node scripts/verify-office-boundary.mjs",
    );
    expect(packageJson.scripts.verify).toBe(
      "npm run lint && npm run typecheck && npm test && npm run test:e2e && npm run build && npm run verify:manifest && npm run verify:office-boundary && npm run verify:network",
    );
  });

  it("locks the Microsoft manifest validator and Mac certificate tooling", () => {
    expect(packageJson.devDependencies["@axe-core/playwright"]).toBe("4.12.1");
    expect(packageJson.devDependencies).not.toHaveProperty(
      "@microsoft/m365agentstoolkit-cli",
    );
    expect(packageJson.devDependencies["office-addin-dev-certs"]).toBe("2.0.10");
    expect(packageJson.devDependencies["fs-extra"]).toBe("7.0.1");
    expect(packageJson.devDependencies).not.toHaveProperty(
      "office-addin-dev-settings",
    );
    expect(packageJson.devDependencies["office-addin-manifest"]).toBe("2.1.6");
  });

  it("ships a pinned production-only installer toolchain", () => {
    const installerPackage = JSON.parse(
      readFileSync(resolve("installer-tools/package.json"), "utf8"),
    ) as {
      dependencies: Record<string, string>;
      private: boolean;
      type: string;
    };

    expect(installerPackage.private).toBe(true);
    expect(installerPackage.type).toBe("module");
    expect(installerPackage.dependencies).toEqual({
      "office-addin-dev-certs": "2.0.10",
    });
  });

  it.each([
    ["task pane", appCss],
    ["Quick Trace", quickTraceCss],
  ])("keeps the reduced-motion contract in the %s stylesheet", (_surface, css) => {
    expect(css).toContain("@media (prefers-reduced-motion: reduce)");
    expect(css).toContain("animation-duration: 0.01ms !important;");
    expect(css).toContain("animation-iteration-count: 1 !important;");
    expect(css).toContain("scroll-behavior: auto !important;");
    expect(css).toContain("transition-duration: 0.01ms !important;");
  });
});
