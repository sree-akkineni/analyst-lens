import {
  appendFileSync,
  chmodSync,
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { spawnSync, type SpawnSyncReturns } from "node:child_process";
import { afterEach, describe, expect, it } from "vitest";

const cliPath = resolve("scripts/analyst-lens");
const temporaryRoots: string[] = [];

type LifecycleFixture = {
  appRoot: string;
  commandRecord: string;
  env: NodeJS.ProcessEnv;
  home: string;
  installedCliPath: string;
  manifestRegistration: string;
  plistPath: string;
  sourceRoot: string;
  run: (command: string) => SpawnSyncReturns<string>;
  runInstalled: (command: string) => SpawnSyncReturns<string>;
};

const writeExecutable = (path: string, source: string) => {
  writeFileSync(path, source);
  chmodSync(path, 0o755);
};

const createLifecycleFixture = ({
  offlineCacheMiss = false,
}: { offlineCacheMiss?: boolean } = {}): LifecycleFixture => {
  const root = mkdtempSync(join(tmpdir(), "analyst-lens-lifecycle-"));
  temporaryRoots.push(root);
  const home = join(root, "Home With Spaces");
  const sourceRoot = join(root, "Source With Spaces");
  const fakeBin = join(root, "fake-bin");
  const mockState = join(root, "mock-state");
  const toolingSource = join(root, "tooling-source");
  mkdirSync(join(sourceRoot, "dist"), { recursive: true });
  mkdirSync(join(sourceRoot, "dist-host"), { recursive: true });
  mkdirSync(join(sourceRoot, "public", "icons"), { recursive: true });
  mkdirSync(fakeBin, { recursive: true });
  mkdirSync(join(toolingSource, "node_modules", ".bin"), { recursive: true });
  mkdirSync(join(toolingSource, "node_modules", "office-addin-dev-certs"), {
    recursive: true,
  });
  mkdirSync(home, { recursive: true });
  mkdirSync(mockState, { recursive: true });

  writeFileSync(
    join(sourceRoot, "package.json"),
    JSON.stringify({ version: "0.1.0" }),
  );
  writeFileSync(
    join(toolingSource, "package.json"),
    JSON.stringify({
      private: true,
      dependencies: { "office-addin-dev-certs": "2.0.10" },
    }),
  );
  writeFileSync(
    join(toolingSource, "package-lock.json"),
    JSON.stringify({ lockfileVersion: 3, packages: {} }),
  );
  writeFileSync(
    join(
      toolingSource,
      "node_modules",
      "office-addin-dev-certs",
      "runtime.js",
    ),
    "export const runtime = true;\n",
  );
  writeFileSync(join(sourceRoot, "manifest.xml"), readFileSync("manifest.xml"));
  writeFileSync(join(sourceRoot, "dist", "taskpane.html"), "Analyst Lens");
  writeFileSync(join(sourceRoot, "dist", "quick-trace.html"), "Quick Trace");
  writeFileSync(join(sourceRoot, "dist-host", "server.js"), "// host");
  for (const size of [16, 32, 80, 192]) {
    writeFileSync(join(sourceRoot, "public", "icons", `icon-${size}.png`), "png");
  }

  const record = join(mockState, "commands.log");
  writeExecutable(
    join(fakeBin, "launchctl"),
    `#!/bin/bash
set -eu
printf 'launchctl <%s>\n' "$*" >> "${record}"
case "\${1:-}" in
  print)
    test -f "${join(mockState, "service")}" || exit 1
    printf 'pid = %s\n' "\${ANALYST_LENS_TEST_PID}"
    ;;
  bootstrap)
    touch "${join(mockState, "service")}"
    ;;
  bootout)
    rm -f "${join(mockState, "service")}"
    ;;
  kickstart) ;;
  *) exit 2 ;;
esac
`,
  );
  writeExecutable(
    join(fakeBin, "curl"),
    `#!/bin/bash
set -eu
printf 'curl <%s>\n' "$*" >> "${record}"
test -f "${join(mockState, "service")}"
printf '{"name":"Analyst Lens","version":"%s"}' "\${ANALYST_LENS_TEST_HEALTH_VERSION:-0.1.0}"
`,
  );
  writeExecutable(
    join(fakeBin, "npm"),
    `#!/bin/bash
set -eu
printf 'npm <%s>\n' "$*" >> "${record}"
if [[ "\${ANALYST_LENS_TEST_OFFLINE_CACHE_MISS:-}" == "1" && "$*" == *'ci --offline'* ]]; then
  printf 'npm error code ENOTCACHED\n' >&2
  exit 1
fi
case "$*" in
  'run build') exit 0 ;;
  *) exit 2 ;;
esac
`,
  );
  writeExecutable(
    join(toolingSource, "node_modules", ".bin", "office-addin-dev-certs"),
    `#!/bin/bash
set -eu
printf 'dev-certs <%s>\n' "$*" >> "${record}"
if [[ "\${1:-}" == "verify" ]]; then
  test -f "${join(home, ".office-addin-dev-certs", "localhost.crt")}"
  test -f "${join(home, ".office-addin-dev-certs", "localhost.key")}"
  exit 0
fi
test "\${1:-}" = "install"
if [[ "\${ANALYST_LENS_TEST_CERT_INSTALL_FAIL:-}" == "1" ]]; then
  printf 'certificate install failed\n' >&2
  exit 1
fi
mkdir -p "${join(home, ".office-addin-dev-certs")}"
printf cert > "${join(home, ".office-addin-dev-certs", "localhost.crt")}"
printf key > "${join(home, ".office-addin-dev-certs", "localhost.key")}"
if [[ "\${ANALYST_LENS_TEST_OWNERSHIP_WRITE_FAIL:-}" == "1" ]]; then
  mkdir -p "${join(home, "Library", "Application Support", "Analyst Lens", "registration.owner.next")}"
fi
`,
  );

  const appRoot = join(home, "Library", "Application Support", "Analyst Lens");
  const installedCliPath = join(appRoot, "scripts", "analyst-lens");
  const plistPath = join(
    home,
    "Library",
    "LaunchAgents",
    "com.analystlens.local-host.plist",
  );
  const manifestRegistration = join(
    home,
    "Library",
    "Containers",
    "com.microsoft.Excel",
    "Data",
    "Documents",
    "wef",
    "5db4c51e-9c38-4d8b-976b-2d0f707e27f5.manifest.xml",
  );
  const env: NodeJS.ProcessEnv = {
    ...process.env,
    ANALYST_LENS_CURL: join(fakeBin, "curl"),
    ANALYST_LENS_LAUNCHCTL: join(fakeBin, "launchctl"),
    ANALYST_LENS_NPM: join(fakeBin, "npm"),
    ANALYST_LENS_SOURCE_ROOT: sourceRoot,
    ANALYST_LENS_TEST_OFFLINE_CACHE_MISS: offlineCacheMiss ? "1" : "0",
    ANALYST_LENS_TEST_PID: String(process.pid),
    ...(offlineCacheMiss
      ? {}
      : { ANALYST_LENS_TOOLING_SOURCE: toolingSource }),
    HOME: home,
  };

  return {
    appRoot,
    commandRecord: record,
    env,
    home,
    installedCliPath,
    manifestRegistration,
    plistPath,
    sourceRoot,
    run: (command: string) =>
      spawnSync(cliPath, [command], { encoding: "utf8", env }),
    runInstalled: (command: string) => {
      const installedEnv = { ...env };
      delete installedEnv.ANALYST_LENS_SOURCE_ROOT;
      delete installedEnv.ANALYST_LENS_TOOLING_SOURCE;
      return spawnSync(installedCliPath, [command], {
        encoding: "utf8",
        env: installedEnv,
      });
    },
  };
};

afterEach(() => {
  for (const root of temporaryRoots.splice(0)) {
    rmSync(root, { force: true, recursive: true });
  }
});

describe("Mac lifecycle CLI", () => {
  it("requires the Quick Trace dialog shell before deploying a release", () => {
    expect(readFileSync(cliPath, "utf8"))
      .toContain('[[ -f "${source_root}/dist/quick-trace.html" ]]');
  });

  it("installs once, reports healthy, and quotes paths in launchd arguments", () => {
    const fixture = createLifecycleFixture();

    const installed = fixture.run("install");
    const healthy = fixture.run("install");

    expect(installed.status, installed.stderr).toBe(0);
    expect(installed.stdout.trim()).toBe("installed");
    expect(healthy.status, healthy.stderr).toBe(0);
    expect(healthy.stdout.trim()).toBe("healthy");
    expect(existsSync(fixture.manifestRegistration)).toBe(true);
    const plist = readFileSync(fixture.plistPath, "utf8");
    expect(plist).toContain(
      `<string>${join(fixture.appRoot, "release", "runtime", "server.js")}</string>`,
    );
    expect(plist).not.toContain("/bin/sh");
    expect(plist).not.toContain(" -c ");
  });

  it("status fails unless launchd, PID, certificate, manifest, and health agree", () => {
    const fixture = createLifecycleFixture();
    expect(fixture.run("install").status).toBe(0);
    expect(fixture.run("status").stdout.trim()).toBe("healthy");

    rmSync(fixture.manifestRegistration);
    const missingManifest = fixture.run("status");

    expect(missingManifest.status).not.toBe(0);
    expect(missingManifest.stdout).toBe("");
  });

  it("status rejects a changed launch agent", () => {
    const fixture = createLifecycleFixture();
    expect(fixture.run("install").status).toBe(0);

    appendFileSync(fixture.plistPath, "\n<!-- changed -->\n");
    const changedLaunchAgent = fixture.run("status");

    expect(changedLaunchAgent.status).not.toBe(0);
    expect(changedLaunchAgent.stdout).toBe("");
  });

  it("status rejects a missing PID, untrusted certificate, or wrong health version", () => {
    const fixture = createLifecycleFixture();
    expect(fixture.run("install").status).toBe(0);

    fixture.env.ANALYST_LENS_TEST_PID = "999999999";
    expect(fixture.run("status").status).not.toBe(0);
    fixture.env.ANALYST_LENS_TEST_PID = String(process.pid);

    rmSync(join(fixture.home, ".office-addin-dev-certs", "localhost.crt"));
    expect(fixture.run("status").status).not.toBe(0);
    writeFileSync(
      join(fixture.home, ".office-addin-dev-certs", "localhost.crt"),
      "cert",
    );

    fixture.env.ANALYST_LENS_TEST_HEALTH_VERSION = "9.9.9";
    expect(fixture.run("status").status).not.toBe(0);
  });

  it("repairs unhealthy state and uninstalls only Analyst Lens artifacts", () => {
    const fixture = createLifecycleFixture();
    expect(fixture.run("install").status).toBe(0);
    rmSync(fixture.plistPath);
    const repaired = fixture.run("repair");
    expect(repaired.status, repaired.stderr).toBe(0);
    expect(repaired.stdout.trim()).toBe("repaired");

    const sibling = join(
      fixture.home,
      "Library",
      "Application Support",
      "Other App",
      "keep.txt",
    );
    mkdirSync(dirname(sibling), { recursive: true });
    writeFileSync(sibling, "keep");

    const first = fixture.run("uninstall");
    const second = fixture.run("uninstall");

    expect(first.status, first.stderr).toBe(0);
    expect(first.stdout.trim()).toBe("uninstalled");
    expect(second.status, second.stderr).toBe(0);
    expect(second.stdout.trim()).toBe("uninstalled");
    expect(existsSync(fixture.appRoot)).toBe(false);
    expect(existsSync(fixture.plistPath)).toBe(false);
    expect(existsSync(fixture.manifestRegistration)).toBe(false);
    expect(readFileSync(sibling, "utf8")).toBe("keep");
  });

  it("supports status, repair, and uninstall after the source checkout is deleted", () => {
    const fixture = createLifecycleFixture();
    expect(fixture.run("install").status).toBe(0);
    rmSync(fixture.sourceRoot, { recursive: true });

    const status = fixture.runInstalled("status");
    expect(status.status, status.stderr).toBe(0);
    expect(status.stdout.trim()).toBe("healthy");

    rmSync(fixture.plistPath);
    const repaired = fixture.runInstalled("repair");
    expect(repaired.status, repaired.stderr).toBe(0);
    expect(repaired.stdout.trim()).toBe("repaired");

    const uninstalled = fixture.runInstalled("uninstall");
    expect(uninstalled.status, uninstalled.stderr).toBe(0);
    expect(uninstalled.stdout.trim()).toBe("uninstalled");
    expect(existsSync(fixture.appRoot)).toBe(false);
  });

  it("uses only locked certificate tooling and has no cloud installer path", () => {
    const cli = readFileSync(cliPath, "utf8");
    const normalized = cli.replace(/\\\s*\n\s*/gu, " ");

    expect(normalized).not.toMatch(
      /\bnpx\b|"\$npm_command"\s+exec\b|\bnpm\s+exec\b|--package(?:=|\s)|"\$npm_command"\s+install\b|\b(?:pnpm|yarn)\s+(?:add|dlx|install)\b|registry\.npmjs\.org/iu,
    );
    expect(normalized).toMatch(
      /"\$npm_command"\s+ci\s+--offline\s+--ignore-scripts\s+--omit=dev/iu,
    );
    expect(normalized).toMatch(
      /node_modules\/\.bin\/office-addin-dev-certs"\s+verify/iu,
    );
    expect(normalized).not.toMatch(/security_command.*verify-cert/iu);
    expect(normalized).not.toMatch(
      /\batk\b|login\.microsoftonline\.com|titles\.prod\.mos\.microsoft\.com|graph\.microsoft\.com/iu,
    );

    const fixture = createLifecycleFixture();
    expect(fixture.run("install").status).toBe(0);
    writeFileSync(fixture.commandRecord, "");
    expect(fixture.runInstalled("status").status).toBe(0);
    expect(readFileSync(fixture.commandRecord, "utf8")).not.toMatch(/npm </u);
  });

  it("fails clearly when required installer tools are absent from the offline cache", () => {
    const fixture = createLifecycleFixture({ offlineCacheMiss: true });

    const installed = fixture.run("install");

    expect(installed.status).not.toBe(0);
    expect(installed.stderr).toContain("offline dependency cache miss");
    expect(existsSync(fixture.appRoot)).toBe(false);
  });

  it("rolls back a new local payload when certificate installation fails", () => {
    const fixture = createLifecycleFixture();
    fixture.env.ANALYST_LENS_TEST_CERT_INSTALL_FAIL = "1";

    const installed = fixture.run("install");

    expect(installed.status).not.toBe(0);
    expect(installed.stderr).toContain("certificate install failed");
    expect(existsSync(fixture.appRoot)).toBe(false);
    expect(existsSync(fixture.plistPath)).toBe(false);
    expect(existsSync(fixture.manifestRegistration)).toBe(false);
  });

  it("rolls back locally when the XML registration path cannot be created", () => {
    const fixture = createLifecycleFixture();
    const registrationDirectory = dirname(fixture.manifestRegistration);
    mkdirSync(dirname(registrationDirectory), { recursive: true });
    writeFileSync(registrationDirectory, "foreign path blocker");

    const installed = fixture.run("install");

    expect(installed.status).not.toBe(0);
    expect(existsSync(fixture.appRoot)).toBe(false);
    expect(readFileSync(registrationDirectory, "utf8")).toBe(
      "foreign path blocker",
    );
    expect(readFileSync(fixture.commandRecord, "utf8")).not.toMatch(/atk </u);
  });

  it("removes its exact WEF link when ownership persistence fails", () => {
    const fixture = createLifecycleFixture();
    fixture.env.ANALYST_LENS_TEST_OWNERSHIP_WRITE_FAIL = "1";

    const installed = fixture.run("install");

    expect(installed.status).not.toBe(0);
    expect(existsSync(fixture.manifestRegistration)).toBe(false);
    expect(existsSync(fixture.appRoot)).toBe(false);
  });

  it("rejects changes to the persisted installer entrypoints", () => {
    const fixture = createLifecycleFixture();
    expect(fixture.run("install").status).toBe(0);
    appendFileSync(
      join(
        fixture.appRoot,
        "tooling",
        "node_modules",
        ".bin",
        "office-addin-dev-certs",
      ),
      "\n// changed\n",
    );

    expect(fixture.runInstalled("status").status).not.toBe(0);
  });

  it("rejects changes anywhere in the persisted production tool tree", () => {
    const fixture = createLifecycleFixture();
    expect(fixture.run("install").status).toBe(0);
    appendFileSync(
      join(
        fixture.appRoot,
        "tooling",
        "node_modules",
        "office-addin-dev-certs",
        "runtime.js",
      ),
      "// changed\n",
    );

    expect(fixture.runInstalled("status").status).not.toBe(0);
  });

  it("rejects changed installed manifest content", () => {
    const changedManifest = createLifecycleFixture();
    expect(changedManifest.run("install").status).toBe(0);
    appendFileSync(join(changedManifest.appRoot, "manifest.xml"), "\n");
    expect(changedManifest.runInstalled("status").status).not.toBe(0);
  });

  it("rejects a copied registration with a different inode", () => {
    const changedInode = createLifecycleFixture();
    expect(changedInode.run("install").status).toBe(0);
    const registeredManifest = readFileSync(changedInode.manifestRegistration);
    rmSync(changedInode.manifestRegistration);
    writeFileSync(changedInode.manifestRegistration, registeredManifest);
    expect(changedInode.runInstalled("status").status).not.toBe(0);
  });

  it("rejects registration metadata that points outside the owned XML path", () => {
    const changedOwner = createLifecycleFixture();
    expect(changedOwner.run("install").status).toBe(0);
    const ownerPath = join(changedOwner.appRoot, "registration.owner");
    const owner = readFileSync(ownerPath, "utf8").replace(
      changedOwner.manifestRegistration,
      `${changedOwner.manifestRegistration}.foreign`,
    );
    writeFileSync(ownerPath, owner);
    expect(changedOwner.runInstalled("status").status).not.toBe(0);
  });

  it("refuses uninstall when ownership metadata is corrupt or registration ownership is unbound", () => {
    const corruptMetadata = createLifecycleFixture();
    expect(corruptMetadata.run("install").status).toBe(0);
    writeFileSync(
      join(corruptMetadata.appRoot, "registration.owner"),
      "corrupt\n",
    );

    const corruptUninstall = corruptMetadata.runInstalled("uninstall");

    expect(corruptUninstall.status).not.toBe(0);
    expect(corruptUninstall.stderr).toMatch(/ownership|metadata/iu);
    expect(existsSync(corruptMetadata.appRoot)).toBe(true);
    expect(existsSync(corruptMetadata.plistPath)).toBe(true);
    expect(existsSync(corruptMetadata.manifestRegistration)).toBe(true);

    const unboundRegistration = createLifecycleFixture();
    expect(unboundRegistration.run("install").status).toBe(0);
    rmSync(unboundRegistration.manifestRegistration);
    writeFileSync(unboundRegistration.manifestRegistration, "foreign manifest");

    const unboundUninstall = unboundRegistration.runInstalled("uninstall");

    expect(unboundUninstall.status).not.toBe(0);
    expect(unboundUninstall.stderr).toMatch(/ownership|not bound/iu);
    expect(readFileSync(unboundRegistration.manifestRegistration, "utf8")).toBe(
      "foreign manifest",
    );
    expect(existsSync(unboundRegistration.appRoot)).toBe(true);
    expect(existsSync(unboundRegistration.plistPath)).toBe(true);
  });

  it("contains no Excel macro-security mutation", () => {
    const cli = readFileSync(cliPath, "utf8");

    expect(cli).not.toMatch(/VBAWarnings|macro.?security|Trust Center/i);
  });
});
