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

const verifier = resolve("scripts/verify-no-outbound.mjs");
const roots: string[] = [];

const createFixture = (source = "") => {
  const root = mkdtempSync(join(tmpdir(), "analyst-lens-network-"));
  roots.push(root);
  mkdirSync(join(root, "src"), { recursive: true });
  mkdirSync(join(root, "scripts"), { recursive: true });
  mkdirSync(join(root, "dist-host"), { recursive: true });
  writeFileSync(
    join(root, "manifest.json"),
    JSON.stringify({
      $schema:
        "https://developer.microsoft.com/json-schemas/teams/v1.25/MicrosoftTeams.schema.json",
      extensions: [
        {
          runtimes: [
            { code: { page: "https://127.0.0.1:43119/taskpane.html" } },
          ],
        },
      ],
    }),
  );
  writeFileSync(
    join(root, "manifest.xml"),
    '<OfficeApp><SourceLocation DefaultValue="https://127.0.0.1:43119/taskpane.html" /></OfficeApp>',
  );
  writeFileSync(join(root, "src", "runtime.ts"), source);
  writeFileSync(
    join(root, "package.json"),
    JSON.stringify({ name: "analyst-lens-fixture", scripts: {} }),
  );
  return root;
};

afterEach(() => {
  for (const root of roots.splice(0)) {
    rmSync(root, { force: true, recursive: true });
  }
});

describe("network allowlist verifier", () => {
  it("permits only loopback and the Microsoft Office.js runtime dependencies", () => {
    const root = createFixture(
      [
        'const officeJs = "https://appsforoffice.microsoft.com/lib/1/hosted/office.js";',
        'const officeRuntimeRoot = "https://appsforoffice.microsoft.com/lib/1/hosted/";',
        'const microsoftAjax = "https://ajax.aspnetcdn.com/ajax/3.5/MicrosoftAjax.js";',
        'const quickTrace = "https://127.0.0.1:43119/quick-trace.html";',
      ].join("\n"),
    );

    const result = spawnSync(process.execPath, [verifier, root], {
      encoding: "utf8",
    });

    expect(result.status, result.stderr).toBe(0);
    expect(result.stdout.trim()).toBe(
      "network allowlist ok: 127.0.0.1, appsforoffice.microsoft.com/lib/1/hosted/, ajax.aspnetcdn.com/ajax/3.5/MicrosoftAjax.js",
    );
  });

  it("rejects an unapproved path on the Office.js CDN", () => {
    const root = createFixture(
      'const unexpected = "https://appsforoffice.microsoft.com/unapproved.js";',
    );

    const result = spawnSync(process.execPath, [verifier, root], {
      encoding: "utf8",
    });

    expect(result.status).not.toBe(0);
    expect(result.stderr).toContain(
      "unexpected outbound destination: https://appsforoffice.microsoft.com/unapproved.js",
    );
  });

  it("fails closed on an unexpected HTTP destination", () => {
    const root = createFixture(
      'void fetch("https://telemetry.example.com/workbook");',
    );

    const result = spawnSync(process.execPath, [verifier, root], {
      encoding: "utf8",
    });

    expect(result.status).not.toBe(0);
    expect(result.stderr).toContain(
      "unexpected outbound destination: https://telemetry.example.com/workbook",
    );
  });

  it("distinguishes inert dependency metadata from a generated network call", () => {
    const root = createFixture();
    const distRoot = join(root, "dist");
    mkdirSync(distRoot, { recursive: true });
    writeFileSync(
      join(distRoot, "taskpane.js"),
      'const xmlNamespace="http://www.w3.org/2000/svg"; const reactHelp="https://react.dev/errors/";',
    );

    const inertMetadata = spawnSync(process.execPath, [verifier, root], {
      encoding: "utf8",
    });

    expect(inertMetadata.status, inertMetadata.stderr).toBe(0);

    writeFileSync(
      join(distRoot, "taskpane.js"),
      'void fetch("https://react.dev/errors/418");',
    );

    const networkCall = spawnSync(process.execPath, [verifier, root], {
      encoding: "utf8",
    });

    expect(networkCall.status).not.toBe(0);
    expect(networkCall.stderr).toContain(
      "unexpected outbound destination: https://react.dev/errors/418",
    );
  });

  it("distinguishes Zod schema metadata from a generated network call", () => {
    const root = createFixture();
    const distRoot = join(root, "dist");
    mkdirSync(distRoot, { recursive: true });
    writeFileSync(
      join(distRoot, "quick-trace.js"),
      [
        'const draft4="http://json-schema.org/draft-04/schema#";',
        'const draft7="http://json-schema.org/draft-07/schema#";',
        'const draft2020="https://json-schema.org/draft/2020-12/schema";',
        'const ipv6Template=`http://[${value}]`;',
      ].join(""),
    );

    const inertMetadata = spawnSync(process.execPath, [verifier, root], {
      encoding: "utf8",
    });

    expect(inertMetadata.status, inertMetadata.stderr).toBe(0);

    writeFileSync(
      join(distRoot, "quick-trace.js"),
      'void fetch("https://json-schema.org/draft/2020-12/schema");',
    );

    const networkCall = spawnSync(process.execPath, [verifier, root], {
      encoding: "utf8",
    });

    expect(networkCall.status).not.toBe(0);
    expect(networkCall.stderr).toContain(
      "unexpected outbound destination: https://json-schema.org/draft/2020-12/schema",
    );
  });

  it("rejects optional-chained fetch calls in the generated JavaScript bundle", () => {
    const root = createFixture();
    const distRoot = join(root, "dist");
    mkdirSync(distRoot, { recursive: true });
    writeFileSync(
      join(distRoot, "taskpane.js"),
      'window.fetch?.("https://telemetry.example.com/optional");',
    );

    const result = spawnSync(process.execPath, [verifier, root], {
      encoding: "utf8",
    });

    expect(result.status).not.toBe(0);
    expect(result.stderr).toContain(
      "unexpected outbound destination: https://telemetry.example.com/optional",
    );
  });

  it("rejects an indirect outbound destination used by a generated network call", () => {
    const root = createFixture();
    const distRoot = join(root, "dist");
    mkdirSync(distRoot, { recursive: true });
    writeFileSync(
      join(distRoot, "taskpane.js"),
      'const endpoint="https://telemetry.example.com/indirect";fetch(endpoint);',
    );

    const result = spawnSync(process.execPath, [verifier, root], {
      encoding: "utf8",
    });

    expect(result.status).not.toBe(0);
    expect(result.stderr).toContain(
      "unexpected outbound destination: https://telemetry.example.com/indirect",
    );
  });

  it.each([
    [
      "member expression",
      'const config={endpoint:"https://telemetry.example.com/member"};fetch(config.endpoint);',
      "https://telemetry.example.com/member",
    ],
    [
      "URL constructor",
      'fetch(new URL("https://telemetry.example.com/constructed"));',
      "https://telemetry.example.com/constructed",
    ],
  ])("rejects a generated %s destination", (_kind, source, destination) => {
    const root = createFixture();
    const distRoot = join(root, "dist");
    mkdirSync(distRoot, { recursive: true });
    writeFileSync(join(distRoot, "taskpane.js"), source);

    const result = spawnSync(process.execPath, [verifier, root], {
      encoding: "utf8",
    });

    expect(result.status).not.toBe(0);
    expect(result.stderr).toContain(`unexpected outbound destination: ${destination}`);
  });

  it.each([
    ["HTML", "taskpane.html", '<img src="https://telemetry.example.com/pixel.png">'],
    ["CSS", "taskpane.css", 'body{background:url("https://telemetry.example.com/bg.png")}'],
  ])("rejects an outbound destination in built %s", (_kind, filename, contents) => {
    const root = createFixture();
    const distRoot = join(root, "dist");
    mkdirSync(distRoot, { recursive: true });
    writeFileSync(join(distRoot, filename), contents);

    const result = spawnSync(process.execPath, [verifier, root], {
      encoding: "utf8",
    });

    expect(result.status).not.toBe(0);
    expect(result.stderr).toContain("unexpected outbound destination");
  });

  it.each([
    ["shipped lifecycle scripts", "scripts/analyst-lens"],
    ["the shipped host bundle", "dist-host/server.js"],
    ["the Mac XML manifest", "manifest.xml"],
    ["package configuration", "package.json"],
  ])("scans %s for outbound destinations", (_surface, relativePath) => {
    const root = createFixture();
    const destination = `https://telemetry.example.com/${encodeURIComponent(relativePath)}`;
    const contents = relativePath === "package.json"
      ? JSON.stringify({
          name: "analyst-lens-fixture",
          scripts: { postinstall: `curl ${destination}` },
        })
      : `void fetch("${destination}");`;
    writeFileSync(join(root, relativePath), contents);

    const result = spawnSync(process.execPath, [verifier, root], {
      encoding: "utf8",
    });

    expect(result.status).not.toBe(0);
    expect(result.stderr).toContain(
      `unexpected outbound destination: ${destination}`,
    );
  });

  it("rejects Microsoft OAuth and title registration endpoints in lifecycle scripts", () => {
    const root = createFixture();
    const oauthUrl =
      "https://login.microsoftonline.com/common/oauth2/v2.0/authorize";
    const titleUrl =
      "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps";
    writeFileSync(
      join(root, "scripts", "analyst-lens"),
      `oauth_url="${oauthUrl}"\ntitle_url="${titleUrl}"\n`,
    );

    const installerOnly = spawnSync(process.execPath, [verifier, root], {
      encoding: "utf8",
    });

    expect(installerOnly.status).not.toBe(0);
    expect(installerOnly.stderr).toContain(
      `unexpected outbound destination: ${oauthUrl}`,
    );
    expect(installerOnly.stderr).toContain(
      `unexpected outbound destination: ${titleUrl}`,
    );
  });

  it("ignores registry metadata in installer lockfiles", () => {
    const root = createFixture();
    const toolingRoot = join(root, "installer-tools");
    mkdirSync(toolingRoot, { recursive: true });
    writeFileSync(
      join(toolingRoot, "package-lock.json"),
      JSON.stringify({
        packages: {
          "node_modules/example": {
            resolved: "https://registry.npmjs.org/example/-/example-1.0.0.tgz",
          },
        },
      }),
    );

    const result = spawnSync(process.execPath, [verifier, root], {
      encoding: "utf8",
    });

    expect(result.status, result.stderr).toBe(0);
  });
});
