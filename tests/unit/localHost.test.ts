import { mkdtempSync, rmSync, symlinkSync, writeFileSync } from "node:fs";
import {
  createServer,
  type IncomingMessage,
  request as httpRequest,
  type Server,
  type ServerResponse,
} from "node:http";
import { request as httpsRequest } from "node:https";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { spawnSync } from "node:child_process";
import { afterEach, describe, expect, it } from "vitest";
import {
  createHost,
  runtimeOptionsFromEnvironment,
} from "../../local-host/server";

const fixtureDirs: string[] = [];
const servers: Server[] = [];

const createFixtureDir = () => {
  const fixtureDir = mkdtempSync(join(tmpdir(), "analyst-lens-host-"));
  fixtureDirs.push(fixtureDir);
  return fixtureDir;
};

afterEach(() => {
  for (const server of servers.splice(0)) {
    server.close();
  }
  for (const fixtureDir of fixtureDirs.splice(0)) {
    rmSync(fixtureDir, { force: true, recursive: true });
  }
});

const listen = async (server: Server) => {
  servers.push(server);
  await new Promise<void>((resolve, reject) => {
    server.once("error", reject);
    server.listen(0, "127.0.0.1", resolve);
  });
  const address = server.address();
  if (address === null || typeof address === "string") {
    throw new Error("Expected an ephemeral TCP address");
  }
  return `http://127.0.0.1:${address.port}`;
};

const listenFor = async (app: ReturnType<typeof createHost>) => {
  const { handleRequest } = app as typeof app & {
    handleRequest?: (request: IncomingMessage, response: ServerResponse) => void;
  };
  if (handleRequest === undefined) {
    throw new Error("createHost must expose a request handler");
  }
  return listen(createServer(handleRequest));
};

const request = async (
  baseUrl: string,
  path: string,
  init?: RequestInit,
) => fetch(`${baseUrl}${path}`, init);

const requestWithHost = async (baseUrl: string, hostHeader: string) => {
  const target = new URL(baseUrl);
  return new Promise<number>((resolve, reject) => {
    const outgoing = httpRequest(
      {
        headers: { Host: hostHeader },
        hostname: target.hostname,
        method: "GET",
        path: "/health",
        port: target.port,
      },
      (incoming) => {
        incoming.resume();
        resolve(incoming.statusCode ?? 0);
      },
    );
    outgoing.once("error", reject);
    outgoing.end();
  });
};

describe("local HTTPS host", () => {
  it("reads the launchd runtime contract from explicit environment paths", () => {
    expect(
      runtimeOptionsFromEnvironment({
        ANALYST_LENS_ASSET_DIR: "/tmp/Analyst Lens/assets",
        ANALYST_LENS_CERT_FILE: "/tmp/Analyst Lens/localhost.crt",
        ANALYST_LENS_KEY_FILE: "/tmp/Analyst Lens/localhost.key",
        ANALYST_LENS_VERSION: "0.1.0",
      }),
    ).toEqual({
      assetDir: "/tmp/Analyst Lens/assets",
      certFile: "/tmp/Analyst Lens/localhost.crt",
      keyFile: "/tmp/Analyst Lens/localhost.key",
      version: "0.1.0",
    });
  });

  it("binds only to loopback and exposes no workbook endpoint", () => {
    const app = createHost({ assetDir: createFixtureDir(), version: "0.1.0" });

    expect(app.routes()).toEqual(["GET /health", "GET /*asset"]);
    expect(app.bindHost).toBe("127.0.0.1");
  });

  it("starts an HTTPS listener on the production loopback port", async () => {
    const fixtureDir = createFixtureDir();
    const certFile = join(fixtureDir, "localhost.crt");
    const keyFile = join(fixtureDir, "localhost.key");
    const certificate = spawnSync(
      "/usr/bin/openssl",
      [
        "req",
        "-x509",
        "-newkey",
        "rsa:2048",
        "-nodes",
        "-keyout",
        keyFile,
        "-out",
        certFile,
        "-subj",
        "/CN=localhost",
        "-days",
        "1",
      ],
      { encoding: "utf8" },
    );
    expect(certificate.status, certificate.stderr).toBe(0);

    const app = createHost({ assetDir: fixtureDir, version: "0.1.0" });
    const { start } = app as typeof app & {
      start?: (options: {
        certFile: string;
        keyFile: string;
        port: number;
      }) => Server;
    };
    if (start === undefined) {
      throw new Error("createHost must expose an HTTPS start function");
    }
    const server = start({ certFile, keyFile, port: 0 });
    servers.push(server);
    await new Promise<void>((resolve, reject) => {
      server.once("error", reject);
      server.once("listening", resolve);
    });
    const address = server.address();
    if (address === null || typeof address === "string") {
      throw new Error("Expected an ephemeral HTTPS address");
    }

    const health = await new Promise<{ status: number; body: string }>(
      (resolve, reject) => {
        const outgoing = httpsRequest(
          {
            hostname: "127.0.0.1",
            path: "/health",
            port: address.port,
            rejectUnauthorized: false,
          },
          (incoming) => {
            const chunks: Buffer[] = [];
            incoming.on("data", (chunk: Buffer) => chunks.push(chunk));
            incoming.once("end", () =>
              resolve({
                body: Buffer.concat(chunks).toString("utf8"),
                status: incoming.statusCode ?? 0,
              }),
            );
          },
        );
        outgoing.once("error", reject);
        outgoing.end();
      },
    );

    expect(health.status).toBe(200);
    expect(JSON.parse(health.body)).toEqual({
      name: "Analyst Lens",
      version: "0.1.0",
    });
  });

  it("serves version health only for a loopback Host header", async () => {
    const app = createHost({ assetDir: createFixtureDir(), version: "0.1.0" });
    const baseUrl = await listenFor(app);

    const response = await request(baseUrl, "/health");

    expect(response.status).toBe(200);
    expect(response.headers.get("cache-control")).toBe("no-store");
    await expect(response.json()).resolves.toEqual({
      name: "Analyst Lens",
      version: "0.1.0",
    });
  });

  it("rejects a non-loopback Host header", async () => {
    const app = createHost({ assetDir: createFixtureDir(), version: "0.1.0" });
    const baseUrl = await listenFor(app);

    const status = await requestWithHost(baseUrl, "attacker.example");

    expect(status).toBe(421);
  });

  it("serves content-hashed files as immutable from the asset root", async () => {
    const assetDir = createFixtureDir();
    writeFileSync(join(assetDir, "taskpane-ZxYwV987.js"), "void 0;");
    const app = createHost({ assetDir, version: "0.1.0" });
    const baseUrl = await listenFor(app);

    const response = await request(baseUrl, "/taskpane-ZxYwV987.js");

    expect(response.status).toBe(200);
    expect(response.headers.get("cache-control")).toBe(
      "public, max-age=31536000, immutable",
    );
    expect(response.headers.get("x-content-type-options")).toBe("nosniff");
    await expect(response.text()).resolves.toBe("void 0;");
  });

  it("never caches stable HTML entrypoints", async () => {
    const assetDir = createFixtureDir();
    writeFileSync(join(assetDir, "taskpane.html"), "<h1>Analyst Lens</h1>");
    writeFileSync(join(assetDir, "quick-trace.html"), "<h1>Quick Trace</h1>");
    const app = createHost({ assetDir, version: "0.1.0" });
    const baseUrl = await listenFor(app);

    const [taskpane, quickTrace] = await Promise.all([
      request(baseUrl, "/taskpane.html"),
      request(baseUrl, "/quick-trace.html"),
    ]);

    expect(taskpane.status).toBe(200);
    expect(taskpane.headers.get("cache-control")).toBe("no-store");
    expect(quickTrace.status).toBe(200);
    expect(quickTrace.headers.get("cache-control")).toBe("no-store");
  });

  it("blocks browser connections outside the local host", async () => {
    const assetDir = createFixtureDir();
    writeFileSync(join(assetDir, "taskpane.html"), "<h1>Analyst Lens</h1>");
    writeFileSync(join(assetDir, "quick-trace.html"), "<h1>Quick Trace</h1>");
    const app = createHost({ assetDir, version: "0.1.0" });
    const baseUrl = await listenFor(app);

    const [taskpane, quickTrace] = await Promise.all([
      request(baseUrl, "/taskpane.html"),
      request(baseUrl, "/quick-trace.html"),
    ]);

    const expectedPolicy =
      "default-src 'self'; script-src 'self' https://appsforoffice.microsoft.com/lib/1/hosted/ https://ajax.aspnetcdn.com/ajax/3.5/MicrosoftAjax.js; connect-src 'self' https://127.0.0.1:43119; img-src 'self' data:; style-src 'self'; object-src 'none'; base-uri 'none'";
    expect(taskpane.headers.get("content-security-policy")).toBe(expectedPolicy);
    expect(quickTrace.headers.get("content-security-policy")).toBe(expectedPolicy);
  });

  it("rejects non-GET methods, missing assets, and traversal attempts", async () => {
    const assetDir = createFixtureDir();
    const app = createHost({ assetDir, version: "0.1.0" });
    const baseUrl = await listenFor(app);

    const [post, missing, traversal] = await Promise.all([
      request(baseUrl, "/health", { method: "POST" }),
      request(baseUrl, "/workbook"),
      request(baseUrl, "/..%2Fpackage.json"),
    ]);

    expect(post.status).toBe(405);
    expect(post.headers.get("allow")).toBe("GET");
    expect(missing.status).toBe(404);
    expect(traversal.status).toBe(404);
  });

  it("returns 404 for a real asset symlink that escapes the asset root", async () => {
    const assetDir = createFixtureDir();
    const outsideDir = createFixtureDir();
    const outsideAsset = join(outsideDir, "secret.txt");
    writeFileSync(outsideAsset, "must not be served");
    symlinkSync(outsideAsset, join(assetDir, "escape.txt"));
    const app = createHost({ assetDir, version: "0.1.0" });
    const baseUrl = await listenFor(app);

    const response = await request(baseUrl, "/escape.txt");

    expect(response.status).toBe(404);
  });
});
