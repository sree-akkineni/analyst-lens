import {
  readFileSync,
  realpathSync,
  statSync,
} from "node:fs";
import type { IncomingMessage, ServerResponse } from "node:http";
import { createServer as createHttpsServer } from "node:https";
import { homedir } from "node:os";
import { basename, extname, isAbsolute, join, relative, resolve } from "node:path";
import { pathToFileURL } from "node:url";

export type HostOptions = {
  assetDir: string;
  version: string;
};

export type StartOptions = {
  certFile?: string;
  keyFile?: string;
  port?: number;
};

type RuntimeEnvironment = Readonly<Record<string, string | undefined>>;

const requiredEnvironmentValue = (
  environment: RuntimeEnvironment,
  name: string,
) => {
  const value = environment[name];
  if (value === undefined || value === "") {
    throw new Error(`Missing required environment variable ${name}`);
  }
  return value;
};

export const runtimeOptionsFromEnvironment = (
  environment: RuntimeEnvironment,
) => ({
  assetDir: requiredEnvironmentValue(environment, "ANALYST_LENS_ASSET_DIR"),
  certFile: requiredEnvironmentValue(environment, "ANALYST_LENS_CERT_FILE"),
  keyFile: requiredEnvironmentValue(environment, "ANALYST_LENS_KEY_FILE"),
  version: requiredEnvironmentValue(environment, "ANALYST_LENS_VERSION"),
});

const mimeTypes: Readonly<Record<string, string>> = {
  ".css": "text/css; charset=utf-8",
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml; charset=utf-8",
};

const contentSecurityPolicy = [
  "default-src 'self'",
  "script-src 'self' https://appsforoffice.microsoft.com/lib/1/hosted/ https://ajax.aspnetcdn.com/ajax/3.5/MicrosoftAjax.js",
  "connect-src 'self' https://127.0.0.1:43119",
  "img-src 'self' data:",
  "style-src 'self'",
  "object-src 'none'",
  "base-uri 'none'",
].join("; ");

const isLoopbackHost = (hostHeader: string | undefined) => {
  if (hostHeader === undefined) {
    return false;
  }
  const match = /^(127\.0\.0\.1|localhost|\[::1\])(?::([0-9]{1,5}))?$/iu.exec(
    hostHeader,
  );
  if (match === null) {
    return false;
  }
  const port = match[2];
  return port === undefined || Number(port) <= 65_535;
};

const end = (response: ServerResponse, statusCode: number) => {
  response.statusCode = statusCode;
  response.end();
};

const isWithin = (root: string, candidate: string) => {
  const pathFromRoot = relative(root, candidate);
  return (
    pathFromRoot === "" ||
    (!pathFromRoot.startsWith("..") && !isAbsolute(pathFromRoot))
  );
};

export const createHost = ({ assetDir, version }: HostOptions) => {
  const assetRoot = realpathSync(assetDir);
  const bindHost = "127.0.0.1" as const;

  const handleRequest = (request: IncomingMessage, response: ServerResponse) => {
    if (!isLoopbackHost(request.headers.host)) {
      end(response, 421);
      return;
    }
    if (request.method !== "GET") {
      response.setHeader("Allow", "GET");
      end(response, 405);
      return;
    }

    response.setHeader("Content-Security-Policy", contentSecurityPolicy);

    const requestUrl = new URL(request.url ?? "/", "http://127.0.0.1");
    if (requestUrl.pathname === "/health") {
      const body = JSON.stringify({ name: "Analyst Lens", version });
      response.statusCode = 200;
      response.setHeader("Cache-Control", "no-store");
      response.setHeader("Content-Type", "application/json; charset=utf-8");
      response.end(body);
      return;
    }

    let pathname: string;
    try {
      pathname = decodeURIComponent(requestUrl.pathname);
    } catch {
      end(response, 404);
      return;
    }
    const candidate = resolve(assetRoot, `.${pathname}`);
    if (!isWithin(assetRoot, candidate)) {
      end(response, 404);
      return;
    }

    try {
      const realCandidate = realpathSync(candidate);
      if (!isWithin(assetRoot, realCandidate) || !statSync(realCandidate).isFile()) {
        end(response, 404);
        return;
      }
      response.statusCode = 200;
      const assetName = basename(realCandidate);
      const isContentHashed = /[-.][A-Za-z0-9_-]{8,}\./u.test(
        assetName,
      );
      response.setHeader(
        "Cache-Control",
        isContentHashed
          ? "public, max-age=31536000, immutable"
          : "no-store",
      );
      response.setHeader("X-Content-Type-Options", "nosniff");
      response.setHeader(
        "Content-Type",
        mimeTypes[extname(realCandidate).toLowerCase()] ??
          "application/octet-stream",
      );
      response.end(readFileSync(realCandidate));
    } catch {
      end(response, 404);
    }
  };

  const start = ({
    certFile =
      process.env.ANALYST_LENS_CERT_FILE ??
      join(homedir(), ".office-addin-dev-certs", "localhost.crt"),
    keyFile =
      process.env.ANALYST_LENS_KEY_FILE ??
      join(homedir(), ".office-addin-dev-certs", "localhost.key"),
    port = 43119,
  }: StartOptions = {}) =>
    createHttpsServer(
      {
        cert: readFileSync(certFile),
        key: readFileSync(keyFile),
      },
      handleRequest,
    ).listen(port, bindHost);

  return {
    bindHost,
    handleRequest,
    routes: () => ["GET /health", "GET /*asset"] as const,
    start,
  };
};

const entryPoint = process.argv[1];
if (
  entryPoint !== undefined &&
  pathToFileURL(resolve(entryPoint)).href === import.meta.url
) {
  const { assetDir, certFile, keyFile, version } =
    runtimeOptionsFromEnvironment(process.env);
  const server = createHost({ assetDir, version }).start({ certFile, keyFile });
  server.once("listening", () => {
    console.log("Analyst Lens host listening on https://127.0.0.1:43119");
  });
}
