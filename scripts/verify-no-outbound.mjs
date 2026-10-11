#!/usr/bin/env node

import {
  existsSync,
  readFileSync,
  readdirSync,
  statSync,
} from "node:fs";
import { extname, join, relative, resolve } from "node:path";

const root = resolve(process.argv[2] ?? ".");
const officeJsUrl =
  "https://appsforoffice.microsoft.com/lib/1/hosted/office.js";
const officeRuntimeRoot =
  "https://appsforoffice.microsoft.com/lib/1/hosted/";
const microsoftAjaxUrl =
  "https://ajax.aspnetcdn.com/ajax/3.5/MicrosoftAjax.js";
const allowedOfficeDestinations = new Set([
  officeJsUrl,
  officeRuntimeRoot,
  microsoftAjaxUrl,
]);
const sourceExtensions = new Set([
  ".css",
  ".html",
  ".js",
  ".json",
  ".mjs",
  ".ts",
  ".tsx",
  ".xml",
]);
const urlPattern = /https?:\/\/[^\s"'`<>\\)]+/gu;
const httpScheme = "http:";
const httpsScheme = "https:";
const generatedInertMetadataUrls = new Set([
  `${httpScheme}//www.w3.org/1998/Math/MathML`,
  `${httpScheme}//www.w3.org/1999/xlink`,
  `${httpScheme}//www.w3.org/2000/svg`,
  `${httpScheme}//www.w3.org/XML/1998/namespace`,
  `${httpScheme}//json-schema.org/draft-04/schema#`,
  `${httpScheme}//json-schema.org/draft-07/schema#`,
  `${httpsScheme}//json-schema.org/draft/2020-12/schema`,
  `${httpsScheme}//react.dev/errors/`,
]);
const isGeneratedInertMetadataUrl = (destination) =>
  generatedInertMetadataUrls.has(destination) ||
  /^https?:\/\/\[\$\{[^}]+\}\]$/u.test(destination);
const generatedNetworkPatterns = [
  /\b(?:fetch|sendBeacon|EventSource|WebSocket)\s*(?:\?\.)?\s*\(\s*["'`](https?:\/\/[^\s"'`<>\\)]+)["'`]/gu,
  /\.open\s*(?:\?\.)?\s*\(\s*["'`][^"'`]*["'`]\s*,\s*["'`](https?:\/\/[^\s"'`<>\\)]+)["'`]/gu,
  /\b(?:importScripts|open|SharedWorker|Worker|location\.(?:assign|replace))\s*(?:\?\.)?\s*\(\s*["'`](https?:\/\/[^\s"'`<>\\)]+)["'`]/gu,
  /\.(?:href|src)\s*=\s*["'`](https?:\/\/[^\s"'`<>\\)]+)["'`]/gu,
];
const generatedVariableSinkPatterns = [
  /\b(?:fetch|sendBeacon|EventSource|WebSocket)\s*(?:\?\.)?\s*\(\s*([A-Za-z_$][\w$]*)\b/gu,
  /\.open\s*(?:\?\.)?\s*\(\s*["'`][^"'`]*["'`]\s*,\s*([A-Za-z_$][\w$]*)\b/gu,
  /\b(?:importScripts|open|SharedWorker|Worker|location\.(?:assign|replace))\s*(?:\?\.)?\s*\(\s*([A-Za-z_$][\w$]*)\b/gu,
  /\.(?:href|src)\s*=\s*([A-Za-z_$][\w$]*)\b/gu,
];
const destinations = [];

const escapeRegularExpression = (value) =>
  value.replace(/[.*+?^${}()|[\]\\]/gu, "\\$&");

const recordDestination = (value, sourcePath) => {
  destinations.push({ destination: value, sourcePath });
};

const collectUrls = (value, sourcePath, key = "") => {
  if (key === "$schema") {
    return;
  }
  if (typeof value === "string") {
    if (/^https?:\/\//u.test(value)) {
      recordDestination(value, sourcePath);
    }
    return;
  }
  if (Array.isArray(value)) {
    for (const item of value) {
      collectUrls(item, sourcePath);
    }
    return;
  }
  if (value !== null && typeof value === "object") {
    for (const [childKey, child] of Object.entries(value)) {
      collectUrls(child, sourcePath, childKey);
    }
  }
};

const scanFile = (path, generatedBundle = false) => {
  const sourcePath = relative(root, path);
  const fileSource = readFileSync(path, "utf8");
  const source = extname(path) === ".xml"
    ? fileSource.replace(/\sxmlns(?::[A-Za-z0-9_-]+)?="[^"]+"/gu, "")
    : fileSource;
  if (generatedBundle && extname(path) === ".js") {
    for (const pattern of generatedNetworkPatterns) {
      for (const match of source.matchAll(pattern)) {
        recordDestination(match[1], sourcePath);
      }
    }
    for (const pattern of generatedVariableSinkPatterns) {
      for (const match of source.matchAll(pattern)) {
        const identifier = match[1];
        const sinkIndex = match.index;
        if (identifier === undefined || sinkIndex === undefined) continue;
        const localSource = source.slice(Math.max(0, sinkIndex - 512), sinkIndex);
        const assignmentPattern = new RegExp(
          `\\b${escapeRegularExpression(identifier)}\\s*=\\s*["'\`](https?:\\/\\/[^\\s"'\`<>\\\\)]+)["'\`]`,
          "gu",
        );
        let destination;
        for (const assignment of localSource.matchAll(assignmentPattern)) {
          destination = assignment[1];
        }
        if (
          destination !== undefined &&
          !isGeneratedInertMetadataUrl(destination)
        ) {
          recordDestination(destination, sourcePath);
        }
      }
    }
    for (const match of source.matchAll(urlPattern)) {
      const destination = match[0];
      if (!isGeneratedInertMetadataUrl(destination)) {
        recordDestination(destination, sourcePath);
      }
    }
    return;
  }
  for (const match of source.matchAll(urlPattern)) {
    recordDestination(match[0], sourcePath);
  }
};

const scanPath = (
  path,
  scanExtensionless = false,
  generatedBundle = false,
) => {
  if (!existsSync(path)) {
    return;
  }
  if (statSync(path).isDirectory()) {
    for (const entry of readdirSync(path, { withFileTypes: true })) {
      if (
        entry.name === "node_modules" ||
        entry.name === "package-lock.json" ||
        entry.name === "npm-shrinkwrap.json"
      ) {
        continue;
      }
      scanPath(join(path, entry.name), scanExtensionless, generatedBundle);
    }
    return;
  }
  if (sourceExtensions.has(extname(path)) || (scanExtensionless && extname(path) === "")) {
    scanFile(path, generatedBundle);
  }
};

const manifestPath = join(root, "manifest.json");
if (!existsSync(manifestPath)) {
  console.error(`missing manifest: ${manifestPath}`);
  process.exit(1);
}
collectUrls(
  JSON.parse(readFileSync(manifestPath, "utf8")),
  "manifest.json",
);

for (const [relativePath, scanExtensionless, generatedBundle] of [
  ["src", false, false],
  ["local-host", false, false],
  ["public", false, false],
  ["dist", false, true],
  ["dist-host", false, false],
  ["scripts", true, false],
  ["installer-tools", true, false],
  ["manifest.xml", false, false],
  ["package.json", false, false],
  ["vite.config.ts", false, false],
]) {
  scanPath(join(root, relativePath), scanExtensionless, generatedBundle);
}

const failures = new Set();
for (const { destination } of destinations) {
  let url;
  try {
    url = new URL(destination);
  } catch {
    failures.add(`unexpected outbound destination: ${destination}`);
    continue;
  }

  const isLoopback =
    url.hostname === "127.0.0.1" ||
    url.hostname === "localhost" ||
    url.hostname === "[::1]";
  if (isLoopback || allowedOfficeDestinations.has(url.href)) {
    continue;
  }

  failures.add(`unexpected outbound destination: ${destination}`);
}

if (failures.size > 0) {
  for (const failure of [...failures].sort()) {
    console.error(failure);
  }
  process.exit(1);
}

console.log(
  "network allowlist ok: 127.0.0.1, appsforoffice.microsoft.com/lib/1/hosted/, ajax.aspnetcdn.com/ajax/3.5/MicrosoftAjax.js",
);
