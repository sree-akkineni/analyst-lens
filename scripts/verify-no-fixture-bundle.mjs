#!/usr/bin/env node

import {
  existsSync,
  readFileSync,
  readdirSync,
  statSync,
} from "node:fs";
import { extname, join, relative, resolve } from "node:path";

const root = resolve(process.argv[2] ?? "dist");
const textExtensions = new Set([".css", ".html", ".js"]);
const fixtureMarkers = ["sample-model", "simulatedWorkbook", "quickTraceFixture"];
const failures = [];

const scan = (path) => {
  if (statSync(path).isDirectory()) {
    readdirSync(path).forEach((entry) => scan(join(path, entry)));
    return;
  }
  const relativePath = relative(root, path);
  if (fixtureMarkers.some((marker) => relativePath.includes(marker))) {
    failures.push(relativePath);
  }
  if (!textExtensions.has(extname(path))) return;
  const source = readFileSync(path, "utf8");
  if (fixtureMarkers.some((marker) => source.includes(marker))) {
    failures.push(relativePath);
  }
};

if (!existsSync(root)) {
  console.error(`production bundle is missing: ${root}`);
  process.exit(1);
}

scan(root);

if (failures.length > 0) {
  console.error(
    `production bundle contains standalone fixture code: ${[...new Set(failures)].sort().join(", ")}`,
  );
  process.exit(1);
}

console.log("production bundle excludes standalone fixture code");
