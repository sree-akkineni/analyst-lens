#!/usr/bin/env node

import { existsSync, readFileSync, readdirSync } from "node:fs";
import { extname, join, relative, resolve, sep } from "node:path";

const root = resolve(process.argv[2] ?? ".");
const sourceRoot = join(root, "src");
const allowedRoot = join(sourceRoot, "platform", "office") + sep;
const sourceExtensions = new Set([".ts", ".tsx"]);
const officeCallPattern = /\b(?:Excel\s*\.\s*run|Office\s*\.)/u;
const violations = [];

const scan = (path) => {
  for (const entry of readdirSync(path, { withFileTypes: true })) {
    const candidate = join(path, entry.name);
    if (entry.isDirectory()) {
      scan(candidate);
      continue;
    }
    if (
      !sourceExtensions.has(extname(candidate)) ||
      candidate.startsWith(allowedRoot)
    ) {
      continue;
    }
    if (officeCallPattern.test(readFileSync(candidate, "utf8"))) {
      violations.push(relative(root, candidate));
    }
  }
};

if (!existsSync(sourceRoot)) {
  console.error(`missing source root: ${sourceRoot}`);
  process.exit(1);
}

scan(sourceRoot);

if (violations.length > 0) {
  for (const path of violations.sort()) {
    console.error(`Office boundary violation: ${path}`);
  }
  process.exit(1);
}

console.log(
  "Office boundary ok: src/platform/office is the only Office.js caller",
);
