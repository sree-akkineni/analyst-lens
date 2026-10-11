/**
 * Captures reproducible current-source UI evidence from the fixture bundle.
 *
 * Usage:
 *   node scripts/capture-ui-screenshots.mjs docs/qa/ui-modernization/current
 *
 * Requires a fixture bundle served at http://127.0.0.1:4173
 */
import { createHash } from "node:crypto";
import { execFileSync } from "node:child_process";
import { mkdir, readFile, rm, writeFile } from "node:fs/promises";
import { resolve } from "node:path";
import { chromium } from "playwright";

const outDir = process.argv[2] ?? "docs/qa/ui-modernization/current";
const expectedOutDir = resolve("docs/qa/ui-modernization/current");
if (resolve(outDir) !== expectedOutDir) {
  throw new Error(`Capture target must be ${expectedOutDir}`);
}

const base = "http://127.0.0.1:4173";
const sourceCommit = execFileSync("git", ["rev-parse", "HEAD"], { encoding: "utf8" }).trim();
const branch = execFileSync("git", ["branch", "--show-current"], { encoding: "utf8" }).trim();
const dirty = execFileSync("git", ["status", "--porcelain"], { encoding: "utf8" }).trim();
if (dirty !== "") {
  throw new Error("Current-source capture requires a clean checkout");
}

const cases = [
  { name: "taskpane-narrow-formula", width: 360, height: 900, url: "/taskpane.html?fixture=sample-model", ready: "text=DCF $F$42" },
  { name: "taskpane-narrow-hardcode", width: 360, height: 900, url: "/taskpane.html?fixture=sample-model&selection=hardcode", ready: "text=Inputs $B$7" },
  { name: "taskpane-wide-formula", width: 760, height: 900, url: "/taskpane.html?fixture=sample-model", ready: "text=DCF $F$42" },
  { name: "quick-trace-in-420x390", width: 420, height: 390, url: "/quick-trace.html?fixture=sample-model&direction=precedent", ready: "[role='tree']" },
  { name: "quick-trace-in-900x700", width: 900, height: 700, url: "/quick-trace.html?fixture=sample-model&direction=precedent", ready: "[role='tree']" },
  { name: "quick-trace-out-420x390", width: 420, height: 390, url: "/quick-trace.html?fixture=sample-model&direction=dependent", ready: "[role='tree']" },
];

await rm(outDir, { force: true, recursive: true });
for (const theme of ["light", "dark"]) {
  await mkdir(`${outDir}/${theme}`, { recursive: true });
}

const browser = await chromium.launch();
const captures = [];
for (const theme of ["light", "dark"]) {
  for (const item of cases) {
    const page = await browser.newPage({
      colorScheme: theme,
      viewport: { width: item.width, height: item.height },
    });
    const fixtureUrl = `${base}${item.url}`;
    await page.goto(fixtureUrl);
    await page.waitForSelector(item.ready, { state: "visible", timeout: 10000 });
    const overflow = await page.evaluate(() => Math.max(
      document.body.scrollWidth - document.body.clientWidth,
      document.documentElement.scrollWidth - document.documentElement.clientWidth,
    ));
    if (overflow > 0) {
      throw new Error(`${theme}/${item.name} overflows by ${overflow}px`);
    }

    const path = `${outDir}/${theme}/${item.name}.png`;
    await page.screenshot({ path });
    const sha256 = createHash("sha256").update(await readFile(path)).digest("hex");
    captures.push({
      filename: `${theme}/${item.name}.png`,
      fixtureUrl,
      name: item.name,
      sha256,
      theme,
      viewport: { width: item.width, height: item.height },
    });
    await page.close();
    console.log(`captured ${theme}/${item.name} ${item.width}x${item.height}`);
  }
}
await browser.close();

await writeFile(`${outDir}/manifest.json`, `${JSON.stringify({
  branch,
  captures,
  sourceCommit,
  status: "clean",
  timestamp: new Date().toISOString(),
}, null, 2)}\n`);
await writeFile(
  `${outDir}/manifest.sha256`,
  `${captures.map(({ filename, sha256 }) => `${sha256}  ${outDir}/${filename}`).sort().join("\n")}\n`,
);
