import react from "@vitejs/plugin-react";
import { defineConfig, type Plugin } from "vite";

export const analystLensFixturesEnabled =
  process.env.ANALYST_LENS_FIXTURES === "1";

export const taskpaneHtml = `<!doctype html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>Analyst Lens</title>
    <link rel="stylesheet" href="/assets/taskpane.css" />
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
  </head>
  <body>
    <div id="root"></div>
    <script type="module" src="/assets/taskpane.js"></script>
  </body>
</html>
`;

export const quickTraceHtml = `<!doctype html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>Analyst Lens Quick Trace</title>
    <link rel="stylesheet" href="/assets/quick-trace.css" />
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
  </head>
  <body>
    <div id="quick-trace-root"></div>
    <script type="module" src="/assets/quick-trace.js"></script>
  </body>
</html>`;

export const taskpaneAssetFileName = (
  assetInfo: { names?: string[] },
): string => assetInfo.names?.some((name) => name.endsWith(".css")) === true
  ? "assets/taskpane.css"
  : "assets/[name]-[hash][extname]";

export const quickTraceAssetFileName = (
  assetInfo: { names?: string[] },
): string => assetInfo.names?.some((name) => name.endsWith("quickTrace.css")) === true
  ? "assets/quick-trace.css"
  : "assets/[name]-[hash][extname]";

export const analystLensAssetFileName = (
  assetInfo: { names?: string[] },
): string => assetInfo.names?.some((name) => name.endsWith("quickTrace.css")) === true
  ? quickTraceAssetFileName(assetInfo)
  : taskpaneAssetFileName(assetInfo);

export const analystLensEntryFileName = (
  chunkInfo: { name: string },
): string => chunkInfo.name === "quickTrace"
  ? "assets/quick-trace.js"
  : chunkInfo.name === "taskpane"
    ? "assets/taskpane.js"
    : "assets/[name]-[hash].js";

const emitAnalystLensShells: Plugin = {
  name: "emit-analyst-lens-shells",
  generateBundle() {
    this.emitFile({ type: "asset", fileName: "taskpane.html", source: taskpaneHtml });
    this.emitFile({ type: "asset", fileName: "quick-trace.html", source: quickTraceHtml });
  },
};

export default defineConfig({
  define: {
    __ANALYST_LENS_FIXTURES__: JSON.stringify(analystLensFixturesEnabled),
  },
  plugins: [react(), emitAnalystLensShells],
  build: {
    emptyOutDir: true,
    rollupOptions: {
      input: { quickTrace: "src/quickTrace.tsx", taskpane: "src/taskpane.tsx" },
      output: {
        assetFileNames: analystLensAssetFileName,
        entryFileNames: analystLensEntryFileName,
      },
    },
  },
});
