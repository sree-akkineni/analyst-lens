import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import { allCommands } from "../../src/domain/commands/commandRegistry";

type UnifiedManifest = {
  authorization: {
    permissions: {
      resourceSpecific: Array<{ name: string; type: string }>;
    };
  };
  icons: { color: string; outline: string };
  extensions: Array<{
    requirements: { scopes: string[] };
    runtimes: Array<{ lifetime: string; type: string }>;
    ribbons: Array<{
      tabs: Array<{
        groups: Array<{
          icons: Array<{ size: number; url: string }>;
        }>;
      }>;
    }>;
  }>;
};

type ShortcutsManifest = {
  actions: Array<{ id: string; name: string; type: string }>;
  shortcuts: Array<{
    action: string;
    key: { default: string; mac: string };
  }>;
};

const readManifest = () =>
  JSON.parse(readFileSync(resolve("manifest.json"), "utf8")) as UnifiedManifest;

const readMacManifest = () => readFileSync(resolve("manifest.xml"), "utf8");

const pngDimensions = (path: string) => {
  const png = readFileSync(resolve(path));
  expect(png.subarray(0, 8)).toEqual(
    Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]),
  );
  return { height: png.readUInt32BE(20), width: png.readUInt32BE(16) };
};

describe("unified Microsoft 365 manifest", () => {
  it("uses a long shared runtime for Excel on Mac", () => {
    const manifest = readManifest();

    expect(manifest.extensions[0]?.runtimes[0]).toMatchObject({
      lifetime: "long",
      type: "general",
    });
  });

  it("limits the extension to Excel and declares its document permission", () => {
    const manifest = readManifest();

    expect(manifest.extensions[0]?.requirements.scopes).toEqual(["workbook"]);
    expect(manifest.authorization.permissions.resourceSpecific).toEqual([
      { name: "Document.ReadWrite.User", type: "Delegated" },
    ]);
  });

  it("ships correctly sized package and ribbon PNG icons", () => {
    const manifest = readManifest();

    expect(manifest.icons).toEqual({
      color: "public/icons/icon-192.png",
      outline: "public/icons/icon-32.png",
    });
    for (const size of [16, 32, 80, 192]) {
      expect(pngDimensions(`public/icons/icon-${size}.png`)).toEqual({
        height: size,
        width: size,
      });
    }
    expect(manifest.extensions[0]?.ribbons[0]?.tabs[0]?.groups[0]?.icons).toEqual(
      [
        { size: 16, url: "https://127.0.0.1:43119/icons/icon-16.png" },
        { size: 32, url: "https://127.0.0.1:43119/icons/icon-32.png" },
        { size: 80, url: "https://127.0.0.1:43119/icons/icon-80.png" },
      ],
    );
  });
});

describe("add-in-only Mac manifest", () => {
  it("bumps the Office identity for the expanded Quick Trace shortcut catalog", () => {
    const manifest = readMacManifest();

    expect(manifest).toContain("<Version>1.0.1.0</Version>");
  });

  it("targets only Excel and grants document read/write access", () => {
    const manifest = readMacManifest();

    expect(manifest).toContain('xsi:type="TaskPaneApp"');
    expect(manifest).toContain("<Id>5db4c51e-9c38-4d8b-976b-2d0f707e27f5</Id>");
    expect(manifest.match(/<Host Name="Workbook"\s*\/>/gu)).toHaveLength(1);
    expect(manifest).not.toMatch(/<Host Name="(?:Document|Presentation)"/u);
    expect(manifest).toContain("<Permissions>ReadWriteDocument</Permissions>");
  });

  it("uses one long shared runtime for the task pane, commands, and shortcuts", () => {
    const manifest = readMacManifest();

    expect(manifest).toContain(
      '<Runtime resid="Taskpane.Url" lifetime="long" />',
    );
    expect(manifest).toContain('<FunctionFile resid="Taskpane.Url" />');
    expect(manifest).toContain(
      '<SourceLocation resid="Taskpane.Url" />',
    );
    expect(manifest).toContain(
      '<ExtendedOverrides Url="https://127.0.0.1:43119/shortcuts.json" />',
    );
  });

  it("exposes both floating Quick Trace commands on the Excel ribbon", () => {
    const manifest = readMacManifest();

    expect(manifest).toContain("<FunctionName>OpenQuickTracePrecedents</FunctionName>");
    expect(manifest).toContain("<FunctionName>OpenQuickTraceDependents</FunctionName>");
    expect(manifest).toContain("floating precedent audit");
    expect(manifest).toContain("floating dependent audit");
    const unified = readFileSync(resolve("manifest.json"), "utf8");
    expect(unified).toContain("OpenQuickTracePrecedents");
    expect(unified).toContain("OpenQuickTraceDependents");
  });

  it("keeps every hosted resource on the loopback HTTPS origin", () => {
    const manifest = readMacManifest();
    const urls = [
      ...manifest.matchAll(
        /(?:DefaultValue|Url)="(https:\/\/[^"]+)"|<AppDomain>(https:\/\/[^<]+)</gu,
      ),
    ].map((match) => match[1] ?? match[2] ?? "");

    expect(urls.length).toBeGreaterThan(0);
    expect(
      urls.every((url) => {
        const parsed = new URL(url);
        return (
          parsed.protocol === "https:" &&
          parsed.hostname === "127.0.0.1" &&
          parsed.port === "43119"
        );
      }),
    ).toBe(true);
  });

  it("serves every registered Mac shortcut action from the local bundle", () => {
    const shortcuts = JSON.parse(
      readFileSync(resolve("public/shortcuts.json"), "utf8"),
    ) as ShortcutsManifest;
    const commands = allCommands.filter((command) => command.actionId !== undefined);

    expect(shortcuts.actions).toEqual(commands.map((command) => ({
      id: command.actionId,
      name: command.label,
      type: "ExecuteFunction",
    })));
    expect(shortcuts.shortcuts).toEqual(commands.map((command) => ({
      action: command.actionId,
      key: {
        default: command.macShortcut
          ?.replace("Command", "Ctrl")
          .replace("Option", "Alt"),
        mac: command.macShortcut,
      },
    })));
    expect(shortcuts.shortcuts[0]).toEqual({
      action: "OpenAnalystLens",
      key: {
        default: "Ctrl+Alt+L",
        mac: "Command+Option+L",
      },
    });
    expect(shortcuts.actions).toHaveLength(21);
    expect(shortcuts.shortcuts).toContainEqual({
      action: "OpenQuickTracePrecedents",
      key: { default: "Ctrl+Alt+P", mac: "Ctrl+Option+P" },
    });
    expect(shortcuts.shortcuts).toContainEqual({
      action: "OpenQuickTraceDependents",
      key: { default: "Ctrl+Alt+D", mac: "Ctrl+Option+D" },
    });
    expect(shortcuts.actions).toContainEqual(expect.objectContaining({
      id: "ClearFormats",
    }));
    expect(shortcuts.shortcuts).toContainEqual({
      action: "ClearBorders",
      key: {
        default: "Ctrl+Alt+Shift+N",
        mac: "Ctrl+Option+Shift+N",
      },
    });
    for (const shortcut of shortcuts.shortcuts) {
      expect(shortcut.key.default.split("+").at(-1)).toMatch(/^(?:[A-Z0-9]|Up|Down)$/);
      expect(shortcut.key.mac.split("+").at(-1)).toMatch(/^(?:[A-Z0-9]|Up|Down)$/);
    }
  });
});
