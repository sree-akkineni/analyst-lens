import AxeBuilder from "@axe-core/playwright";
import { expect, test, type Locator, type Page } from "@playwright/test";

const expectNoHorizontalOverflow = async (page: Page) => {
  const overflow = await page.evaluate(() => ({
    body: document.body.scrollWidth - document.body.clientWidth,
    document: document.documentElement.scrollWidth - document.documentElement.clientWidth,
  }));
  expect(overflow).toEqual({ body: 0, document: 0 });
};

const expectAccessible = async (page: Page) => {
  const result = await new AxeBuilder({ page }).withTags(["wcag2a", "wcag2aa"]).analyze();
  const blocking = result.violations.filter((violation) =>
    violation.impact === "critical" || violation.impact === "serious");
  expect(blocking).toEqual([]);
};

const token = (page: Page, name: string) => page.evaluate((property) =>
  getComputedStyle(document.documentElement).getPropertyValue(property).trim(), name);

const expectedTokens = {
  light: {
    "--al-bg": "#f5f5f7",
    "--al-surface": "#fff",
    "--al-surface-2": "#f2f2f4",
    "--al-surface-raised": "#fff",
    "--al-surface-sunken": "#ececef",
    "--al-ink": "#1d1d1f",
    "--al-ink-soft": "#424245",
    "--al-ink-mute": "#5d5d62",
    "--al-ink-faint": "#6e6e73",
    "--al-border": "#d2d2d7",
    "--al-border-strong": "#a8a8ad",
    "--al-border-subtle": "#e5e5ea",
    "--al-live": "#006b5f",
    "--al-live-dim": "#007a6e",
    "--al-live-soft": "#e2f3f0",
    "--al-annotation": "#8a4b08",
    "--al-annotation-soft": "#f7eadb",
    "--al-input-blue": "#0057b7",
    "--al-warning": "#805000",
    "--al-warning-soft": "#fff3d6",
    "--al-danger": "#b42318",
    "--al-danger-soft": "#fdecea",
    "--al-ok": "#006b5f",
    "--al-focus": "#0067c0",
    "--al-on-accent": "#fff",
    "--al-shadow-pop": "0 12px 40px #0000002e",
    "--al-shadow-sticky": "0 6px 18px #00000024",
    "--al-glow-live": "0 0 6px #006b5f52",
  },
  dark: {
    "--al-bg": "#0e0f11",
    "--al-surface": "#15171a",
    "--al-surface-2": "#1b1e22",
    "--al-surface-raised": "#21252a",
    "--al-surface-sunken": "#0b0c0e",
    "--al-ink": "#e9ebee",
    "--al-ink-soft": "#c7cbd2",
    "--al-ink-mute": "#a7adb7",
    "--al-ink-faint": "#969ca6",
    "--al-border": "#34383e",
    "--al-border-strong": "#4b5058",
    "--al-border-subtle": "#26292e",
    "--al-live": "#5eead4",
    "--al-live-dim": "#2dd4bf",
    "--al-live-soft": "#1a2e2a",
    "--al-annotation": "#e6b980",
    "--al-annotation-soft": "#2a2419",
    "--al-input-blue": "#8ab4ff",
    "--al-warning": "#f0c77b",
    "--al-warning-soft": "#2a2419",
    "--al-danger": "#ff9b9b",
    "--al-danger-soft": "#2a1a1a",
    "--al-ok": "#5eead4",
    "--al-focus": "#6ee7d6",
    "--al-on-accent": "#071513",
    "--al-shadow-pop": "0 12px 40px #0000008c",
    "--al-shadow-sticky": "0 6px 18px #0006",
    "--al-glow-live": "0 0 6px #5eead499",
  },
} as const;

type ColorScheme = keyof typeof expectedTokens;
type TokenName = keyof typeof expectedTokens.light;

const tokenNames = Object.keys(expectedTokens.light) as TokenName[];
const surfaceTokenNames = [
  "--al-bg",
  "--al-surface",
  "--al-surface-2",
  "--al-surface-raised",
  "--al-surface-sunken",
] as const satisfies readonly TokenName[];

const tokens = async (page: Page, names: readonly TokenName[] = tokenNames) =>
  Object.fromEntries(await Promise.all(names.map(async (name) => [name, await token(page, name)])));

const resolvedColor = (page: Page, color: string) => page.evaluate((value) => {
  const probe = document.createElement("span");
  probe.style.color = value;
  document.body.append(probe);
  const resolved = getComputedStyle(probe).color;
  probe.remove();
  return resolved;
}, color);

const focusWithKeyboard = async (page: Page, control: Locator) => {
  await page.evaluate(() => {
    if (document.activeElement instanceof HTMLElement) document.activeElement.blur();
  });
  for (let index = 0; index < 50; index += 1) {
    await page.keyboard.press("Tab");
    if (await control.evaluate((element) => element === document.activeElement)) break;
  }
  await expect(control).toBeFocused();
};

const expectVisibleFocus = async (page: Page, control: Locator, focusToken: string) => {
  const outline = await control.evaluate((element) => {
    const style = getComputedStyle(element);
    return {
      color: style.outlineColor,
      style: style.outlineStyle,
      width: style.outlineWidth,
    };
  });
  expect(outline.style).not.toBe("none");
  expect(Number.parseFloat(outline.width)).toBeGreaterThan(0);
  expect(outline.color).toBe(await resolvedColor(page, focusToken));
};

const taskPaneSizes = [{ width: 360, height: 900 }, { width: 760, height: 900 }];
const quickTraceSizes = [{ width: 420, height: 390 }, { width: 900, height: 700 }];
const taskPaneStates = [
  { heading: "DCF $F$42", name: "formula", selection: "formula" },
  { heading: "Inputs $B$7", name: "hardcode", selection: "hardcode" },
  { heading: "DCF $A$1", name: "blank", selection: "blank" },
  { heading: "Checks $H$5", name: "error", selection: "error" },
  { heading: "2 × 2 selection", name: "multi", selection: "multi" },
  { heading: "DCF $F$42", name: "command palette open", selection: "formula" },
  { heading: "DCF $F$42", name: "dependency map open", selection: "formula" },
  { heading: "DCF $F$42", name: "Model health open and completed", selection: "formula" },
  { heading: "DCF $F$42", name: "Key outputs with a pin", selection: "formula" },
] as const;

type TaskPaneState = (typeof taskPaneStates)[number];

const expectKeyboardReachable = async (control: Locator) => {
  await expect(control).toBeVisible();
  await expect(control).toBeEnabled();
  expect(await control.evaluate((element) => (element as HTMLElement).tabIndex))
    .toBeGreaterThanOrEqual(0);
};

const prepareTaskPaneState = async (
  page: Page,
  state: TaskPaneState,
): Promise<readonly Locator[]> => {
  const launcher = page.getByRole("button", {
    name: /Open commands for this cell.*Command-K/u,
  });
  const keyOutputs = page.getByRole("button", { name: /Key outputs/u });
  const speedTools = page.getByRole("button", { name: /Speed tools/u });
  const modelHealth = page.getByRole("button", { name: /Model health/u });

  if (state.name === "command palette open") {
    await launcher.click();
    const dialog = page.getByRole("dialog", { name: "Command palette" });
    await expect(dialog).toBeVisible();
    expect(await dialog.evaluate((element) => element.contains(document.activeElement))).toBe(true);
    return [
      dialog.getByRole("searchbox", { name: "Find a command" }),
      dialog.getByRole("button", { name: "Close command palette" }),
    ];
  }
  if (state.name === "dependency map open") {
    await page.getByRole("button", { name: "Open dependency map" }).click();
    const dialog = page.getByRole("dialog", { name: "Dependency map" });
    await expect(dialog).toBeVisible();
    expect(await dialog.evaluate((element) => element.contains(document.activeElement))).toBe(true);
    return [
      dialog.getByRole("searchbox", { name: "Search dependency map" }),
      dialog.getByRole("button", { name: "Return to origin from dependency map" }),
      dialog.getByRole("button", { name: "Close dependency map" }),
    ];
  }
  if (state.name === "Model health open and completed") {
    await modelHealth.click();
    await page.getByRole("button", { name: "Run model health scan" }).click();
    await expect(page.getByText(/Scan complete\./u)).toBeVisible();
    await expect(page.getByRole("list", { name: "Health findings" })).toBeVisible();
    return [
      page.getByRole("button", { name: "Export local Markdown" }),
      page.getByRole("button", { name: /ERR_VALUE$/u }),
    ];
  }
  if (state.name === "Key outputs with a pin") {
    await page.getByRole("button", { name: "Pin key output" }).click();
    await expect(keyOutputs).toHaveAttribute("aria-expanded", "true");
    await expect(page.getByRole("heading", { name: "DCF $F$42", exact: true })).toHaveCount(2);
    return [
      page.getByRole("button", { name: "Refresh key outputs" }),
      page.getByRole("button", { name: "Unpin DCF $F$42" }),
    ];
  }
  if (state.name === "multi") {
    await expect(launcher).toHaveCount(0);
    return [keyOutputs, speedTools, modelHealth];
  }
  return [launcher, keyOutputs, speedTools, modelHealth];
};

for (const colorScheme of Object.keys(expectedTokens) as ColorScheme[]) {
  test(`task pane resolves exact ${colorScheme} semantic tokens and focus`, async ({ page }) => {
    await page.emulateMedia({ colorScheme });
    await page.setViewportSize({ width: 360, height: 900 });
    await page.goto("/taskpane.html?fixture=sample-model");
    await expect(page.getByRole("heading", { name: "DCF $F$42" })).toBeVisible();
    expect(await tokens(page)).toEqual(expectedTokens[colorScheme]);

    const pinButton = page.getByRole("button", { name: "Pin key output" });
    await focusWithKeyboard(page, pinButton);
    await expectVisibleFocus(page, pinButton, expectedTokens[colorScheme]["--al-focus"]);
  });

  test(`Quick Trace resolves exact ${colorScheme} surfaces, accents, and focus`, async ({ page }) => {
    await page.emulateMedia({ colorScheme });
    await page.goto("/quick-trace.html?fixture=sample-model&direction=precedent");
    const expected = expectedTokens[colorScheme];
    expect(await tokens(page, surfaceTokenNames)).toEqual(
      Object.fromEntries(surfaceTokenNames.map((name) => [name, expected[name]])),
    );

    const live = await resolvedColor(page, expected["--al-live"]);
    const onAccent = await resolvedColor(page, expected["--al-on-accent"]);
    const pressedButton = page.getByRole("button", { name: "Trace In" });
    const selectedRow = page.getByRole("treeitem", { selected: true });
    await expect(pressedButton).toHaveAttribute("aria-pressed", "true");
    await expect(pressedButton).toHaveCSS("background-color", live);
    await expect(pressedButton).toHaveCSS("color", onAccent);
    await expect(selectedRow).toHaveCSS("background-color", live);
    await expect(selectedRow).toHaveCSS("color", onAccent);

    await focusWithKeyboard(page, pressedButton);
    await expectVisibleFocus(page, pressedButton, expected["--al-focus"]);
  });

  for (const viewport of taskPaneSizes) {
    test(`task pane state matrix remains accessible in ${colorScheme} at ${viewport.width} by ${viewport.height}`, async ({
      page,
    }) => {
      await page.emulateMedia({ colorScheme });
      await page.setViewportSize(viewport);
      for (const state of taskPaneStates) {
        await test.step(state.name, async () => {
          await page.goto(
            `/taskpane.html?fixture=sample-model&selection=${state.selection}`,
          );

          await expect(page.getByRole("heading", { name: state.heading })).toBeVisible();
          await expect(page.getByRole("button", { name: /Key outputs/u })).toBeVisible();
          await expect(page.getByRole("button", { name: /Speed tools/u })).toBeVisible();
          await expect(page.getByRole("button", { name: /Model health/u })).toBeVisible();
          const stateControls = await prepareTaskPaneState(page, state);
          for (const control of stateControls) await expectKeyboardReachable(control);

          if (state.name === "formula" && viewport.width === 360) {
            const annotatedDependencies = page
              .getByRole("list", { name: "Direct precedents" })
              .locator(".annotation-button");
            for (const rowIndex of [3, 4]) {
              const row = annotatedDependencies.nth(rowIndex);
              const layout = await row.evaluate((element) => {
                const body = element.querySelector(".annotation-body");
                const reference = element.querySelector(".annotation-ref");
                const gloss = element.querySelector(".annotation-gloss");
                const value = element.querySelector(".annotation-val");
                if (
                  !(body instanceof HTMLElement) ||
                  !(reference instanceof HTMLElement) ||
                  !(gloss instanceof HTMLElement) ||
                  !(value instanceof HTMLElement)
                ) {
                  throw new Error("Expected annotated dependency layout");
                }
                const rectangle = (target: HTMLElement) => {
                  const bounds = target.getBoundingClientRect();
                  return { left: bounds.left, right: bounds.right, width: bounds.width };
                };
                return {
                  body: rectangle(body),
                  bodyMinWidth: getComputedStyle(body).minWidth,
                  gloss: rectangle(gloss),
                  reference: rectangle(reference),
                  value: rectangle(value),
                };
              });
              expect(layout.bodyMinWidth).toBe("0px");
              expect(layout.body.width).toBeGreaterThan(0);
              expect(layout.reference.left).toBeGreaterThanOrEqual(layout.body.left);
              expect(layout.reference.right).toBeLessThanOrEqual(layout.body.right);
              expect(layout.gloss.left).toBeGreaterThanOrEqual(layout.body.left);
              expect(layout.gloss.right).toBeLessThanOrEqual(layout.body.right);
              expect(layout.body.right).toBeLessThan(layout.value.left);
            }
          }
          await expectNoHorizontalOverflow(page);
          await expectAccessible(page);
        });
      }
    });
  }

  for (const viewport of quickTraceSizes) {
    test(`Quick Trace direction matrix remains accessible in ${colorScheme} at ${viewport.width} by ${viewport.height}`, async ({
      page,
    }) => {
      await page.emulateMedia({ colorScheme });
      await page.setViewportSize(viewport);
      for (const direction of [
        { fixture: "precedent", label: "Trace In" },
        { fixture: "dependent", label: "Trace Out" },
      ] as const) {
        await test.step(direction.label, async () => {
          await page.goto(`/quick-trace.html?fixture=sample-model&direction=${direction.fixture}`);

          const tree = page.getByRole("tree", {
            name: `${direction.label} dependencies`,
          });
          const currentRoot = page.getByRole("region", { name: "Current root DCF $F$42" });
          const traceIn = page.getByRole("button", { name: "Trace In" });
          const traceOut = page.getByRole("button", { name: "Trace Out" });
          const returnToOrigin = page.getByRole("button", { name: "Return to origin" });
          const close = page.getByRole("button", { name: "Close" });

          await expect(tree).toBeVisible();
          await expect(currentRoot).toBeVisible();
          await expect(traceIn).toBeVisible();
          await expect(traceOut).toBeVisible();
          await expect(returnToOrigin).toBeVisible();
          await expect(close).toBeVisible();
          for (const control of [tree, traceIn, traceOut, returnToOrigin, close]) {
            await focusWithKeyboard(page, control);
          }
          await expectNoHorizontalOverflow(page);
          await expectAccessible(page);
        });
      }
    });
  }
}

test("task pane honors reduced motion", async ({ page }) => {
  await page.emulateMedia({ reducedMotion: "reduce" });
  await page.goto("/taskpane.html?fixture=sample-model");
  const durations = await page.evaluate(() => {
    const style = getComputedStyle(document.querySelector(".panel-chevron") as Element);
    return { animation: style.animationDuration, transition: style.transitionDuration };
  });
  const toMilliseconds = (value: string) => value.endsWith("ms")
    ? Number.parseFloat(value)
    : Number.parseFloat(value) * 1000;
  expect(toMilliseconds(durations.animation)).toBeLessThanOrEqual(0.01);
  expect(toMilliseconds(durations.transition)).toBeLessThanOrEqual(0.01);
});
