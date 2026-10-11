import AxeBuilder from "@axe-core/playwright";
import { expect, test, type Page } from "@playwright/test";

const expectAccessible = async (page: Page) => {
  const result = await new AxeBuilder({ page }).withTags(["wcag2a", "wcag2aa"]).analyze();
  const blocking = result.violations.filter((violation) =>
    violation.impact === "critical" || violation.impact === "serious");
  expect(blocking).toEqual([]);
};

test("trace a branch and return to origin", async ({ page }) => {
  await page.goto("/taskpane.html?fixture=sample-model");
  await expect(page.getByRole("heading", { name: "DCF $F$42" })).toBeVisible();
  const precedents = page.getByRole("list", { name: "Direct precedents" });
  await expect(precedents).toBeVisible();
  await expect(precedents.getByRole("button", { name: "Inputs $B$7" })).toBeVisible();
  await expect(page.getByRole("button", {
    name: /Open commands for this cell.*Command-K/u,
  })).toBeVisible();
  await expect(page.getByText("Ask about this cell…")).toHaveCount(0);
  await page.getByRole("button", {
    name: /Open commands for this cell.*Command-K/u,
  }).click();
  await expect(page.getByRole("dialog", { name: "Command palette" })).toBeVisible();
  await expectAccessible(page);
  await page.getByRole("button", { name: "Close command palette" }).click();
  await precedents.getByRole("button", { name: "Inputs $B$7" }).click();
  await expect(page.getByRole("heading", { name: "Inputs $B$7" })).toBeVisible();
  const dependents = page.getByRole("list", { name: "Direct dependents" });
  await expect(dependents).toBeVisible();
  await dependents.getByRole("button", { name: "DCF $F$42" }).click();
  await expect(page.getByRole("heading", { name: "DCF $F$42" })).toBeVisible();
  await page.getByRole("button", { name: "Open dependency map" }).click();
  const map = page.getByRole("dialog", { name: "Dependency map" });
  await map.getByRole("button", { name: "Expand precedents for DCF $F$42" }).click();
  await map.getByRole("button", { name: "Navigate to Inputs $B$7" }).click();
  await map.getByRole("button", { name: "Return to origin from dependency map" }).click();
  await expect(page.getByRole("heading", { name: "DCF $F$42" })).toBeVisible();
});

for (const overlay of [
  {
    closeName: "Close dependency map",
    dialogName: "Dependency map",
    openerName: "Open dependency map",
  },
  {
    closeName: "Close command palette",
    dialogName: "Command palette",
    openerName: /Open commands for this cell.*Command-K/u,
  },
  {
    closeName: "Close settings",
    dialogName: "Inspector settings",
    openerName: "Settings",
  },
] as const) {
  test(`Command-K stays within the active ${overlay.dialogName} modal`, async ({ page }) => {
    await page.goto("/taskpane.html?fixture=sample-model");
    const opener = page.getByRole("button", { name: overlay.openerName });
    await opener.click();
    const dialog = page.getByRole("dialog", { name: overlay.dialogName });
    const close = dialog.getByRole("button", { name: overlay.closeName });
    await expect(dialog).toBeVisible();
    await close.focus();

    await page.keyboard.press("Meta+k");

    await expect(page.getByRole("dialog")).toHaveCount(1);
    await expect(dialog).toBeVisible();
    expect(await dialog.evaluate((element) => element.contains(document.activeElement))).toBe(true);

    await close.click();
    await expect(page.getByRole("dialog")).toHaveCount(0);
    await expect(opener).toBeFocused();
  });
}

for (const fixture of [
  { heading: "DCF $F$42", selection: "formula" },
  { heading: "Inputs $B$7", selection: "hardcode" },
  { heading: "DCF $A$1", selection: "blank" },
  { heading: "Checks $H$5", selection: "error" },
  { heading: "2 × 2 selection", selection: "multi" },
] as const) {
  test(`renders the ${fixture.selection} selection state`, async ({ page }) => {
    await page.goto(`/taskpane.html?fixture=sample-model&selection=${fixture.selection}`);
    await expect(page.getByRole("heading", { name: fixture.heading })).toBeVisible();
    await expectAccessible(page);
  });
}

test("shows cycle, external, and dynamic graph evidence", async ({ page }) => {
  await page.goto("/taskpane.html?fixture=sample-model");
  await page.getByRole("button", { name: "Open dependency map" }).click();
  const map = page.getByRole("dialog", { name: "Dependency map" });

  await expect(map.getByText("External leaf")).toBeVisible();
  await expect(map.getByText("Dynamic unresolved leaf")).toBeVisible();
  await map.getByRole("button", { name: "Expand precedents for DCF $C$10" }).click();
  await map.getByRole("button", { name: "Expand precedents for DCF $C$11" }).click();
  await expect(map.getByRole("list", { name: "Dependency edges" })).toContainText("Cycle");
  await expectAccessible(page);
});
