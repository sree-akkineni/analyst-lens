import AxeBuilder from "@axe-core/playwright";
import { expect, test, type Page } from "@playwright/test";

const expectAccessible = async (page: Page) => {
  const result = await new AxeBuilder({ page }).withTags(["wcag2a", "wcag2aa"]).analyze();
  const blocking = result.violations.filter((violation) =>
    violation.impact === "critical" || violation.impact === "serious");
  expect(blocking).toEqual([]);
};

test("traces, expands, switches direction, reroots, and returns to origin", async ({ page }) => {
  await page.goto("/quick-trace.html?fixture=sample-model&direction=precedent");
  const tree = page.getByRole("tree", { name: "Trace In dependencies" });
  await expect(tree).toBeVisible();
  await expect(page.getByText("Reference", { exact: true })).toBeVisible();
  await expect(page.getByText("Value", { exact: true })).toBeVisible();
  await expect(page.getByText("Argument", { exact: true })).toHaveCount(0);
  await expectAccessible(page);

  await tree.press("ArrowDown");
  await tree.press("ArrowDown");
  await expect(page.getByRole("treeitem", { selected: true })).toContainText("$D$20");
  await tree.press("ArrowRight");
  await expect(page.getByRole("treeitem")).toHaveCount(7);
  await tree.press("ArrowLeft");
  await expect(page.getByRole("treeitem")).toHaveCount(6);
  await tree.press("ArrowRight");
  await expect(page.getByRole("treeitem")).toHaveCount(7);
  await tree.press("Enter");
  await expect(page.getByRole("region", { name: "Current root DCF $D$20" })).toBeVisible();
  await tree.press("t");
  const traceOut = page.getByRole("tree", { name: "Trace Out dependencies" });
  await expect(traceOut).toBeVisible();
  await expectAccessible(page);
  await traceOut.press("o");
  await expect(page.getByRole("region", { name: "Current root DCF $F$42" })).toBeVisible();
});

test("shows cycles and unresolved leaves without recursive expansion", async ({ page }) => {
  await page.goto("/quick-trace.html?fixture=sample-model&direction=precedent");
  await expect(page.getByText("External", { exact: true })).toBeVisible();
  await expect(page.getByText("Dynamic", { exact: true })).toBeVisible();
  await page.getByRole("button", { name: /Expand DCF.*\$C\$10/u }).click();
  await page.getByRole("button", { name: /Expand DCF.*\$C\$11/u }).click();
  await expect(page.getByText("Cycle", { exact: true })).toBeVisible();
});

for (const viewport of [{ width: 420, height: 390 }, { width: 900, height: 700 }]) {
  test(`remains operable at ${viewport.width} by ${viewport.height}`, async ({ page }) => {
    await page.setViewportSize(viewport);
    await page.goto("/quick-trace.html?fixture=sample-model&direction=precedent");
    await expect(page.getByRole("button", { name: "Return to origin" })).toBeVisible();
    await expect(page.getByRole("button", { name: "Close" })).toBeVisible();
  });
}

test("keeps the selected-reference action fully reachable above the footer at 420 by 390", async ({
  page,
}) => {
  await page.setViewportSize({ width: 420, height: 390 });
  await page.goto("/quick-trace.html?fixture=sample-model&direction=dependent");
  const action = page.getByRole("button", { name: "Make selected reference the root" });
  const secondary = page.locator(".quick-trace-secondary");
  const footer = page.locator(".quick-trace-footer");
  await expect(action).toBeVisible();
  const bounds = await Promise.all([
    action.boundingBox(),
    secondary.boundingBox(),
    footer.boundingBox(),
  ]);
  const [actionBox, secondaryBox, footerBox] = bounds;
  expect(actionBox).not.toBeNull();
  expect(secondaryBox).not.toBeNull();
  expect(footerBox).not.toBeNull();
  if (actionBox === null || secondaryBox === null || footerBox === null) return;
  expect(actionBox.y).toBeGreaterThanOrEqual(secondaryBox.y);
  expect(actionBox.y + actionBox.height).toBeLessThanOrEqual(
    secondaryBox.y + secondaryBox.height,
  );
  expect(actionBox.y + actionBox.height).toBeLessThanOrEqual(footerBox.y);
});

test("closes and reopens without stale root or expansion state", async ({ browser }) => {
  const context = await browser.newContext();
  let page = await context.newPage();
  await page.goto("/quick-trace.html?fixture=sample-model&direction=precedent");
  const tree = page.getByRole("tree", { name: "Trace In dependencies" });
  await tree.press("ArrowDown");
  await tree.press("ArrowDown");
  await tree.press("Enter");
  await expect(page.getByRole("region", { name: "Current root DCF $D$20" })).toBeVisible();
  await page.getByRole("button", { name: "Close" }).click();
  await page.close();

  page = await context.newPage();
  await page.goto("/quick-trace.html?fixture=sample-model&direction=precedent");
  await expect(page.getByRole("region", { name: "Current root DCF $F$42" })).toBeVisible();
  await expect(page.getByRole("region", { name: "Current root DCF $D$20" })).toHaveCount(0);
  await context.close();
});
