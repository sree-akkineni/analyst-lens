import AxeBuilder from "@axe-core/playwright";
import { expect, test, type Page } from "@playwright/test";

const expectAccessible = async (page: Page) => {
  const result = await new AxeBuilder({ page }).withTags(["wcag2a", "wcag2aa"]).analyze();
  const blocking = result.violations.filter((violation) =>
    violation.impact === "critical" || violation.impact === "serious");
  expect(blocking).toEqual([]);
};

test("cancels a health scan without partial findings, then completes", async ({ page }) => {
  await page.goto("/taskpane.html?fixture=sample-model");

  await page.getByRole("button", { name: /Model health/u }).click();
  await expectAccessible(page);
  await page.getByRole("button", { name: "Run model health scan" }).click();
  await page.getByRole("button", { name: "Cancel health scan" }).click();
  await expect(page.getByText("Scan cancelled. No partial findings were kept."))
    .toBeVisible();
  await expect(page.getByRole("list", { name: "Health findings" })).toHaveCount(0);

  await page.getByRole("button", { name: "Run model health scan" }).click();
  await expect(page.getByText(/Scan complete\./u)).toBeVisible();
  await expect(page.getByRole("list", { name: "Health findings" })).toBeVisible();
  await expect(page.getByRole("button", { name: /ERR_VALUE$/u })).toBeVisible();
  await expectAccessible(page);
});

test("pins an output, shows assumption impact, and saves metadata only after opt-in", async ({
  page,
}) => {
  await page.goto("/taskpane.html?fixture=sample-model");
  await page.getByRole("button", { name: "Pin key output" }).click();
  await expect(page.getByRole("button", { name: /Key outputs/u })).toHaveAttribute("aria-expanded", "true");
  await expect(page.getByRole("heading", { name: "DCF $F$42", exact: true })).toHaveCount(2);
  await expectAccessible(page);

  await page.getByRole("button", { name: "Inputs $B$7", exact: true }).click();
  await expect(page.getByRole("heading", { name: "Inputs $B$7" })).toBeVisible();
  await expect(page.getByText("Affected by current selection", { exact: true })).toBeVisible();
  await expect(page.getByText("1 affected by current selection")).toBeVisible();

  const save = page.getByRole("button", { name: "Save watchlist with workbook" });
  await expect(save).toBeDisabled();
  await page.getByRole("checkbox", {
    name: "Store only addresses, labels, and categories in this workbook",
  }).check();
  await save.click();
  await expect(page.getByText("Watchlist metadata saved with this workbook.")).toBeVisible();
});
