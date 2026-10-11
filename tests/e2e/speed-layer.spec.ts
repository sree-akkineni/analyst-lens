import { expect, test } from "@playwright/test";

test("applies and reverts a visible format command", async ({ page }) => {
  await page.goto("/taskpane.html?fixture=sample-model");

  await page.getByRole("button", { name: "Input blue" }).click();
  await expect(page.getByRole("status").filter({ hasText: "Input blue applied" }))
    .toBeVisible();
  await page.getByRole("button", { name: "Revert last Analyst Lens action" }).click();
  await expect(page.getByRole("status").filter({ hasText: "Last Analyst Lens action reverted" }))
    .toBeVisible();
});

test("reports a protected-sheet write failure without claiming success", async ({ page }) => {
  await page.goto("/taskpane.html?fixture=sample-model&selection=protected");
  await expect(page.getByRole("heading", { name: "Protected $B$2" })).toBeVisible();

  await page.getByRole("button", { name: "Formula black" }).click();

  await expect(page.getByRole("alert")).toContainText("protected fixture sheet");
  await expect(page.getByText("Formula black applied")).toHaveCount(0);
  await page.getByRole("button", { name: "Group rows", exact: true }).click();
  await expect(page.getByRole("alert")).toContainText("protected fixture sheet");
});

test("uses one-layer manual inverse copy and keeps visibility out of Revert history", async ({ page }) => {
  await page.goto("/taskpane.html?fixture=sample-model");

  await page.getByRole("button", { name: "Input blue" }).click();
  await page.getByRole("button", { name: "Open All Groups" }).click();
  await expect(page.getByRole("status").filter({
    hasText: "Group definitions were not changed",
  })).toBeVisible();
  await page.getByRole("button", { name: "Revert last Analyst Lens action" }).click();
  await expect(page.getByRole("status").filter({
    hasText: "Last Analyst Lens action reverted",
  })).toBeVisible();

  await page.getByRole("button", { name: "Group rows", exact: true }).click();
  await expect(page.getByRole("status").filter({ hasText: "Ungroup rows once" })).toBeVisible();
  await page.getByRole("button", { name: "Revert last Analyst Lens action" }).click();
  await expect(page.getByRole("alert")).toContainText("Ungroup rows once");
});
