import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { useState } from "react";
import { afterEach, describe, expect, it } from "vitest";
import { Disclosure } from "../../src/ui/Disclosure";

afterEach(cleanup);

const StatefulChild = () => {
  const [value, setValue] = useState(0);
  return <button onClick={() => setValue((current) => current + 1)} type="button">Count {value}</button>;
};

describe("Disclosure", () => {
  it("uses native hidden while retaining child state", () => {
    render(<Disclosure eyebrow="Evidence" title="Model health" titleId="health"><StatefulChild /></Disclosure>);
    const toggle = screen.getByRole("button", { name: /Model health/u });
    const body = document.getElementById("health-body");
    expect(toggle.getAttribute("aria-expanded")).toBe("false");
    expect(body?.hidden).toBe(true);
    expect(screen.queryByRole("button", { name: "Count 0" })).toBeNull();

    fireEvent.click(toggle);
    fireEvent.click(screen.getByRole("button", { name: "Count 0" }));
    fireEvent.click(toggle);
    expect(body?.hidden).toBe(true);
    fireEvent.click(toggle);
    expect(screen.getByRole("button", { name: "Count 1" })).toBeTruthy();
  });

  it("auto-opens once when openWhen changes from false to true", () => {
    const { rerender } = render(
      <Disclosure eyebrow="Session" openWhen={false} title="Key outputs" titleId="outputs"><StatefulChild /></Disclosure>,
    );
    rerender(<Disclosure eyebrow="Session" openWhen title="Key outputs" titleId="outputs"><StatefulChild /></Disclosure>);
    const toggle = screen.getByRole("button", { name: /Key outputs/u });
    expect(toggle.getAttribute("aria-expanded")).toBe("true");
    fireEvent.click(toggle);
    rerender(<Disclosure eyebrow="Session" openWhen title="Key outputs" titleId="outputs"><StatefulChild /></Disclosure>);
    expect(toggle.getAttribute("aria-expanded")).toBe("false");
  });
});
