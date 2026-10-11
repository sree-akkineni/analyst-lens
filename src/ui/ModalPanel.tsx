import { useEffect, useRef, type ReactNode } from "react";

type ModalPanelProps = {
  children: ReactNode;
  labelledBy: string;
  onClose(): void;
};

const focusableSelector = [
  "button:not(:disabled)",
  "input:not(:disabled)",
  "select:not(:disabled)",
  "textarea:not(:disabled)",
  "a[href]",
  "[tabindex]:not([tabindex='-1'])",
].join(",");

export function ModalPanel({ children, labelledBy, onClose }: ModalPanelProps) {
  const panel = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const current = panel.current;
    if (current === null || current.contains(document.activeElement)) return;
    current.querySelector<HTMLElement>(focusableSelector)?.focus();
  }, []);

  return (
    <div
      aria-labelledby={labelledBy}
      aria-modal="true"
      className="overlay-panel"
      onKeyDown={(event) => {
        if (event.key === "Escape") {
          event.preventDefault();
          onClose();
          return;
        }
        if (event.key !== "Tab") return;
        const current = panel.current;
        if (current === null) return;
        const focusable = [...current.querySelectorAll<HTMLElement>(focusableSelector)];
        const first = focusable[0];
        const last = focusable.at(-1);
        if (first === undefined || last === undefined) {
          event.preventDefault();
          return;
        }
        if (event.shiftKey && document.activeElement === first) {
          event.preventDefault();
          last.focus();
        } else if (!event.shiftKey && document.activeElement === last) {
          event.preventDefault();
          first.focus();
        }
      }}
      ref={panel}
      role="dialog"
    >
      {children}
    </div>
  );
}
