import { useEffect, useRef, useState, type ReactNode } from "react";

type DisclosureProps = {
  children: ReactNode;
  className?: string | undefined;
  defaultOpen?: boolean | undefined;
  eyebrow: string;
  meta?: string | undefined;
  openWhen?: boolean | undefined;
  title: string;
  titleId: string;
};

/**
 * Progressive-disclosure section. The header is a real <button> so keyboard
 * users get a focusable, Space/Enter-operable toggle; the region is wired to
 * the button via aria-controls/aria-expanded. Content stays mounted while
 * closed (hidden) to preserve any internal state.
 */
export function Disclosure({
  children,
  className,
  defaultOpen = false,
  eyebrow,
  meta,
  openWhen = false,
  title,
  titleId,
}: DisclosureProps) {
  const [open, setOpen] = useState(defaultOpen);
  const previousOpenWhen = useRef(openWhen);
  const bodyId = `${titleId}-body`;

  useEffect(() => {
    if (openWhen && !previousOpenWhen.current) setOpen(true);
    previousOpenWhen.current = openWhen;
  }, [openWhen]);

  return (
    <section aria-labelledby={titleId} className={`panel${className === undefined ? "" : ` ${className}`}`} data-open={open}>
      <h2 className="panel-heading">
        <button
          aria-controls={bodyId}
          aria-expanded={open}
          className="panel-toggle"
          onClick={() => setOpen((current) => !current)}
          type="button"
        >
          <span className="panel-title">
            <span className="section-label">{eyebrow}</span>
            <span className="panel-title-text" id={titleId}>{title}</span>
          </span>
          {meta === undefined ? null : <span className="panel-meta">{meta}</span>}
          <span aria-hidden="true" className="panel-chevron">›</span>
        </button>
      </h2>
      <div aria-labelledby={titleId} className="panel-body" hidden={!open} id={bodyId} role="region">{children}</div>
    </section>
  );
}
