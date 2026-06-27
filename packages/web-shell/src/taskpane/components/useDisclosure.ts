import { useCallback, useEffect, useRef, useState } from 'react';

/**
 * Click-toggle disclosure for the pane's interactive flyouts (the surface action menu, the context
 * tray, the skills list, the quick-action drawer). It replaces the old pure-CSS `:hover` popovers,
 * whose trigger→popover gap dropped the hover mid-reach and whose `pointer-events: none` made
 * interactive content unclickable. Opens on an explicit click; closes on `Escape` or a pointer-down
 * outside the container. Read-only tooltips stay on hover/focus — this is only for menus you act in.
 *
 * Wire it up: attach `containerRef` to the disclosure root, drive its `data-open` attribute from
 * `open`, point the trigger button's `onClick` at `toggle` with `aria-expanded={open}`. The CSS
 * reveals the popover on `[data-open='true']`.
 */
export function useDisclosure<T extends HTMLElement = HTMLDivElement>(): {
  open: boolean;
  toggle: () => void;
  close: () => void;
  containerRef: React.RefObject<T>;
} {
  const [open, setOpen] = useState(false);
  const containerRef = useRef<T>(null);
  const toggle = useCallback(() => setOpen((value) => !value), []);
  const close = useCallback(() => setOpen(false), []);

  useEffect(() => {
    if (!open) return undefined;
    const onPointerDown = (event: PointerEvent): void => {
      const root = containerRef.current;
      if (root && !root.contains(event.target as Node)) setOpen(false);
    };
    const onKeyDown = (event: KeyboardEvent): void => {
      if (event.key === 'Escape') setOpen(false);
    };
    // Capture phase so a click on another disclosure's trigger closes this one before that one opens.
    document.addEventListener('pointerdown', onPointerDown, true);
    document.addEventListener('keydown', onKeyDown, true);
    return () => {
      document.removeEventListener('pointerdown', onPointerDown, true);
      document.removeEventListener('keydown', onKeyDown, true);
    };
  }, [open]);

  return { open, toggle, close, containerRef };
}
