---
name: frontend-design
description: >-
  Design and build polished, production-quality frontend UI — landing pages,
  marketing sites, web apps, components, and design systems. Use this whenever
  the task is to create, restyle, or review a user-facing interface (HTML/CSS,
  React/Vue/Svelte, Tailwind, design tokens), when a change needs to look
  intentional rather than default, or when the user mentions "make it look
  good", "modern UI", "landing page", "redesign", "brand", "responsive", or
  "accessible". Read this BEFORE writing the first line of markup or choosing
  colors, spacing, or type.
---

# Frontend Design

A method for producing frontend that looks intentional and considered — not
template-default. The goal is work that a designer would sign off on: coherent
type, deliberate spacing, restrained color, and interactions that feel solid.

## When to use

- Building a landing page, marketing site, or hero section
- Creating or restyling components (buttons, cards, forms, nav, tables)
- Establishing a design system or design tokens for a project
- Reviewing an existing UI for polish and accessibility
- Any request to "make it look good / modern / premium"

If the project already ships a design system or brand guide, **follow it** —
this skill fills the gaps, it does not override an owner's tokens.

## Core principles

1. **Commit to a direction.** Pick one point of view (editorial, brutalist,
   soft/rounded, technical, luxe) and make every decision serve it. A page
   that hedges reads as default. Mixed metaphors read as broken.
2. **Type carries the design.** Most of the visual quality of a page is the
   typography. Get the type scale, line-height, measure, and weight contrast
   right before touching color or decoration.
3. **Space is a feature, not leftover.** Use a consistent spacing scale and
   let elements breathe. Generous, rhythmic whitespace is the fastest path
   from "fine" to "designed".
4. **Restraint with color.** One dominant neutral, one accent, a couple of
   supporting tones. Color should guide attention, not decorate everything.
5. **Depth is earned.** Shadows, borders, and blur should map to real
   elevation and hierarchy — not scatter for effect.
6. **Motion confirms, never distracts.** Transitions exist to acknowledge
   state changes (hover, focus, open/close). Keep them fast (120–240ms) and
   respect `prefers-reduced-motion`.
7. **Accessible by construction.** Contrast, focus states, semantics, and
   keyboard operability are part of "looks good", not a later pass.

## Design tokens — start here

Define these before writing components. Reuse them everywhere; never hardcode
one-off values.

```css
:root {
  /* Spacing — a 4px base, geometric-ish scale */
  --space-1: 0.25rem;  --space-2: 0.5rem;  --space-3: 0.75rem;
  --space-4: 1rem;     --space-6: 1.5rem;  --space-8: 2rem;
  --space-12: 3rem;    --space-16: 4rem;   --space-24: 6rem;

  /* Type scale — ~1.2–1.25 ratio */
  --text-xs: 0.75rem;  --text-sm: 0.875rem; --text-base: 1rem;
  --text-lg: 1.125rem; --text-xl: 1.5rem;   --text-2xl: 2rem;
  --text-3xl: 2.5rem;  --text-4xl: 3.5rem;

  /* Radii */
  --radius-sm: 0.375rem; --radius-md: 0.625rem; --radius-lg: 1rem;
  --radius-full: 999px;

  /* Neutrals (light) — one ramp, used for text, borders, surfaces */
  --bg: #ffffff; --surface: #f7f7f5; --border: #e6e6e1;
  --text: #1a1a17; --text-muted: #6b6b63;

  /* Accent — pick ONE brand color; derive hover/active by shifting L */
  --accent: #b45309; --accent-hover: #92400e; --accent-contrast: #ffffff;

  /* Elevation */
  --shadow-sm: 0 1px 2px rgb(0 0 0 / 0.06);
  --shadow-md: 0 4px 16px rgb(0 0 0 / 0.08);
  --shadow-lg: 0 12px 40px rgb(0 0 0 / 0.12);
}

@media (prefers-color-scheme: dark) {
  :root {
    --bg: #14130f; --surface: #1e1c17; --border: #2e2b23;
    --text: #f2efe6; --text-muted: #a19b8c;
  }
}
```

> Swap the accent and neutral hues for the project's brand. For 도하커피 and
> similar warm/artisanal brands, warm neutrals (stone/sand) with a single
> roasted accent (amber/terracotta) work well. Keep the *structure* above.

## Typography rules

- **Two families max**: one for display/headings, one for body. A single
  well-set family is often better than two chosen carelessly.
- **Set a measure**: body copy at `max-width: 60–75ch`. Long lines kill
  readability more than any font choice.
- **Line-height**: ~1.5–1.65 for body, ~1.1–1.25 for large headings.
- **Weight contrast** creates hierarchy cheaply: pair a heavy heading (600–800)
  with a regular body (400). Avoid faux-bolding via all-caps everywhere.
- **Tighten large type**: apply slightly negative `letter-spacing`
  (-0.01em to -0.02em) on headings ≥ `--text-2xl`.

## Layout & spacing

- Use a consistent max content width (e.g. `--content: 72rem`) with generous
  horizontal padding that scales down on mobile.
- Build with CSS Grid / Flexbox; align to a spacing scale, not arbitrary px.
- Vertical rhythm: section padding should feel roomy
  (`--space-16` to `--space-24` between major sections).
- Establish clear hierarchy: one primary CTA per view, obvious focal point.

## Color usage

- 60/30/10: ~60% neutral surface, ~30% secondary/text, ~10% accent.
- Reserve the accent for actions and highlights, not backgrounds of everything.
- Verify text/background contrast meets WCAG AA (4.5:1 body, 3:1 large text).
- In dark mode, don't invert — re-derive a ramp; pure black + pure white
  vibrate. Use near-black surfaces and off-white text.

## Components — quality bar

Every interactive element needs all states designed:
`default · hover · focus-visible · active · disabled` (and `loading` where
relevant). A button without a visible focus ring is unfinished.

```css
.btn {
  display: inline-flex; align-items: center; gap: var(--space-2);
  padding: var(--space-3) var(--space-6);
  border-radius: var(--radius-md);
  font-weight: 600; line-height: 1;
  background: var(--accent); color: var(--accent-contrast);
  border: 1px solid transparent;
  transition: background 160ms ease, transform 160ms ease;
}
.btn:hover { background: var(--accent-hover); }
.btn:active { transform: translateY(1px); }
.btn:focus-visible {
  outline: 2px solid var(--accent);
  outline-offset: 2px;
}
.btn:disabled { opacity: 0.5; cursor: not-allowed; }
```

## Responsiveness

- Mobile-first: base styles for small screens, layer up with `min-width`
  media queries or container queries.
- Test the real breakpoints your content needs, not device names.
- Never let the page scroll horizontally. Wide content (tables, code, diagrams)
  scrolls inside its own `overflow-x: auto` container.
- Tap targets ≥ 44×44px.

## Accessibility checklist

- [ ] Semantic HTML (`<button>`, `<nav>`, `<main>`, headings in order)
- [ ] Visible `:focus-visible` state on every interactive element
- [ ] Text contrast ≥ WCAG AA; don't rely on color alone to convey meaning
- [ ] Images have `alt`; decorative images `alt=""`
- [ ] Forms have associated `<label>`s
- [ ] `prefers-reduced-motion` honored for non-essential motion
- [ ] Keyboard-operable: logical tab order, no traps

## Workflow

1. **Clarify the direction.** Confirm brand, tone, and any existing tokens or
   references. If the project has a brand/design doc, read it first.
2. **Set tokens.** Establish spacing, type scale, neutrals, and one accent.
3. **Block out layout.** Structure and hierarchy before decoration.
4. **Style with the tokens.** Components, all states, responsive.
5. **Polish pass.** Spacing rhythm, type detail, motion, empty/loading/error
   states.
6. **A11y + verify.** Run the checklist; view the result in a browser and at
   mobile width before calling it done.

## Anti-patterns to avoid

- Default framework look with no intentional choices
- More than 2 type families or a jumble of font sizes off-scale
- Accent color applied everywhere until nothing stands out
- Shadows on everything; borders + shadows stacked without purpose
- Center-aligned long paragraphs
- Interactive elements with no hover/focus states
- Fixed pixel widths that break on mobile
- Motion that's slow, bouncy, or fires on every scroll

## Verifying the result

Don't declare a UI done from the code alone. Open it in a browser, resize to a
narrow width, tab through it with the keyboard, and check dark mode if
supported. If a `verify` or `run` skill is available in the project, use it to
drive the actual page.
