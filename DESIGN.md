---
name: TrioBridge AI / 三桥 AI
description: A compact document proofing workspace for Word, Excel, and PowerPoint.
colors:
  ink: "#233b37"
  muted: "#64736e"
  accent: "#126b61"
  accent-hover: "#0c544b"
  accent-soft: "#e8f2ee"
  surface: "#fff"
  canvas: "#f7f9f7"
  line: "#dae3de"
  error: "#a73030"
  error-bg: "#fff0ed"
  error-hover: "#812727"
  page-background: "#edf2f0"
  sidebar: "#eef5f0"
  nav-text: "#536b5d"
  nav-hover: "#e1ede5"
  utility-active: "#dfebe2"
  input-border: "#bdcdc3"
  placeholder: "#6d7b73"
  selected-border: "#81ad98"
  focus: "#d0a23f"
  result-empty: "#f5f8f5"
  result-ready: "#fbfcfa"
  pending: "#80631d"
  success: "#255c40"
  success-bg: "#e9f3ec"
typography:
  headline:
    fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', 'Noto Sans SC', sans-serif"
    fontSize: "28px"
    fontWeight: 650
    lineHeight: 1.35
    letterSpacing: "-0.025em"
  headline-narrow:
    fontSize: "24px"
    fontWeight: 650
    lineHeight: 1.35
    letterSpacing: "-0.025em"
  title:
    fontSize: "16px"
    fontWeight: 650
    lineHeight: 1.45
  dialog-title:
    fontSize: "22px"
    fontWeight: 650
    lineHeight: 1.45
  body:
    fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', 'Noto Sans SC', sans-serif"
    fontSize: "14px"
    fontWeight: 400
    lineHeight: 1.6
  content:
    fontSize: "13px"
    lineHeight: 1.95
  label:
    fontSize: "13px"
    fontWeight: 500
    lineHeight: 1.6
  button:
    fontSize: "14px"
    fontWeight: 550
    lineHeight: 1.4
  help:
    fontSize: "12px"
    fontWeight: 400
    lineHeight: 1.7
  metadata:
    fontSize: "11px"
    lineHeight: 1.6
  formula:
    fontFamily: "ui-monospace, SFMono-Regular, Consolas, monospace"
    fontSize: "16px"
    lineHeight: 1.95
rounded:
  badge: "6px"
  task: "7px"
  control: "8px"
  navigation: "9px"
  model-strip: "10px"
  radius: "12px"
  dialog: "16px"
  shell: "18px"
spacing:
  "7": "7px"
  "8": "8px"
  "10": "10px"
  "12": "12px"
  "14": "14px"
  "16": "16px"
  "18": "18px"
  "20": "20px"
  "24": "24px"
  "26": "26px"
  "28": "28px"
  "34": "34px"
components:
  button-primary:
    backgroundColor: "{colors.accent}"
    textColor: "{colors.surface}"
    typography: "{typography.button}"
    rounded: "{rounded.control}"
    padding: "11px 16px"
  button-primary-hover:
    backgroundColor: "{colors.accent-hover}"
  button-secondary:
    backgroundColor: "{colors.surface}"
    textColor: "{colors.ink}"
    typography: "{typography.button}"
    rounded: "{rounded.control}"
    padding: "11px 16px"
  button-secondary-hover:
    backgroundColor: "{colors.accent-soft}"
  button-danger:
    backgroundColor: "{colors.error}"
    textColor: "{colors.surface}"
    typography: "{typography.button}"
    rounded: "{rounded.control}"
    padding: "11px 16px"
  button-danger-hover:
    backgroundColor: "{colors.error-hover}"
  button-text:
    backgroundColor: "transparent"
    textColor: "{colors.accent}"
    padding: "8px 2px"
  button-text-hover:
    textColor: "{colors.accent-hover}"
  button-icon:
    backgroundColor: "transparent"
    textColor: "{colors.muted}"
    rounded: "{rounded.control}"
    width: "40px"
    height: "40px"
  button-icon-hover:
    backgroundColor: "{colors.accent-soft}"
    textColor: "{colors.accent}"
  input:
    backgroundColor: "{colors.surface}"
    textColor: "{colors.ink}"
    rounded: "{rounded.control}"
    padding: "10px 12px"
  input-disabled:
    backgroundColor: "{colors.canvas}"
    textColor: "{colors.muted}"
  navigation:
    backgroundColor: "transparent"
    textColor: "{colors.nav-text}"
    rounded: "{rounded.navigation}"
    padding: "10px 12px"
  navigation-active:
    backgroundColor: "{colors.accent}"
    textColor: "{colors.surface}"
  task-option:
    backgroundColor: "{colors.surface}"
    textColor: "#54655b"
    rounded: "{rounded.task}"
    padding: "9px 11px"
  task-option-selected:
    backgroundColor: "{colors.accent-soft}"
    textColor: "{colors.accent}"
  result-card:
    backgroundColor: "{colors.result-empty}"
    textColor: "{colors.ink}"
    rounded: "{rounded.radius}"
    padding: "18px 20px"
  result-card-ready:
    backgroundColor: "{colors.result-ready}"
  comparison-region:
    backgroundColor: "{colors.canvas}"
    textColor: "{colors.ink}"
    rounded: "{rounded.control}"
    padding: "12px"
---

# Design System: TrioBridge AI / 三桥 AI

## Overview

**Creative North Star: "Document proofing workspace"**

This is an implementation interpretation of the current interface, not a user-approved brand direction. Pale sage surroundings, deep green actions, and compact system typography keep document inputs and their previews easy to distinguish. The interface uses ordinary labeled controls and visible pending/applied states.

The confirmed identity is TrioBridge AI / 三桥 AI with Chinese as the primary interface language. The user confirmed Mac and Windows, including perpetual Office; that requirement is not evidence of host compatibility testing. Visual source of truth: [src/styles.css](src/styles.css), [App.tsx](src/App.tsx), and [components](src/components). The composition contract remains in [.impeccable/surfaces/src-app-tsx.md](.impeccable/surfaces/src-app-tsx.md).

**Key Characteristics:**

- Compact system typography with roomy multiline document text.
- Pale tonal containers, thin borders, and deep green actions.
- Explicit selection, pending review, applied, error, and focus states.
- Desktop columns that become a vertical workflow in narrow panes.

The frontmatter records reused CSS values. Unprefixed color names match root custom properties; descriptive aliases cover repeated literals. Spacing and most radius names are documentation aliases, not existing CSS variables. The sidecar adds motion, breakpoints, depth, and component previews; no tonal ramps have been invented.

## Colors

### Primary

`accent` is the deep green of primary actions, active host navigation, links, and applied status. `accent-hover` darkens interactive actions; `accent-soft` supports selected tasks and secondary-control hover states.

### Neutral

`ink` carries main text; `muted` supports labels and guidance. White `surface` sits above the pale `page-background`, `sidebar`, and `canvas`. `line` divides sections, while `input-border` gives fields a stronger edge. Result containers use `result-empty` and the lighter `result-ready`.

Pending review uses `pending`; successful notices use `success` on `success-bg`; errors use `error` on `error-bg`. The gold `focus` outline is a keyboard cue. These are semantic states, not additional brand accents.

**The Visible State Rule.** Retain the text and icon cues accompanying color: selected tasks have a check, and results say “待确认” or “已写入”.

## Typography

The system stack resolves to the platform UI font, with Segoe UI and Noto Sans SC fallbacks; there is no bundled display face. The hierarchy is practical rather than a mathematical type scale.

- `headline` is the page heading; `headline-narrow` applies at the narrow breakpoint.
- `title` is the section heading, with `dialog-title` for the write confirmation.
- `body` establishes the inherited text. Controls and settings labels use the smaller `label`; document output uses the airy `content` line height.
- `help` and `metadata` carry guidance, provenance, and counts. Counts use tabular numerals where implemented.
- `formula` uses a monospace stack and the accent color. General document previews retain the system face even inside a `pre` element.

## Layout

The desktop shell has a maximum width of 1260px, a 212px navigation column, and 32px vertical outer margins. At widths of at least 1100px, horizontal margins are at least 24px. Main content uses 33px by 34px padding. The work area splits input and preview into `1fr / 0.96fr` columns with a 28px gap.

At 1099px and below the shell fills the viewport, loses its outer border and rounding, and uses a 180px navigation column; content gutters become 24px. At 850px and below the workspace stacks input above preview. At 600px and below navigation moves above the content, host buttons form a row, utility navigation becomes underlined tabs, and content gutters become 18px. The CSS body minimum is 280px. Settings remain a single column with a 620px maximum width.

Task options stay in two columns. Long result text and target labels wrap. Confirmation comparisons stack at 600px; each text region scrolls independently and remains keyboard-focusable. The modal itself is bounded by the viewport and can scroll.

## Elevation & Depth

Resting surfaces use tonal separation and one-pixel borders without shadows. The confirmation dialog is the exception: its shadow is `0 18px 70px #163c3433` and its backdrop is `#183b3266`. These exact depth values live in the sidecar.

Button background and text transitions use the root `--motion` value. Loading indicators rotate at one second per revolution. The reduced-motion media query reduces animation and transition duration and limits animation to one iteration.

## Shapes

Controls use the `control` radius, task choices the slightly smaller `task` radius, and host navigation the `navigation` radius. The model strip, result container, dialog, and desktop shell step up in rounding. Borders, not lift or scale effects, define control edges. The small connection dot is circular; rectangular task options are not pills.

## Components

### Buttons

Primary, secondary, and danger actions share a minimum height of 44px, the `button` type role, and the frontmatter padding. Secondary actions keep a thin `line` border. Text actions use a 36px minimum height and underline on hover; icon buttons are 40px squares. Disabled buttons use opacity 0.48 and a not-allowed cursor. The CSS has no separate pressed animation.

### Inputs / Fields

Inputs and selects have a 44px minimum height, `input-border`, and white fill. Textareas resize vertically and use a 1.85 line height. The source textarea has 14px padding and a 187px minimum height, reduced to 165px in narrow panes. Disabled fields use muted text on `canvas`; errors appear in labeled notices rather than an invented red field-border variant.

Model connection starts with a provider selector and API key. Presets fill the service name, endpoint and model without sending a request; the destination remains visible as wrapping help text. A single “连接并保存” action tests before saving. Advanced settings disclose editable names, endpoints, model IDs, timeouts and an offline save action. Custom profiles open these fields immediately; existing saved values remain intact. The saved-profile toolbar appears only after at least one profile exists. Changing the provider or endpoint clears its credential and remember-key choice.

### Navigation

Host buttons pair an icon with a name. The active host uses deep green with white content and a trailing chevron on desktop; selected disabled host buttons retain full opacity. Utility navigation uses a pale green active fill on desktop and an accent underline in narrow panes. Existing disabled and hover treatments remain distinct from selection.

### Task choices and status

Task choices use outlined rectangular buttons, a soft green selected fill, the `selected-border`, stronger text weight, and a check icon. A pending result label becomes an applied label only after writing. Connection and result badges are informational, not clickable filters.

### Cards / Containers

The preview is a bordered container using `--radius`. It changes from the empty to the ready surface when content exists and from centered empty guidance to document text and actions. Its keyboard-visible outline uses the selected-border color at 2px with a 3px offset; standard controls instead use the gold 3px outline with a 3px offset.

### Write confirmation

The native dialog shows the named target and before/after content before the action row. Comparison regions have 100px minimum and 260px maximum heights, reduced to a 170px maximum in narrow panes. Their `tabIndex=0` and gold focus outline make scrollable content reachable from the keyboard. The dialog initially focuses “返回检查”, cycles through its focus targets, and restores focus on close. Busy and error states govern the action buttons.

## Do's and Don'ts

### Do:

- **Do** reuse the root CSS custom properties and existing component variants.
- **Do** retain labels, icons, and focus outlines alongside color states.
- **Do** preserve the input-to-preview order when the workspace stacks.
- **Do** keep long confirmation content scrollable and keyboard-focusable.

### Don't:

- **Don't** describe this extracted aesthetic as a user-approved brand system.
- **Don't** replace document preview text with monospace unless it is a formula.
- **Don't** present pending output with the applied-state label before a write succeeds.
- **Don't** infer Office host verification from the browser layout or screenshots.
