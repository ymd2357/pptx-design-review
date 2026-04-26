# PPTX-specific Fix Guide

This document explains concrete fix techniques for PPTX slide decks
(PowerPoint / Google Slides / Keynote) rather than web frameworks.

---

## Layout & Alignment

### Fixing Overlap

- Use **Align** and **Distribute** to normalize positions.
- Prefer **Slide Master** adjustments if the issue repeats across slides.
- Reduce text box size or split content across slides.

### Grid Consistency

- Define a **baseline grid** (e.g., 8pt) and snap objects to it.
- Reuse **layout templates** to keep margins consistent.

---

## Text Handling

### Text Clipping

- Increase text box height or reduce font size by 1–2pt.
- Use **auto-fit** only if it does not distort hierarchy.
- Convert long paragraphs into bullet lists across multiple slides.

### Typography Consistency

- Centralize fonts in **Slide Master**.
- Use 2–3 font sizes for hierarchy (Title / Subtitle / Body).
- Avoid mixing font families within the same deck.

---

## Images & Media

### Aspect Ratio Distortion

- Always **crop** images instead of stretching.
- Lock aspect ratio during resize.

### Low Resolution

- Replace with higher-resolution assets.
- Avoid scaling up images above 100%.
- For print output, ensure source assets are 300dpi equivalent.

### Embedded Media

- Prefer static frames if playback is uncertain.
- If video is required, confirm playback environment (PowerPoint/Keynote) and codecs.

---

## Color & Consistency

### Palette Unification

- Define 1–2 brand primaries and 1–2 neutrals.
- Replace outlier colors with nearest palette values.

### Contrast

- Increase text contrast against background.
- Avoid placing text over complex imagery without a solid overlay.

---

## Component Consistency

### Repeated Elements

- Convert to **Slide Master** or duplicate from a canonical slide.
- Ensure identical spacing, size, and alignment for repeating components.

---

## Export / Delivery

### File Format Validation

- If output is PDF, verify fonts are embedded or substituted predictably.
- Test on a second machine to catch missing fonts.

---

## Fix Principles

1. **Minimal Changes**: Fix only what is necessary for clarity.
2. **Respect Existing Patterns**: Use the deck’s current style system.
3. **Avoid Breaking Slides**: Confirm fixes don’t introduce new issues.
4. **Document Changes**: Record slide numbers and fixes.
