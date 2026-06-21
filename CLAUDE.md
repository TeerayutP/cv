# CV Project

Single-file static CV for Teerayut Panyoheang, served as a plain HTML fragment.

## Structure

```
index.html   — entire CV: inline CSS + markup, no build step
```

## Dev

Uses VS Code Live Server on port **5533**.  
Open `index.html` with Live Server or any static file server — no install required.

## Architecture

- **No framework, no bundler, no package.json.** Pure HTML + CSS.
- Full HTML document: `<!DOCTYPE html>`, `<html lang="en">`, `<head>`, `<body>`.
- All styles are inline `<style>` in `<head>`.
- Fonts loaded via `<link>` tags (not `@import`) with `preconnect` hints for faster first paint. `DM Sans` (body) and `DM Serif Display` (name heading).
- **Mobile-first CSS**: base styles target mobile, `@media (min-width: 420px)` and `@media (min-width: 600px)` add desktop layout.
- **Dark mode**: `@media (prefers-color-scheme: dark)` overrides `:root` tokens and badge colors. No JS required.
- CSS custom properties defined in `:root` — all color/background values go through tokens.
- Print-ready: `.print-btn` hides on print, `print-color-adjust: exact` preserves badge colors, `@page { margin: 1.5cm }` for consistent PDF margins.

## Constraints

- **Do not change content** — copy, bullet points, dates, and job details are owner-managed.
- Keep the single-file approach unless there is a clear reason to split.
- Print/PDF output must remain clean — test `@media print` after any layout change.
