# pptx2html

A tool for converting PowerPoint (PPTX) files to HTML5.

## Structure

This is an npm workspaces monorepo:

- `packages/` — library code with no direct runtime output (parsers, converters, shared utilities).
- `apps/` — runnable code (web pages, Electron apps, Node CLIs) that consume the packages.

## Tooling

- TypeScript
- webpack (bundling for apps)
- ESLint (linting)
- Prettier (formatting)
- Vitest (testing)

## Getting started

```bash
npm install
npm run build
npm run lint
npm run format
npm test
```
