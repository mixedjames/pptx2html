import { defineConfig, devices } from '@playwright/test';

/**
 * A deliberately minimal real-browser regression suite — see `packages/to-html5/CLAUDE.md`'s "Key
 * design decision: three behavioral scroll-mode bugs" for why this exists at all: three real bugs
 * this project actually shipped (a containing-block/letterbox sizing bug, a Web Animations
 * composite-order bug, a dead-scroll-zone bug) were each invisible to `happy-dom` — this package's
 * whole unit-test environment implements no real layout and no real Web Animations at all — and
 * were only ever found by driving a real Chromium by hand. This suite exists to keep them found
 * automatically instead. Scoped narrowly to those two failure *classes* (real layout/containment,
 * real WAAPI composite/timing behavior) rather than broad pixel-diffing visual regression — cheaper
 * to write and far less flaky, and it's what the actual bugs were.
 *
 * Chromium only, one project — this isn't testing cross-browser compatibility, just real-vs-fake
 * DOM/CSS/WAAPI behavior that `happy-dom` can't provide at all. `webServer` runs the existing dev
 * server (`apps/pages`'s own `npm start`, i.e. `webpack serve --mode development`) rather than a
 * separate static-file server — one existing command, no new tooling, and it already serves the
 * demo `.pptx` fixtures via the same `CopyWebpackPlugin` the production build uses.
 *
 * Deliberately **not** wired into `npm test`/CI yet — a real-browser suite is slower and a
 * different kind of investment than the rest of this repo's tests; run it explicitly via
 * `npm run test:e2e` (from this directory or the repo root).
 */
export default defineConfig({
  testDir: './e2e',
  fullyParallel: true,
  forbidOnly: !!process.env.CI,
  retries: process.env.CI ? 2 : 0,
  reporter: 'list',
  use: {
    baseURL: 'http://localhost:8081',
    trace: 'retain-on-failure',
  },
  projects: [{ name: 'chromium', use: { ...devices['Desktop Chrome'] } }],
  webServer: {
    command: 'npm start',
    url: 'http://localhost:8081',
    reuseExistingServer: !process.env.CI,
    timeout: 30_000,
  },
});
