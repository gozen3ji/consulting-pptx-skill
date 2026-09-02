import fs from "node:fs/promises";
import path from "node:path";
import { createRequire } from "node:module";

const defaultNodeModules =
  "/Users/kazuki/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules";
const moduleDir = process.env.PLAYWRIGHT_MODULE_DIR || process.env.NODE_PATH || defaultNodeModules;
const localRequire = createRequire(import.meta.url);
const load = (name) => {
  try {
    return localRequire(name);
  } catch {
    return createRequire(`${moduleDir}/`)(name);
  }
};
const { chromium } = load("playwright");

const input = process.argv[2] || "html-css/styleguide.html";
const outputDir = process.argv[3] || "examples/screenshots";

const root = new URL("..", import.meta.url).pathname;
const htmlPath = path.resolve(root, input);
const screenshotsDir = path.resolve(root, outputDir);

await fs.mkdir(screenshotsDir, { recursive: true });

const browser = await chromium.launch({ headless: true });
const page = await browser.newPage({
  viewport: { width: 1800, height: 1100 },
  deviceScaleFactor: 1,
});

await page.goto(`file://${htmlPath}`, { waitUntil: "networkidle" });

const slides = await page.$$(".slide");
const screenshots = [];

for (let i = 0; i < slides.length; i += 1) {
  const output = path.join(
    screenshotsDir,
    `${path.basename(htmlPath, ".html")}-slide-${String(i + 1).padStart(2, "0")}.png`,
  );
  await slides[i].screenshot({ path: output });
  const box = await slides[i].boundingBox();
  screenshots.push({
    slide: i + 1,
    screenshot: output,
    width: Math.round(box.width),
    height: Math.round(box.height),
  });
}

const overflow = await page.evaluate(() =>
  [...document.querySelectorAll(".slide *")]
    .filter((el) => {
      const style = getComputedStyle(el);
      if (style.display === "none" || style.visibility === "hidden") return false;
      return el.scrollWidth > el.clientWidth + 1 || el.scrollHeight > el.clientHeight + 1;
    })
    .map((el) => ({
      className: String(el.className),
      text: (el.textContent || "").trim().slice(0, 100),
      scrollWidth: el.scrollWidth,
      clientWidth: el.clientWidth,
      scrollHeight: el.scrollHeight,
      clientHeight: el.clientHeight,
    })),
);

await browser.close();

const report = {
  htmlPath,
  slideCount: slides.length,
  overflowCount: overflow.length,
  screenshots,
  overflow,
};

const reportPath = path.join(screenshotsDir, `${path.basename(htmlPath, ".html")}-qa.json`);
await fs.writeFile(reportPath, `${JSON.stringify(report, null, 2)}\n`);

console.log(JSON.stringify(report, null, 2));

if (overflow.length > 0) {
  process.exitCode = 1;
}
