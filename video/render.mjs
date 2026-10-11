// Renders demo.html to a video, frame by frame.
//   node render.mjs he            -> demo-he.webm + demo-he-poster.webp
//   node render.mjs he 1.5 6 9    -> stills at those seconds only (for checking)
// Needs playwright-core (taken from the portfolio's node_modules), Edge and ffmpeg.
import { mkdirSync, rmSync, existsSync } from "fs";
import { execFileSync } from "child_process";
import { fileURLToPath, pathToFileURL } from "url";
import path from "path";

const { chromium } = await import(pathToFileURL("E:/Projects/Protfolio/node_modules/playwright-core/index.mjs").href);
const here = path.dirname(fileURLToPath(import.meta.url));
const [lang = "he", ...stills] = process.argv.slice(2);
const FPS = 30;

const browser = await chromium.launch({ executablePath: "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe" });
const page = await browser.newPage({ viewport: { width: 960, height: 620 }, deviceScaleFactor: 1.5 });
await page.goto(pathToFileURL(path.join(here, "demo.html")).href + `?lang=${lang}`);
await page.waitForFunction(() => window.ready);
await page.evaluate(() => document.fonts.ready);

if (stills.length) {
  for (const s of stills) {
    await page.evaluate((t) => window.render(t), Number(s));
    await page.screenshot({ path: path.join(here, `still-${lang}-${s}.png`) });
  }
} else {
  const dir = path.join(here, `frames-${lang}`);
  rmSync(dir, { recursive: true, force: true });
  mkdirSync(dir);
  const duration = await page.evaluate(() => window.DURATION);
  const total = Math.round(duration * FPS);
  for (let i = 0; i < total; i++) {
    await page.evaluate((t) => window.render(t), i / FPS);
    await page.screenshot({ path: path.join(dir, `${String(i).padStart(4, "0")}.png`) });
  }
  const out = path.join(here, `demo-${lang}.webm`);
  execFileSync("ffmpeg", ["-y", "-loglevel", "error", "-framerate", String(FPS), "-i", path.join(dir, "%04d.png"),
    "-c:v", "libvpx-vp9", "-b:v", "0", "-crf", "36", "-row-mt", "1", "-pix_fmt", "yuv420p", "-an", out]);
  await page.evaluate((t) => window.render(t), 9.5);
  await page.screenshot({ path: path.join(here, `demo-${lang}-poster.png`) });
  rmSync(dir, { recursive: true, force: true });
  console.log(`Wrote ${out}`);
}
await browser.close();
