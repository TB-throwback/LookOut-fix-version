import { mkdir, copyFile, rm, readFile, writeFile } from "node:fs/promises";
import path from "node:path";

const root = path.resolve(path.dirname(new URL(import.meta.url).pathname));
const dist = path.join(root, "dist");

await rm(dist, { recursive: true, force: true });

const filesToCopy = [
  ["index.html", "index.html"],
  ["styles.css", "styles.css"],
  ["sw.js", "sw.js"],
  ["manifest.webmanifest", "manifest.webmanifest"],
  ["../src/scripts/mapi_props.js", "mapi_props.js"],
  ["../src/scripts/lookout.mjs", "scripts/lookout.mjs"],
  ["../src/scripts/tnef.mjs", "scripts/tnef.mjs"],
  ["../src/icons/LOicon-32.png", "icons/LOicon-32.png"],
  ["../src/icons/LOicon-48.png", "icons/LOicon-48.png"],
  ["../src/icons/LOicon-64.png", "icons/LOicon-64.png"],
];

const appSource = await readFile(path.join(root, "app.js"), "utf8");
const appBuilt = appSource.replace(
  "../src/scripts/lookout.mjs",
  "./scripts/lookout.mjs",
);
await mkdir(path.join(dist), { recursive: true });
await writeFile(path.join(dist, "app.js"), appBuilt, "utf8");

for (const [srcRel, destRel] of filesToCopy) {
  const src = path.join(root, srcRel);
  const dest = path.join(dist, destRel);
  await mkdir(path.dirname(dest), { recursive: true });
  await copyFile(src, dest);
}
