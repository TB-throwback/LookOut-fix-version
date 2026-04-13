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
  ["../src/icons/LOicon-32.png", "icons/LOicon-32.png"],
  ["../src/icons/LOicon-48.png", "icons/LOicon-48.png"],
  ["../src/icons/LOicon-64.png", "icons/LOicon-64.png"],
  {
    src: "app.js",
    dest: "app.js",
    replace: ["../src/scripts/lookout.mjs", "./scripts/lookout.mjs"],
  },
  {
    src: "../src/scripts/lookout.mjs",
    dest: "scripts/lookout.mjs",
    replace: ["/scripts/tnef.mjs", "./tnef.mjs"],
  },
  ["../src/scripts/tnef.mjs", "scripts/tnef.mjs"],
];

for (const file of filesToCopy) {
  const src = path.join(root, file.src || file[0]);
  const dest = path.join(dist, file.dest || file[1]);
  await mkdir(path.dirname(dest), { recursive: true });
  if (file.replace) {
    let content = await readFile(src, "utf8");
    content = content.replaceAll(file.replace[0], file.replace[1]);
    await writeFile(dest, content, "utf8");
  } else {
    await copyFile(src, dest);
  }
}
