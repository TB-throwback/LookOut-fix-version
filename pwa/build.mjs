import {
  mkdir,
  copyFile,
  rm,
  readFile,
  writeFile,
  readdir,
  stat,
} from "node:fs/promises";
import path from "node:path";

const root = path.resolve(path.dirname(new URL(import.meta.url).pathname));
const dist = path.join(root, "dist");

await rm(dist, { recursive: true, force: true });

const filesToCopy = [
  ["index.html", "index.html"],
  ["styles.css", "styles.css"],
  ["sw.js", "sw.js"],
  ["manifest.webmanifest", "manifest.webmanifest"],
  ["screenshots/", "screenshots/"],
  ["../src/scripts/mapi_props.js", "mapi_props.js"],
  ["../src/icons/", "icons/"],
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

async function copyRecursive(src, dest) {
  const srcStat = await stat(src);
  if (srcStat.isDirectory()) {
    await mkdir(dest, { recursive: true });
    const entries = await readdir(src);
    for (const entry of entries) {
      const srcEntry = path.join(src, entry);
      const destEntry = path.join(dest, entry);
      await copyRecursive(srcEntry, destEntry);
    }
  } else {
    await copyFile(src, dest);
  }
}

for (const file of filesToCopy) {
  const src = path.join(root, file.src || file[0]);
  const dest = path.join(dist, file.dest || file[1]);
  await mkdir(path.dirname(dest), { recursive: true });
  if (file.replace) {
    let content = await readFile(src, "utf8");
    content = content.replaceAll(file.replace[0], file.replace[1]);
    await writeFile(dest, content, "utf8");
  } else {
    const srcStat = await stat(src);
    if (srcStat.isDirectory()) {
      await copyRecursive(src, dest);
    } else {
      await copyFile(src, dest);
    }
  }
}
