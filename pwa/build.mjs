import {
  mkdir,
  copyFile,
  rm,
  readFile,
  writeFile,
  readdir,
  stat,
} from "node:fs/promises";
import { watch as fsWatch } from "node:fs";
import path from "node:path";

const root = path.resolve(path.dirname(new URL(import.meta.url).pathname));
const dist = path.join(root, "dist");
const watchMode = process.argv.includes("--watch");

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

const watchTargets = filesToCopy.map((file) =>
  path.join(root, file.src || file[0]),
);

async function build() {
  // await rm(dist, { recursive: true, force: true });

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
  console.log("Build complete.");
}

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

async function watch() {
  let rebuilding = false;
  let pending = false;
  let rebuildTimeout;

  const triggerRebuild = async () => {
    if (rebuildTimeout) {
      pending = true;
      return;
    }

    // Debounce: wait 100ms for additional changes before rebuilding
    rebuildTimeout = setTimeout(async () => {
      rebuildTimeout = null;

      if (rebuilding) {
        pending = true;
        return;
      }

      rebuilding = true;
      do {
        pending = false;
        await build();
      } while (pending);
      rebuilding = false;
    }, 100);
  };

  const watchers = [];
  for (const target of watchTargets) {
    try {
      const watcher = fsWatch(target, { persistent: true }, () => {
        void triggerRebuild();
      });
      watchers.push(watcher);
    } catch (error) {
      console.error(`Failed to watch ${target}:`, error);
    }
  }

  const stop = () => {
    for (const watcher of watchers) {
      watcher.close();
    }
    if (rebuildTimeout) {
      clearTimeout(rebuildTimeout);
    }
  };

  process.once("SIGINT", stop);
  process.once("SIGTERM", stop);
}

await build();

if (watchMode) {
  await watch();
}
