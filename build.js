// @ts-check
"use strict";

const esbuild = require("esbuild");
const fs = require("fs");
const path = require("path");

const ROOT = __dirname;
const DIST = path.join(ROOT, "dist");
const PORT = 3000;

const DEV_BASE_URL = "https://localhost:3000";
const PROD_BASE_URL = "https://msmsrep.github.io/TextToSvgDiagramsForExcel/dist";

// ── 静的ファイルのコピー ───────────────────────────────────────────────────────
function copyStatics() {
  fs.mkdirSync(DIST, { recursive: true });

  // HTML / CSS
  for (const f of ["taskpane.html", "taskpane.css"]) {
    fs.copyFileSync(path.join(ROOT, "src/taskpane", f), path.join(DIST, f));
  }

  // assets/
  const srcAssets = path.join(ROOT, "assets");
  const dstAssets = path.join(DIST, "assets");
  fs.mkdirSync(dstAssets, { recursive: true });
  for (const f of fs.readdirSync(srcAssets)) {
    fs.copyFileSync(path.join(srcAssets, f), path.join(dstAssets, f));
  }
}

// ── manifest.xml の URL を書き換えて dist/ へ出力 ────────────────────────────
/**
 * @param {string} baseUrl
 */
function copyManifest(baseUrl) {
  const src = fs.readFileSync(path.join(ROOT, "manifest.xml"), "utf8");
  const out = src.replaceAll(DEV_BASE_URL, baseUrl);
  fs.writeFileSync(path.join(DIST, "manifest.xml"), out, "utf8");
}

// ── esbuild 共通オプション ────────────────────────────────────────────────────
/** @type {import('esbuild').BuildOptions} */
const BUNDLE = {
  entryPoints: [path.join(ROOT, "src/taskpane/taskpane.ts")],
  bundle: true,
  outfile: path.join(DIST, "taskpane.js"),
  format: "iife",
  target: "es2017",
  sourcemap: true,
};

// ── 各モード ──────────────────────────────────────────────────────────────────
/**
 * ワンタイムビルド
 * @param {string} baseUrl manifest.xml に埋め込む URL
 */
async function buildOnce(baseUrl) {
  copyManifest(baseUrl);
  await esbuild.build(BUNDLE);
  console.log("  build 完了 → dist/");
}

/** ファイル監視モード */
async function watch() {
  copyManifest(DEV_BASE_URL);
  const ctx = await esbuild.context({ ...BUNDLE, sourcemap: "inline" });
  await ctx.watch();
  console.log("  watch モード開始 (Ctrl+C で停止)");
}

/** esbuild の HTTPS 開発サーバ */
async function serve() {
  copyManifest(DEV_BASE_URL);

  /** @type {{ localhostCertificatePath: string, localhostKeyPath: string }} */
  let certs;
  try {
    // 証明書がインストール済みで有効なことを確認する
    await require("office-addin-dev-certs").getHttpsServerOptions();
    certs = require("office-addin-dev-certs/lib/defaults");
  } catch {
    console.error(
      "\n  [エラー] HTTPS 証明書が見つかりません。先に以下を実行してください:",
    );
    console.error("  npx office-addin-dev-certs install\n");
    process.exit(1);
  }

  const ctx = await esbuild.context(BUNDLE);
  await ctx.serve({
    servedir: DIST,
    port: PORT,
    certfile: certs.localhostCertificatePath,
    keyfile: certs.localhostKeyPath,
  });
  console.log(`\n  Text-to-SVG Diagrams for Excel:`);
  console.log(`  ${DEV_BASE_URL}/taskpane.html`);
  console.log(
    `\n  Excel の [挿入] → [アドイン] → [アドインをアップロード] で manifest.xml を選択\n`,
  );
}

/** @type {Record<string, () => Promise<void>>} */
const MODES = {
  build: () => buildOnce(PROD_BASE_URL),
  dev: () => buildOnce(DEV_BASE_URL),
  watch,
  serve,
};

// ── エントリポイント ───────────────────────────────────────────────────────────
const [, , mode = "build"] = process.argv; // node build.js [build|dev|watch|serve]

(async () => {
  const run = MODES[mode];
  if (!run) {
    throw new Error(`Unknown mode: ${mode} (${Object.keys(MODES).join("|")})`);
  }
  copyStatics();
  await run();
})().catch((e) => {
  console.error(e);
  process.exit(1);
});
