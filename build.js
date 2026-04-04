// @ts-check
"use strict";

const esbuild = require("esbuild");
const fs = require("fs");
const path = require("path");
const http = require("http");
const https = require("https");

const ROOT = __dirname;
const DIST = path.join(ROOT, "dist");
const PORT = 3000;

// ── アイコン生成 ──────────────────────────────────────────────────────────────
// 最小の 1×1 PNG プレースホルダー（assets/ が空の場合のみ生成）
const PLACEHOLDER_PNG = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAAAAAA6fptVAAAACklEQVQI12NgAAAAAgAB4iG8MwAAAABJRU5ErkJggg==",
  "base64",
);

function ensureIcons() {
  const dir = path.join(ROOT, "assets");
  fs.mkdirSync(dir, { recursive: true });
  for (const size of [16, 32, 80]) {
    const file = path.join(dir, `icon-${size}.png`);
    if (!fs.existsSync(file)) {
      fs.writeFileSync(file, PLACEHOLDER_PNG);
      console.log(`  created placeholder: assets/icon-${size}.png`);
    }
  }
}

// ── 静的ファイルのコピー ───────────────────────────────────────────────────────
function copyStatics() {
  fs.mkdirSync(DIST, { recursive: true });

  // HTML
  fs.copyFileSync(
    path.join(ROOT, "src/taskpane/taskpane.html"),
    path.join(DIST, "taskpane.html"),
  );

  // assets/
  const srcAssets = path.join(ROOT, "assets");
  const dstAssets = path.join(DIST, "assets");
  fs.mkdirSync(dstAssets, { recursive: true });
  for (const f of fs.readdirSync(srcAssets)) {
    fs.copyFileSync(path.join(srcAssets, f), path.join(dstAssets, f));
  }
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

// ── エントリポイント ───────────────────────────────────────────────────────────
const [, , mode = "build"] = process.argv; // node build.js [build|watch|serve]

(async () => {
  ensureIcons();
  copyStatics();

  if (mode === "serve") {
    // ── esbuild の HTTP サーバ + HTTPS プロキシ ────────────────────────────
    const ctx = await esbuild.context(BUNDLE);
    const { port: ebPort } = await ctx.serve({ servedir: DIST });
    const host = "localhost";

    let tlsOpts;
    try {
      const devCerts = require("office-addin-dev-certs");
      tlsOpts = await devCerts.getHttpsServerOptions();
    } catch {
      console.error(
        "\n  [エラー] HTTPS 証明書が見つかりません。先に以下を実行してください:",
      );
      console.error("  npx office-addin-dev-certs install\n");
      process.exit(1);
    }

    https
      .createServer(tlsOpts, (req, res) => {
        const proxy = http.request(
          {
            hostname: host,
            port: ebPort,
            path: req.url,
            method: req.method,
            headers: req.headers,
          },
          (r) => {
            res.writeHead(r.statusCode ?? 200, r.headers);
            r.pipe(res, { end: true });
          },
        );
        proxy.on("error", (e) => {
          console.error(e);
          res.writeHead(502);
          res.end();
        });
        req.pipe(proxy, { end: true });
      })
      .listen(PORT, () => {
        console.log(`\n  Mermaid for Excel:`);
        console.log(`  https://localhost:${PORT}/taskpane.html`);
        console.log(
          `\n  Excel の [挿入] → [アドイン] → [アドインをアップロード] で manifest.xml を選択\n`,
        );
      });
  } else if (mode === "watch") {
    // ── ファイル監視モード ────────────────────────────────────────────────
    const ctx = await esbuild.context({ ...BUNDLE, sourcemap: "inline" });
    await ctx.watch();
    console.log("  watch モード開始 (Ctrl+C で停止)");
  } else {
    // ── ワンタイムビルド ──────────────────────────────────────────────────
    const ctx = await esbuild.context(BUNDLE);
    await ctx.rebuild();
    await ctx.dispose();
    console.log("  build 完了 → dist/");
  }
})().catch((e) => {
  console.error(e);
  process.exit(1);
});
