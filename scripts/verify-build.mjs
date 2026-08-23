import assert from "node:assert/strict";
import { existsSync, readFileSync, readdirSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const outputRoot = path.join(projectRoot, "dist");

assert.ok(existsSync(outputRoot), "dist must exist before running build verification");

const files = [];
function walk(directory) {
  for (const entry of readdirSync(directory, { withFileTypes: true })) {
    const entryPath = path.join(directory, entry.name);
    if (entry.isDirectory()) walk(entryPath);
    else files.push(entryPath);
  }
}
walk(outputRoot);

const htmlFiles = files.filter((file) => file.endsWith(".html"));
assert.ok(htmlFiles.length > 0, "build must contain HTML pages");

for (const required of ["robots.txt", "sitemap-index.xml", "sitemap-0.xml"]) {
  assert.ok(existsSync(path.join(outputRoot, required)), `Missing ${required}`);
}

const robots = readFileSync(path.join(outputRoot, "robots.txt"), "utf8");
assert.match(robots, /Sitemap: https:\/\/healeycottage\.com\/sitemap-index\.xml/);
const sitemap = readFileSync(path.join(outputRoot, "sitemap-0.xml"), "utf8");
assert.doesNotMatch(sitemap, /\/request\/success\/?</, "success page must not appear in sitemap");

for (const file of htmlFiles) {
  const html = readFileSync(file, "utf8");
  const relative = path.relative(outputRoot, file);
  const route = relative === "index.html"
    ? "/"
    : `/${relative.replace(/index\.html$/, "").replaceAll(path.sep, "/")}`;
  const expectedCanonical = new URL(route, "https://healeycottage.com").href;

  assert.ok(
    html.includes(`<link rel="canonical" href="${expectedCanonical}">`),
    `${relative} has an incorrect or missing canonical URL`
  );
  assert.ok(
    html.includes(`<meta property="og:url" content="${expectedCanonical}">`),
    `${relative} has an incorrect or missing Open Graph URL`
  );
  for (const property of ["og:title", "og:description", "og:image"]) {
    assert.ok(html.includes(`<meta property="${property}"`), `${relative} is missing ${property}`);
  }
  assert.ok(html.includes('type="application/ld+json"'), `${relative} is missing JSON-LD`);
  assert.equal(
    [...html.matchAll(/<main(?:\s|>)/g)].length,
    1,
    `${relative} must contain exactly one main landmark`
  );
  assert.doesNotMatch(html, /localhost|127\.0\.0\.1|0\.0\.0\.0/);

  for (const match of html.matchAll(/\b(?:href|src)="([^"#]+)"/g)) {
    const target = match[1];
    if (/^(?:https?:|mailto:|tel:|data:|javascript:)/i.test(target)) continue;

    const targetPath = target.split("?")[0];
    const resolved = targetPath.startsWith("/")
      ? path.join(outputRoot, targetPath)
      : path.resolve(path.dirname(file), targetPath);
    const candidates = [resolved, path.join(resolved, "index.html"), `${resolved}.html`];
    assert.ok(
      candidates.some(existsSync),
      `${relative} references missing internal resource ${target}`
    );
  }
}

const successHtml = readFileSync(path.join(outputRoot, "request/success/index.html"), "utf8");
assert.ok(
  successHtml.includes('<meta name="robots" content="noindex, nofollow">'),
  "request success page must not be indexed"
);
const requestHtml = readFileSync(path.join(outputRoot, "request/index.html"), "utf8");
const formAction = requestHtml.match(/<form[^>]+action="([^"]+)"/i)?.[1];
assert.ok(formAction, "request form must include a submission endpoint");
assert.equal(new URL(formAction).protocol, "https:", "request form endpoint must use HTTPS");
assert.equal(files.filter((file) => file.endsWith(".map")).length, 0, "source maps must not ship");

console.log(`Verified ${htmlFiles.length} HTML pages and ${files.length} generated files.`);
