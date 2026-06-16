const { test, before } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("fs");
const path = require("path");
const { execSync } = require("child_process");

const ROOT = path.resolve(__dirname, "..");
const HTML_PATH = path.join(ROOT, "docs", "user-guide.html");

before(() => {
  execSync("node scripts/build-docs.js", { cwd: ROOT, stdio: "pipe" });
});

test("writes user-guide.html", () => {
  assert.ok(fs.existsSync(HTML_PATH), "expected docs/user-guide.html");
});

test("includes document title and key guide content", () => {
  const html = fs.readFileSync(HTML_PATH, "utf8");
  assert.match(html, /<title>MT Data Connector — User Guide<\/title>/);
  assert.match(html, /GETSOURCES/);
  assert.match(html, /RELOADFUNCTIONS/);
  assert.match(html, /First-time setup/);
});

test("renders markdown tables as HTML", () => {
  const html = fs.readFileSync(HTML_PATH, "utf8");
  assert.match(html, /<table>/);
  assert.match(html, /<th>Function<\/th>/);
  assert.match(html, /<th>What it does<\/th>/);
});

test("exports build function for programmatic use", () => {
  const { buildUserGuideHtml } = require("./build-docs");
  const html = buildUserGuideHtml();
  assert.match(html, /<div class="wrap">/);
  assert.match(html, /Generated from markdown at build time/);
});

test("writes support and privacy HTML pages", () => {
  const supportPath = path.join(ROOT, "docs", "support.html");
  const privacyPath = path.join(ROOT, "docs", "privacy-policy.html");
  assert.ok(fs.existsSync(supportPath), "expected docs/support.html");
  assert.ok(fs.existsSync(privacyPath), "expected docs/privacy-policy.html");
  const supportHtml = fs.readFileSync(supportPath, "utf8");
  assert.match(supportHtml, /<title>MT Data Connector — Support<\/title>/);
});
