const { test, before } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("fs");
const path = require("path");
const { execSync } = require("child_process");

const ROOT = path.resolve(__dirname, "..");
const HTML_PATH = path.join(ROOT, "docs", "user-guide.html");

before(() => {
  execSync("node scripts/build-user-guide.js", { cwd: ROOT, stdio: "pipe" });
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
  const { buildUserGuideHtml } = require("./build-user-guide");
  const html = buildUserGuideHtml();
  assert.match(html, /<div class="wrap">/);
  assert.match(html, /Generated from <code>docs\/user-guide\.md<\/code>/);
});
