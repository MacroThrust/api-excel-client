/**
 * Generates HTML documentation pages from markdown under docs/.
 * Used for GitHub Pages (user guide, support, privacy policy).
 */
const fs = require("fs");
const path = require("path");
const { marked } = require("marked");

const ROOT = path.resolve(__dirname, "..");
const DOCS_DIR = path.join(ROOT, "docs");

const PAGE_STYLES = `
    :root {
      color-scheme: light dark;
      --text: #242424;
      --muted: #605e5c;
      --border: #edebe9;
      --accent: #0078d4;
      --bg: #faf9f8;
      --card: #ffffff;
    }
    @media (prefers-color-scheme: dark) {
      :root {
        --text: #f3f2f1;
        --muted: #c8c6c4;
        --border: #3b3a39;
        --accent: #4da6ff;
        --bg: #1b1a19;
        --card: #252423;
      }
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: "Segoe UI", system-ui, -apple-system, sans-serif;
      line-height: 1.55;
      color: var(--text);
      background: var(--bg);
    }
    .wrap {
      max-width: 44rem;
      margin: 0 auto;
      padding: 2rem 1.25rem 3rem;
    }
    .wrap > :first-child { margin-top: 0; }
    h1 { font-size: 1.75rem; margin: 0 0 0.75rem; font-weight: 600; }
    h2 {
      font-size: 1.2rem;
      margin: 2rem 0 0.75rem;
      font-weight: 600;
      padding-top: 0.25rem;
      border-top: 1px solid var(--border);
    }
    .wrap > h2:first-of-type,
    .wrap > hr + h2 { border-top: none; padding-top: 0; }
    h3 { font-size: 1rem; margin: 1.25rem 0 0.5rem; }
    p, li { font-size: 0.95rem; }
    hr { border: none; border-top: 1px solid var(--border); margin: 1.5rem 0; }
    a { color: var(--accent); }
    code, pre {
      font-family: Consolas, "Courier New", monospace;
      font-size: 0.88em;
    }
    code {
      background: var(--border);
      padding: 0.1em 0.35em;
      border-radius: 3px;
    }
    pre {
      background: var(--card);
      border: 1px solid var(--border);
      padding: 0.75rem 1rem;
      overflow-x: auto;
      border-radius: 4px;
    }
    pre code { background: none; padding: 0; }
    table {
      width: 100%;
      border-collapse: collapse;
      font-size: 0.9rem;
      background: var(--card);
      border: 1px solid var(--border);
      border-radius: 4px;
      overflow: hidden;
      margin: 1rem 0;
    }
    th, td {
      text-align: left;
      padding: 0.5rem 0.65rem;
      border-bottom: 1px solid var(--border);
      vertical-align: top;
    }
    th { background: var(--border); font-weight: 600; }
    tr:last-child td { border-bottom: none; }
    footer.doc-footer {
      margin-top: 2.5rem;
      padding-top: 1rem;
      border-top: 1px solid var(--border);
      color: var(--muted);
      font-size: 0.85rem;
    }
`;

/** @type {{ md: string; title: string }[]} */
const DOC_PAGES = [
  { md: "user-guide.md", title: "MT Data Connector — User Guide" },
  { md: "support.md", title: "MT Data Connector — Support" },
  { md: "privacy-policy.md", title: "MT Data Connector — Privacy Policy" },
];

marked.setOptions({
  gfm: true,
  headerIds: true,
  mangle: false,
});

function buildDocHtml(title, markdown) {
  const content = marked.parse(markdown);

  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>${title}</title>
  <style>${PAGE_STYLES}
  </style>
</head>
<body>
  <div class="wrap">
    ${content}
    <footer class="doc-footer">
      <p>Generated from markdown at build time. <a href="user-guide.html">User guide</a> · <a href="support.html">Support</a> · <a href="privacy-policy.html">Privacy</a></p>
    </footer>
  </div>
</body>
</html>
`;
}

function buildUserGuideHtml() {
  const entry = DOC_PAGES[0];
  const markdown = fs.readFileSync(path.join(DOCS_DIR, entry.md), "utf8");
  return buildDocHtml(entry.title, markdown);
}

function main() {
  for (const { md, title } of DOC_PAGES) {
    const mdPath = path.join(DOCS_DIR, md);
    if (!fs.existsSync(mdPath)) {
      console.error(`Missing source file: ${mdPath}`);
      process.exit(1);
    }

    const htmlPath = path.join(DOCS_DIR, md.replace(/\.md$/, ".html"));
    const markdown = fs.readFileSync(mdPath, "utf8");
    fs.writeFileSync(htmlPath, buildDocHtml(title, markdown), "utf8");
    console.log(`Wrote ${path.relative(ROOT, htmlPath)}`);
  }
}

if (require.main === module) {
  main();
}

module.exports = {
  buildUserGuideHtml,
  buildDocHtml,
  DOC_PAGES,
  main,
  MD_PATH: path.join(DOCS_DIR, "user-guide.md"),
  HTML_PATH: path.join(DOCS_DIR, "user-guide.html"),
};
