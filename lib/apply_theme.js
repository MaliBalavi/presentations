// Vendored replacement for the pptx skill's apply_theme.js.
// The deck used to require an absolute path on one machine; this version
// lives in the repo and only needs python3 (stdlib only).
const path = require("path");
const { execFileSync } = require("child_process");

async function applyTheme(pptxPath, theme) {
  execFileSync(
    "python3",
    [path.join(__dirname, "apply_theme.py"), pptxPath, JSON.stringify(theme)],
    { stdio: "inherit" },
  );
}

module.exports = { applyTheme };
