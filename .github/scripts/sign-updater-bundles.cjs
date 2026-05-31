const fs = require("node:fs");
const path = require("node:path");
const cp = require("node:child_process");

const extension = process.argv[2];
if (!extension) {
  throw new Error("Usage: node .github/scripts/sign-updater-bundles.cjs <extension>");
}

const privateKey = process.env.TAURI_SIGNING_PRIVATE_KEY;
const password = process.env.TAURI_SIGNING_PRIVATE_KEY_PASSWORD || "";
const requireSignatures = process.env.REQUIRE_UPDATER_SIGNATURES === "true";

if (!privateKey) {
  const message = "TAURI_SIGNING_PRIVATE_KEY is not set; updater bundle signatures cannot be generated.";
  if (requireSignatures) {
    throw new Error(`${message} Add this GitHub secret before publishing a tag release.`);
  }
  console.log(`${message} Skipping updater signing.`);
  process.exit(0);
}

function walk(dir) {
  if (!fs.existsSync(dir)) {
    return [];
  }

  return fs.readdirSync(dir, { withFileTypes: true }).flatMap((entry) => {
    const fullPath = path.join(dir, entry.name);
    return entry.isDirectory() ? walk(fullPath) : [fullPath];
  });
}

function parseSignature(output) {
  const lines = output
    .split(/\r?\n/u)
    .map((line) => line.trim())
    .filter(Boolean);
  const explicit = lines.findLast((line) => /^signature:/iu.test(line));
  if (explicit) {
    return explicit.replace(/^signature:\s*/iu, "").trim();
  }
  const base64 = lines.findLast((line) => /^[A-Za-z0-9+/=]+$/u.test(line));
  return base64 || lines.at(-1) || "";
}

const bundles = walk("src-tauri/target")
  .filter((file) => file.endsWith(extension))
  .filter((file) => file.includes(`${path.sep}bundle${path.sep}`));

if (bundles.length === 0) {
  throw new Error(`No ${extension} bundles found under src-tauri/target/**/bundle.`);
}

const npx = process.platform === "win32" ? "npx.cmd" : "npx";
for (const bundle of bundles) {
  const output = cp.execFileSync(
    npx,
    ["--no-install", "tauri", "signer", "sign", "-k", privateKey, "-p", password, bundle],
    { encoding: "utf8", stdio: ["ignore", "pipe", "inherit"] },
  );
  const signature = parseSignature(output);
  if (!signature) {
    throw new Error(`Tauri signer did not return a signature for ${bundle}`);
  }
  fs.writeFileSync(`${bundle}.sig`, `${signature}\n`);
  console.log(`Signed ${bundle}`);
}
