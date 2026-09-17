import "dotenv/config";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { mkdir, writeFile, readFile } from "node:fs/promises";
import XLSX from "xlsx";

import { runPipeline } from "./lib/apiClient.js";
import { sampleRows } from "./data/sampleRows.js";

// --- raw ANSI helpers (no extra deps) ---------------------------------
const c = {
  reset: "\x1b[0m",
  bold: "\x1b[1m",
  dim: "\x1b[2m",
  red: "\x1b[31m",
  green: "\x1b[32m",
  yellow: "\x1b[33m",
  blue: "\x1b[34m",
  magenta: "\x1b[35m",
  cyan: "\x1b[36m",
  gray: "\x1b[90m",
};

const paint = (color, text) => `${color}${text}${c.reset}`;
const rule = (char = "-", len = 60) => paint(c.gray, char.repeat(len));
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

function section(title) {
  console.log("\n" + rule("="));
  console.log(paint(c.bold + c.magenta, ` ${title}`));
  console.log(rule("="));
}

// --- setup ---------------------------------------------------------------
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const outputDir = path.join(__dirname, "..", "output");

// Fixed default input location - drop a CSV here and every run (no --file
// needed) picks it up automatically. Falls back to the hardcoded demo
// data in sampleRows.js only if this file hasn't been placed yet.
const DEFAULT_CSV_PATH = path.join(__dirname, "data", "sampleRows.csv");

const args = process.argv.slice(2);
const mock = args.includes("--mock") || process.env.MOCK_MODE === "true";
const fileArgIndex = args.indexOf("--file");
const filePath = fileArgIndex !== -1 ? args[fileArgIndex + 1] : null;

async function readRowsFromFile(targetPath) {
  const ext = path.extname(targetPath).toLowerCase();
  if (ext === ".json") {
    const raw = await readFile(targetPath, "utf8");
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) {
      throw new Error(`${targetPath} must contain a JSON array of product rows`);
    }
    return parsed;
  }

  // .csv / .xlsx - read as a buffer and pass to XLSX.read(), not
  // XLSX.readFile(), which breaks under Node ESM in SheetJS's .mjs build.
  const buffer = await readFile(targetPath);
  const wb = XLSX.read(buffer, { type: "buffer" });
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" });
  if (rows.length === 0) throw new Error(`${targetPath} appears empty`);
  return rows;
}

async function loadRows() {
  if (filePath) return { rows: await readRowsFromFile(filePath), source: filePath };

  try {
    return { rows: await readRowsFromFile(DEFAULT_CSV_PATH), source: "src/data/sampleRows.csv" };
  } catch (err) {
    if (err.code !== "ENOENT") throw err;
    return { rows: sampleRows, source: "src/data/sampleRows.js (built-in demo data)" };
  }
}

async function main() {
  console.log(paint(c.bold + c.cyan, "\nAutoproduct v2 POC"));
  console.log(
    paint(c.dim, "AI-powered product enrichment + Shopify CSV export pipeline demo")
  );
  console.log(
    paint(
      c.dim,
      `Mode: ${mock ? "MOCK (server returns canned AI output)" : "LIVE (server calls OpenAI)"}`
    )
  );
  console.log(paint(c.dim, `Engine: ${process.env.API_BASE_URL || "(API_BASE_URL not set)"}`));

  const { rows, source } = await loadRows();
  console.log(paint(c.dim, `Data source: ${source} (${rows.length} row(s))`));

  section("Requesting pipeline run");
  console.log(paint(c.yellow, "Calling engine API..."));
  const result = await runPipeline({ mode: mock ? "mock" : "live", rows });
  console.log(
    paint(
      c.yellow,
      `Engine responded: ${result.steps.length} products (mode=${result.meta.mode}, model=${result.meta.model})`
    )
  );

  section(`Walking through ${result.steps.length} products`);

  for (const step of result.steps) {
    const { row, enriched, pricing, elapsedMs, index, total } = step;
    const stepLabel = `[${index}/${total}]`;

    console.log("\n" + rule());
    console.log(paint(c.bold + c.blue, `${stepLabel} ${row.rawTitle || "(untitled)"}`));
    console.log(rule());

    console.log(paint(c.gray, "\nRaw supplier row:"));
    console.log(paint(c.gray, `  barcode:         ${row.barcode || "(empty)"}`));
    console.log(paint(c.gray, `  supplier:        ${row.supplierName || "(empty)"}`));
    console.log(paint(c.gray, `  brand:           ${row.brand || "(empty)"}`));
    console.log(paint(c.gray, `  raw title:       ${row.rawTitle || "(empty)"}`));
    console.log(paint(c.gray, `  raw description: ${row.rawDescription || "(empty)"}`));
    console.log(
      paint(c.gray, `  cost price:      ${row.costPrice != null ? "$" + Number(row.costPrice).toFixed(2) : "(empty)"}`)
    );
    console.log(paint(c.gray, `  qty on hand:     ${row.qtyOnHand ?? "(empty)"}`));

    await sleep(150);
    console.log(paint(c.yellow, `\nAI enrichment done (${elapsedMs}ms)`));

    console.log(paint(c.green, "\nEnriched product data:"));
    console.log(paint(c.green, `  title:            ${enriched.title}`));
    console.log(paint(c.green, `  product type:     ${enriched.productType}`));
    console.log(paint(c.green, `  collection:       ${enriched.subCategory}`));
    console.log(paint(c.green, `  vendor:           ${enriched.vendor}`));
    console.log(paint(c.green, `  seo title:        ${enriched.seoTitle}`));
    console.log(paint(c.green, `  seo description:  ${enriched.seoDescription}`));
    console.log(paint(c.green, `  tags:             ${enriched.tags.join(", ")}`));
    console.log(paint(c.green, `  body html:        ${enriched.bodyHtml.slice(0, 90)}...`));

    console.log(paint(c.cyan, "\nCalculated pricing:"));
    if (pricing.price === null) {
      console.log(paint(c.cyan, `  cost: (empty)  ->  no cost price supplied, price left blank`));
    } else {
      console.log(
        paint(
          c.cyan,
          `  cost: $${Number(row.costPrice).toFixed(2)}  ->  target margin: ${pricing.marginPercent}%  ->  price: $${pricing.price.toFixed(2)}`
        )
      );
    }

    await sleep(150);
  }

  // --- summary -------------------------------------------------------
  section("Summary");
  console.log(paint(c.bold, `${result.products.length} products enriched and priced:\n`));

  const col = (str, len) => String(str).padEnd(len).slice(0, len);
  console.log(
    paint(c.bold, col("TITLE", 42) + col("COST", 10) + col("PRICE", 10) + "MARGIN")
  );
  for (const p of result.products) {
    const costCell = p.costPrice != null ? `$${Number(p.costPrice).toFixed(2)}` : "-";
    const priceCell = p.price != null ? `$${Number(p.price).toFixed(2)}` : "-";
    console.log(
      col(p.title, 42) +
        col(costCell, 10) +
        col(priceCell, 10) +
        (p.marginPercent !== null ? `${p.marginPercent}%` : "-")
    );
  }

  // --- export ----------------------------------------------------------
  section("Exporting to CSV + JSON");
  await mkdir(outputDir, { recursive: true });

  const csvPath = path.join(outputDir, "products.csv");
  await writeFile(csvPath, result.csv, "utf8");

  const jsonPath = path.join(outputDir, "products.json");
  await writeFile(jsonPath, JSON.stringify(result.products, null, 2) + "\n", "utf8");

  console.log(paint(c.green, `\nExport complete:`));
  console.log(paint(c.green, `  CSV:  ${csvPath}`));
  console.log(paint(c.green, `  JSON: ${jsonPath}`));
  console.log(
    paint(
      c.dim,
      "\n(POC note: this writes local files only - no Shopify API connection is made.)\n"
    )
  );
}

main().catch((err) => {
  console.error(paint(c.red, `\nFatal error: ${err.message}`));
  process.exit(1);
});
