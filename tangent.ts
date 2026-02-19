#!/usr/bin/env node
/**
 * Tangent – Word to JPG Converter
 * Converts .docx and .dotm files into JPG images (one JPG per page)
 *
 * Requirements:
 *   - LibreOffice (for DOCX → PDF)
 *   - pdftoppm / poppler (for PDF → JPG)
 *   - imagemagick for resizing
 *
 * Usage:
 *   npx tsx tangent.ts [options] [path]
 *
 * Options:
 *   --input,    -i <path>   Input: file or directory (default: current directory)
 *   --output,   -o <path>   Output directory (default: ./output)
 *   --dpi,      -d <num>    Resolution in DPI (default: 150)
 *   --quality,  -q <num>    JPG quality 1–100 (default: 90)
 *   --resize,   -s <WxH>    Resize output to WxH pixels, e.g. 2480x3508
 *                           Use 2480x3508 for A4 at 300 DPI (default: no resize)
 *   --recursive, -r         Search directories recursively
 *   --help,     -h          Show help
 *
 * Examples:
 *   npx tsx tangent.ts -i ./documents -o ./images -d 200 -q 95
 *   npx tsx tangent.ts -i report.docx -s 2480x3508
 *   npx tsx tangent.ts -r -i ./projects
 */

import { execSync, spawnSync } from "child_process";
import * as fs from "fs";
import * as path from "path";
import * as os from "os";

// ─── Types ───────────────────────────────────────────────────────────────────

interface Config {
  inputPath: string;
  outputDir: string;
  dpi: number;
  quality: number;
  resize: { width: number; height: number } | null;
  recursive: boolean;
}

interface ConversionResult {
  source: string;
  pages: number;
  outputFiles: string[];
  error?: string;
}

// ─── Helper Functions ─────────────────────────────────────────────────────────

function printHelp(): void {
  console.log(`
╔══════════════════════════════════════════════════════╗
║               Tangent – Word to JPG                  ║
╚══════════════════════════════════════════════════════╝

Converts .docx and .dotm files into JPG images.
A separate JPG file is created for each page.

USAGE:
  npx tsx tangent.ts [options]

OPTIONS:
  -i, --input <path>      Input: file or directory
                          (default: current directory)
  -o, --output <path>     Output directory
                          (default: ./output)
  -d, --dpi <num>         Resolution in DPI (default: 150)
  -q, --quality <num>     JPG quality 1–100 (default: 90)
  -s, --resize <WxH>      Resize to exact pixel dimensions, e.g. 2480x3508
                          (A4 at 300 DPI = 2480x3508, no resize by default)
  -r, --recursive         Include subdirectories
  -h, --help              Show this help

EXAMPLES:
  npx tsx tangent.ts -i ./documents
  npx tsx tangent.ts -i report.docx -o ./images -d 200
  npx tsx tangent.ts -i report.docx -s 2480x3508
  npx tsx tangent.ts -r -i ./projects -q 95
`);
}

/** Resolves the LibreOffice binary path across platforms. */
function resolveLibreOffice(): string {
  // macOS app bundle locations (Homebrew cask or manual install)
  const macPaths = [
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    `${os.homedir()}/Applications/LibreOffice.app/Contents/MacOS/soffice`,
  ];
  for (const p of macPaths) {
    if (fs.existsSync(p)) return p;
  }
  // PATH-based lookup (Linux, or macOS with a working symlink)
  const which = spawnSync("which", ["libreoffice"], { encoding: "utf8" });
  if (which.status === 0 && which.stdout.trim()) return which.stdout.trim();

  const whichSoffice = spawnSync("which", ["soffice"], { encoding: "utf8" });
  if (whichSoffice.status === 0 && whichSoffice.stdout.trim())
    return whichSoffice.stdout.trim();

  return ""; // not found
}

function checkDependencies(needsResize: boolean): void {
  // LibreOffice
  const libreOfficeBin = resolveLibreOffice();
  if (!libreOfficeBin) {
    console.error(`❌ Error: LibreOffice is not installed.`);
    console.error(`   macOS  → brew install libreoffice`);
    console.error(`   Linux  → sudo apt install libreoffice`);
    process.exit(1);
  }

  // pdftoppm
  const pdftoppm = spawnSync("which", ["pdftoppm"], { encoding: "utf8" });
  if (pdftoppm.status !== 0) {
    console.error(`❌ Error: "pdftoppm" is not installed.`);
    console.error(`   macOS  → brew install poppler`);
    console.error(`   Linux  → sudo apt install poppler-utils`);
    process.exit(1);
  }

  // ImageMagick (only required when --resize is used)
  if (needsResize) {
    const convert = spawnSync("which", ["convert"], { encoding: "utf8" });
    if (convert.status !== 0) {
      console.error(`❌ Error: "convert" (ImageMagick) is not installed.`);
      console.error(`   Required for --resize.`);
      console.error(`   macOS  → brew install imagemagick`);
      console.error(`   Linux  → sudo apt install imagemagick`);
      process.exit(1);
    }
  }
}

function parseArgs(): Config {
  const args = process.argv.slice(2);
  const config: Config = {
    inputPath: process.cwd(),
    outputDir: path.join(process.cwd(), "output"),
    dpi: 150,
    quality: 90,
    resize: null,
    recursive: false,
  };

  for (let i = 0; i < args.length; i++) {
    switch (args[i]) {
      case "-h":
      case "--help":
        printHelp();
        process.exit(0);
        break;
      case "-i":
      case "--input":
        config.inputPath = path.resolve(args[++i]);
        break;
      case "-o":
      case "--output":
        config.outputDir = path.resolve(args[++i]);
        break;
      case "-d":
      case "--dpi":
        config.dpi = parseInt(args[++i], 10);
        break;
      case "-q":
      case "--quality":
        config.quality = parseInt(args[++i], 10);
        break;
      case "-s":
      case "--resize": {
        const raw = args[++i];
        const match = raw.match(/^(\d+)[x×](\d+)$/i);
        if (!match) {
          console.error(`❌ Invalid --resize value "${raw}". Expected format: WxH (e.g. 2480x3508)`);
          process.exit(1);
        }
        config.resize = { width: parseInt(match[1], 10), height: parseInt(match[2], 10) };
        break;
      }
      case "-r":
      case "--recursive":
        config.recursive = true;
        break;
      default:
        // treat positional argument as input path
        if (!args[i].startsWith("-")) {
          config.inputPath = path.resolve(args[i]);
        }
    }
  }

  return config;
}

/** Collects all Word files at the given path */
function collectWordFiles(inputPath: string, recursive: boolean): string[] {
  const stat = fs.statSync(inputPath);
  const extensions = [".docx", ".dotm"];

  if (stat.isFile()) {
    const ext = path.extname(inputPath).toLowerCase();
    if (extensions.includes(ext)) {
      return [inputPath];
    } else {
      console.error(`❌ File is not a Word document: ${inputPath}`);
      process.exit(1);
    }
  }

  const files: string[] = [];

  function scanDir(dir: string): void {
    const entries = fs.readdirSync(dir, { withFileTypes: true });
    for (const entry of entries) {
      const fullPath = path.join(dir, entry.name);
      if (entry.isDirectory() && recursive) {
        scanDir(fullPath);
      } else if (entry.isFile()) {
        const ext = path.extname(entry.name).toLowerCase();
        if (extensions.includes(ext)) {
          files.push(fullPath);
        }
      }
    }
  }

  scanDir(inputPath);
  return files;
}

/** Converts a single Word file to JPGs */
function convertFile(
  filePath: string,
  outputDir: string,
  dpi: number,
  quality: number,
  resize: { width: number; height: number } | null
): ConversionResult {
  const basename = path.basename(filePath, path.extname(filePath));
  const fileOutputDir = path.join(outputDir, basename);
  fs.mkdirSync(fileOutputDir, { recursive: true });

  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "word2jpg-"));

  try {
    // Step 1: DOCX → PDF via LibreOffice
    console.log(`  📄 Converting to PDF...`);
    const libreOfficeBin = resolveLibreOffice();
    const libreResult = spawnSync(
      libreOfficeBin,
      ["--headless", "--convert-to", "pdf", "--outdir", tmpDir, filePath],
      { encoding: "utf8" }
    );

    if (libreResult.status !== 0) {
      throw new Error(
        `LibreOffice error: ${libreResult.stderr || libreResult.stdout}`
      );
    }

    // Locate the PDF file
    const pdfFile = path.join(tmpDir, `${basename}.pdf`);
    if (!fs.existsSync(pdfFile)) {
      // Name may differ – search generically
      const pdfs = fs
        .readdirSync(tmpDir)
        .filter((f) => f.endsWith(".pdf"))
        .map((f) => path.join(tmpDir, f));
      if (pdfs.length === 0) {
        throw new Error("No PDF file found after LibreOffice conversion.");
      }
    }

    const actualPdf = fs
      .readdirSync(tmpDir)
      .filter((f) => f.endsWith(".pdf"))
      .map((f) => path.join(tmpDir, f))[0];

    // Step 2: Determine page count
    let pageCount = 1;
    try {
      const infoResult = spawnSync(
        "identify",
        ["-format", "%n\n", `${actualPdf}[0]`],
        { encoding: "utf8" }
      );
      // Page count via pdfinfo
      const pdfInfoResult = spawnSync("pdfinfo", [actualPdf], {
        encoding: "utf8",
      });
      if (pdfInfoResult.status === 0) {
        const match = pdfInfoResult.stdout.match(/Pages:\s+(\d+)/);
        if (match) pageCount = parseInt(match[1], 10);
      }
    } catch {
      // pdfinfo not available – pdftoppm handles it anyway
    }

    // Step 3: PDF → JPG via pdftoppm (poppler-utils / poppler)
    console.log(`  🖼️  Creating JPG images (${pageCount} page(s), ${dpi} DPI)...`);
    const outputPrefix = path.join(fileOutputDir, `${basename}_page`);

    const convertResult = spawnSync(
      "pdftoppm",
      [
        "-jpeg",
        "-r", String(dpi),
        "-jpegopt", `quality=${quality}`,
        actualPdf,
        outputPrefix,
      ],
      { encoding: "utf8" }
    );

    if (convertResult.status !== 0) {
      throw new Error(
        `pdftoppm error: ${convertResult.stderr || convertResult.stdout}`
      );
    }

    // Collect output files
    const outputFiles = fs
      .readdirSync(fileOutputDir)
      .filter((f) => f.endsWith(".jpg"))
      .sort()
      .map((f) => path.join(fileOutputDir, f));

    // Step 4 (optional): Resize via ImageMagick
    if (resize) {
      console.log(`  📐 Resizing to ${resize.width}×${resize.height} px...`);
      for (const jpgFile of outputFiles) {
        const resizeResult = spawnSync(
          "convert",
          [
            jpgFile,
            "-resize", `${resize.width}x${resize.height}!`,
            "-quality", String(quality),
            jpgFile,
          ],
          { encoding: "utf8" }
        );
        if (resizeResult.status !== 0) {
          throw new Error(
            `ImageMagick resize error: ${resizeResult.stderr || resizeResult.stdout}`
          );
        }
      }
    }

    return {
      source: filePath,
      pages: outputFiles.length,
      outputFiles,
    };
  } catch (err) {
    return {
      source: filePath,
      pages: 0,
      outputFiles: [],
      error: err instanceof Error ? err.message : String(err),
    };
  } finally {
    // Clean up temporary directory
    try {
      fs.rmSync(tmpDir, { recursive: true, force: true });
    } catch {
      // ignore
    }
  }
}

// ─── Main ───────────────────────────────────────────────────────────────────

async function main(): Promise<void> {
  console.log("\n🔄 Tangent – Word to JPG Converter\n");

  const config = parseArgs();

  checkDependencies(config.resize !== null);

  // Validate input path
  if (!fs.existsSync(config.inputPath)) {
    console.error(`❌ Input path not found: ${config.inputPath}`);
    process.exit(1);
  }

  // Collect Word files
  const files = collectWordFiles(config.inputPath, config.recursive);

  if (files.length === 0) {
    console.log("⚠️  No Word files (.docx, .dotm) found.");
    process.exit(0);
  }

  console.log(`📁 Output directory: ${config.outputDir}`);
  console.log(`📋 Files found:      ${files.length}`);
  console.log(`📐 Resolution:       ${config.dpi} DPI`);
  console.log(`🎨 JPG quality:      ${config.quality}%`);
  if (config.resize) {
    console.log(`↔️  Resize to:        ${config.resize.width}×${config.resize.height} px`);
  }
  console.log("");

  // Create output directory
  fs.mkdirSync(config.outputDir, { recursive: true });

  // Run conversions
  const results: ConversionResult[] = [];
  let successCount = 0;
  let errorCount = 0;

  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    const filename = path.basename(file);
    console.log(`[${i + 1}/${files.length}] ${filename}`);

    const result = convertFile(file, config.outputDir, config.dpi, config.quality, config.resize);
    results.push(result);

    if (result.error) {
      console.log(`  ❌ Error: ${result.error}\n`);
      errorCount++;
    } else {
      console.log(
        `  ✅ ${result.pages} page(s) → ${path.relative(process.cwd(), path.dirname(result.outputFiles[0]))}/\n`
      );
      successCount++;
    }
  }

  // Summary
  console.log("─".repeat(50));
  console.log(`✅ Successful: ${successCount} file(s)`);
  if (errorCount > 0) {
    console.log(`❌ Errors:     ${errorCount} file(s)`);
  }

  const totalJpgs = results.reduce((sum, r) => sum + r.pages, 0);
  console.log(`🖼️  Total JPGs: ${totalJpgs}`);
  console.log(`📂 Saved to:   ${config.outputDir}\n`);
}

main().catch((err) => {
  console.error("Unexpected error:", err);
  process.exit(1);
});
