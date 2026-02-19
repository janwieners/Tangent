# Tangent

Converts `.docx` files into JPG images – one file per page.

## Requirements

### macOS (Homebrew)

```bash
brew install libreoffice poppler
```

> **Note:** On first run, macOS Gatekeeper may block LibreOffice. Open the app once manually via Finder and confirm the security dialog – after that it runs headlessly without issues.
>
> Apple Silicon (M1/M2/M3) is fully supported.

### Linux (apt)

```bash
sudo apt install libreoffice poppler-utils
```

### ImageMagick (optional, required for `--resize`)

```bash
# macOS
brew install imagemagick

# Linux
sudo apt install imagemagick
```

### Node.js & tsx (all platforms)

```bash
npm install -g tsx
```

## Usage

```bash
# Convert a single file
tsx convert.ts -i report.docx

# Convert a directory
tsx convert.ts -i ./documents -o ./images

# Recursive + high resolution
tsx convert.ts -r -i ./projects -o ./export -d 300 -q 95

# Resize to A4 at 300 DPI (2480×3508 px)
tsx convert.ts -i report.docx -s 2480x3508

# Help
tsx convert.ts --help
```

## Options

| Option              | Description                               | Default         |
|---------------------|-------------------------------------------|-----------------|
| `-i`, `--input`     | Input file or directory                   | Current dir     |
| `-o`, `--output`    | Output directory                          | `./output`      |
| `-d`, `--dpi`       | Resolution in DPI                         | `150`           |
| `-q`, `--quality`   | JPG quality (1–100)                       | `90`            |
| `-s`, `--resize`    | Resize to exact pixels, e.g. `2480x3508` | no resize       |
| `-r`, `--recursive` | Include subdirectories                    | `false`         |

## Output Structure

For each input file a subdirectory is created inside the output folder:

```
output/
├── Report_Q1/
│   ├── Report_Q1_page-1.jpg
│   ├── Report_Q1_page-2.jpg
│   └── Report_Q1_page-3.jpg
└── Minutes/
    ├── Minutes_page-1.jpg
    └── Minutes_page-2.jpg
```

## How It Works

1. **DOCX/DOTM → PDF** via LibreOffice (headless mode)
2. **PDF → JPG** via `pdftoppm` (poppler-utils / poppler)
