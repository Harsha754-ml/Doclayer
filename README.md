<div align="center">

# DocLayer

**A programmable document automation layer & native barcode studio for Microsoft Word.**

[![Release](https://img.shields.io/github/v/release/Harsha754-ml/Doclayer?color=10B981&label=Release)](https://github.com/Harsha754-ml/Doclayer/releases/latest)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%20%7C%2011-0A0A0A?logo=windows)](https://github.com/Harsha754-ml/Doclayer)
[![.NET](https://img.shields.io/badge/.NET-8.0--windows-512BD4?logo=dotnet)](https://dotnet.microsoft.com/)
[![License](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

</div>

---

## ⚡ Overview

**DocLayer** is a modern Windows desktop application that uses **Microsoft Word** as a programmable document and vector barcode rendering engine via COM automation.

Instead of relying on third-party bitmap generators, DocLayer injects native Word **`DISPLAYBARCODE`** fields directly into Word documents. This produces 100% vector-sharp, print-ready barcodes, multi-barcode document queues, high-DPI metafile previews, and instant Word (`.docx`) and PDF export pipelines.

```
┌──────────────┐     COM Automation     ┌────────────────────────┐     Native Rendering     ┌────────────────────────┐
│   DocLayer   │ ─────────────────────> │     Microsoft Word     │ ───────────────────────> │ Valid .docx / .pdf Doc │
│ Desktop App  │                        │ DISPLAYBARCODE Fields  │                          │ + Real-Time Preview    │
└──────────────┘                        └────────────────────────┘                          └────────────────────────┘
```

---

## 📥 Installation

### Option 1: Modern Setup Wizard (Recommended)

1. Go to the **[Latest GitHub Releases](https://github.com/Harsha754-ml/Doclayer/releases/latest)** page.
2. Download **`DocLayer-v1.0.0-Windows-Setup.zip`**.
3. Extract the ZIP archive and double-click **`DocLayer.Setup.exe`**.
4. Follow the setup wizard:
   - **EULA Agreement**: Review and accept the End-User License Agreement.
   - **Destination**: Choose your installation folder (default: `%LocalAppData%\Programs\DocLayer`).
   - **Shortcuts**: Select whether to create Desktop and Start Menu shortcuts.
   - **System Verification**: Setup automatically checks for .NET 8 Runtime and Microsoft Word COM registration.
   - **Finish**: Click Install, then launch DocLayer directly.

---

### Option 2: Build & Run from Source

#### Prerequisites
- **Windows 10 or 11 (64-bit)**
- **Microsoft Word desktop** (2013 or later) installed and activated
- **[.NET 8.0 SDK](https://dotnet.microsoft.com/download/dotnet/8.0)**

#### Build Steps
```powershell
# 1. Clone the repository
git clone https://github.com/Harsha754-ml/Doclayer.git
cd Doclayer

# 2. Build the application
dotnet build -c Release

# 3. Run DocLayer
dotnet run -c Release
```

#### Package the Installer Bundle
```powershell
# Build the complete installer and distribution zip
pwsh -File installer\build_installer.ps1
```

---

## 🚀 How to Use DocLayer (User Flow)

```
[1. Queue / Select Barcodes] ──> [2. Customize Data & Format] ──> [3. Click GENERATE] ──> [4. Preview & Export]
```

### Step 1: Manage Document Barcodes (Multi-Barcode Queue)
- Look at the **Document Barcodes** card on the left.
- Use **`+ Add Barcode`** to add additional barcode items to your document.
- Use the **Checkboxes** on each item to select or deselect which barcodes to include in the generated document.
- Click **`All`** or **`None`** for quick batch selection.
- Click **`✕`** to delete an item from the queue.

### Step 2: Configure Barcode Data & Options
- Click on any barcode item in the queue to select it for editing.
- **Item Header**: Customize the section title or label that appears above the barcode.
- **Barcode Type**: Select from 8 supported standards:
  - **QR**: Any URL, text, or payload (e.g. `https://example.com`)
  - **CODE128**: High-density alphanumeric ASCII (e.g. `DOC-2026-X89`)
  - **CODE39**: Industrial alphanumeric (e.g. `PART-9872`)
  - **EAN13**: 12 or 13 numeric retail digits (e.g. `5901234123457`)
  - **EAN8**: 7 or 8 numeric digits (e.g. `96385074`)
  - **UPCA**: 11 or 12 numeric product digits (e.g. `012345678905`)
  - **UPCE**: 6 to 8 numeric digits (e.g. `01234565`)
  - **ITF14**: 13 or 14 numeric carton digits (e.g. `10012345678902`)
- **Use Example**: Click the interactive **Use Example** button to immediately insert valid sample data for the chosen barcode type.
- **Options & Advanced Drawer**: Customize Scale (`%`), Error Correction (`L/M/Q/H`), Show Text, Rotation, Colors, and twip height.

### Step 3: Generate Document via Microsoft Word
- Click the primary **`GENERATE (X ITEMS)`** button.
- DocLayer connects to Word via COM, creates a document, inserts the native field codes, forces Word to render the barcodes, and extracts a high-DPI metafile preview onto the white paper sheet.

### Step 4: Preview, Inspect, and Export
- **Paper Canvas Preview**: Zoom and inspect the rendered barcode document in real-time.
- **Copy Field Code**: Copy the exact Word `DISPLAYBARCODE` syntax to your clipboard.
- **Open in Word**: Launch Microsoft Word to inspect and edit the live document with native fields.
- **Save DOCX**: Save the `.docx` document to your chosen folder.
- **Export PDF**: Generate a high-DPI PDF document with vector-sharp barcodes.

### Step 5: Session History & Templates
- Switch to the **History** tab in the sidebar to review past generations in the session and restore them with one click.
- Switch to **Templates** to browse document layouts.
- Switch to **Settings** to toggle Dark/Light mode and explore temporary cache files.

---

## 🗑️ Uninstallation

DocLayer integrates with Windows Programs & Features:
1. Open **Windows Settings** → **Apps** → **Installed Apps** (or Control Panel → Programs and Features).
2. Locate **DocLayer** and click **Uninstall**.
3. *Alternatively*, launch `Uninstall.exe` from your install directory or run:
   ```powershell
   DocLayer.Setup.exe --uninstall
   ```

---

## 🛠️ Supported DISPLAYBARCODE Switches

| Switch | Parameter | Description | Supported Types |
|:------:|:---------:|:------------|:----------------|
| `\s`   | `10-1000` | Scaling factor percentage | All |
| `\q`   | `L/M/Q/H` | QR Error Correction level | QR |
| `\r`   | `0-3`     | Rotation (0°, 90°, 180°, 270°) | All |
| `\t`   | *flag*    | Display human-readable text | All 1D / 2D |
| `\h`   | `twips`   | Height of barcode symbol | 1D Barcodes |
| `\f`   | `0xRRGGBB`| Foreground bar color | All |
| `\b`   | `0xRRGGBB`| Background sheet color | All |
| `\x`   | *flag*    | Fix invalid check digit | EAN13, UPCA |
| `\d`   | *flag*    | Start / Stop characters | CODE39 |
| `\c`   | `STD/2/3` | Case code packaging style | ITF14 |

---

## 📂 Project Architecture

```
WordBarcodeStudio/
├── Assets/                     # Application Icons (.ico, .png, .svg)
├── Converters/                 # WPF Theme & Binding Converters
├── Models/                     # BarcodeEntry, BarcodeOption, History Models
├── Services/
│   ├── WordService.cs          # Microsoft Word COM automation & EMF clipboard extraction
│   ├── BarcodeFieldService.cs  # DISPLAYBARCODE field compiler & format validator
│   ├── PdfService.cs           # Word FixedFormat PDF exporter
│   └── Exceptions.cs           # Diagnostic & COM exception handlers
├── ViewModels/                 # MainViewModel (Navigation, Queue, Commands)
├── Views/                      # MainWindow (Pitch-black shell, sidebar, inspector)
├── installer/
│   ├── DocLayer.Setup/         # Native WPF Setup Wizard & Uninstaller
│   ├── DocLayer.iss            # Inno Setup configuration
│   └── build_installer.ps1     # Automated release packager
├── App.xaml / App.xaml.cs      # Core styling, brushes, geometry icons
└── WordBarcodeStudio.csproj    # .NET 8 WPF project file
```

---

## 📄 License

This project is licensed under the [MIT License](LICENSE).
DocLayer is an independent product and is not affiliated with, sponsored by, or endorsed by Microsoft Corporation.
