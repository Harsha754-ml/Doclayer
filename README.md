# Doclayer

**Generate barcodes using Microsoft Word.**

A Windows desktop prototype that proves a standalone application can drive
**Microsoft Word** through COM automation to render barcodes/QR codes using
Word's native **`DISPLAYBARCODE`** field. No third-party QR/barcode library is
used for generation — **Word itself is the rendering engine**.

---

## What this is

This is a **prototype** built to demonstrate the architecture:

```
User
  ↓
Doclayer (WPF / .NET 8)
  ↓
Microsoft Word COM Automation
  ↓
Word Document + DISPLAYBARCODE field
  ↓
Rendered QR / Barcode
```

The application:

1. Accepts barcode data and options from the UI.
2. Starts Microsoft Word through COM (in the background by default).
3. Creates a temporary Word document.
4. Inserts a real `DISPLAYBARCODE` field programmatically (you never type the field code).
5. Forces Word to update the field so Word renders the barcode.
6. Extracts a preview image from Word (best-effort) and shows it in the app.
7. Lets you **Save DOCX**, **Export PDF**, and **Open in Word**.

---

## Requirements

- **Windows 10/11**
- **Microsoft Word desktop** (2013 or later, 32- or 64-bit) installed
- **.NET 8 SDK** (for building)
- Visual Studio 2022 (recommended) or `dotnet` CLI

> ⚠️ This prototype **requires the desktop version of Microsoft Word**.
> Word must remain installed for the app to use this architecture. There is
> no fallback QR generator — by design.

---

## How it works

- `BarcodeFieldService` converts the UI settings into the Word field code
  (escaping the data correctly). Example for the settings
  `https://example.com`, `QR`, EC `H`, scale `150`:

  ```
  { DISPLAYBARCODE "https://example.com" QR \q H \s 150 }
  ```

  The braces are added by Word; the application inserts the field via
  Word's `Fields.Add` API so Word recognizes it as a real field.

- `WordService` uses **late-bound COM** (`Type.GetTypeFromProgID("Word.Application")`
  + `Activator.CreateInstance`) so the project compiles even when the Word
  Primary Interop Assemblies are not installed on the build machine.

- `PdfService` calls Word's `ExportAsFixedFormat` to produce a PDF containing
  the rendered barcode.

---

## Build

Using Visual Studio:
1. Open `WordBarcodeStudio.csproj`.
2. Set the configuration to **Release** (or **Debug**).
3. Build → Build Solution (or `Ctrl+Shift+B`).

Using the .NET CLI:

```powershell
cd WordBarcodeStudio
dotnet build -c Release
```

The output is placed in `bin\Release\net8.0-windows\`.

---

## Run

From Visual Studio: press **F5** (or Ctrl+F5).

From the CLI:

```powershell
dotnet run -c Release
```

Make sure Microsoft Word is installed and not blocked by an interactive
dialog (e.g. a "first run" activation prompt) when you click **Generate**.

---

## Using the app

1. Enter the barcode data (default: `https://example.com`).
2. Select a **Barcode Type** (QR is default; QR and CODE128 are the focus).
3. For QR, set **Error Correction** (`L/M/Q/H`) and **Scale** (`%`).
4. (Optional) enable **Show encoded text** and set **Rotation**.
5. Click **GENERATE**.
   - Word is started in the background and renders the barcode.
   - A preview appears (best-effort, extracted from Word) and the
     generated field code is shown.
6. Click **Open in Word** to inspect the real document/field.
7. Click **Save DOCX** or **Export PDF** to write a file.

> **Run Word in background** (default on) hides Word. Turn it off to show
> Word for debugging/troubleshooting.

---

## Supported DISPLAYBARCODE switches

The options panel is **data-driven**: it shows only the switches that are valid
for the selected barcode type. All switches are Word's native `DISPLAYBARCODE`
switches (no third-party generation):

| Switch | Meaning | Applies to | UI control |
|--------|---------|------------|------------|
| `\s`   | Scale (10–1000%) | All | Number |
| `\r`   | Rotation (0–3 → 0/90/180/270°) | All | Dropdown |
| `\t`   | Show encoded text | All | Checkbox |
| `\q`   | QR error correction (L/M/Q/H) | QR | Dropdown |
| `\u`   | Unicode data | All | Checkbox |
| `\h`   | Height (twips) | All | Number |
| `\f`   | Foreground color | All | Dropdown |
| `\b`   | Background color | All | Dropdown |
| `\x`   | Fix invalid check digit | EAN/UPC | Checkbox |
| `\d`   | Add Start/Stop chars | CODE39 | Checkbox |
| `\c`   | ITF14 case code style (STD/2/3) | ITF14 | Dropdown |

The preview extracts the barcode image Word actually rendered, so every
combination of these switches is supported by the preview, Save DOCX, and
Export PDF paths.

---

## Project structure

```
WordBarcodeStudio/
│
├── Converters/
│   └── BooleanToVisibilityConverter.cs
├── Models/
│   └── BarcodeSettings.cs
├── Services/
│   ├── WordService.cs          # COM automation lifecycle
│   ├── BarcodeFieldService.cs  # settings -> DISPLAYBARCODE field code
│   ├── PdfService.cs           # Word -> PDF export
│   └── Exceptions.cs           # WordNotAvailable / WordAutomation errors
├── ViewModels/
│   └── MainViewModel.cs        # state, commands, error handling
├── Views/
│   ├── MainWindow.xaml
│   └── MainWindow.xaml.cs
├── App.xaml / App.xaml.cs
├── WordBarcodeStudio.csproj
└── README.md
```

---

## Temporary files

Generated files live in:

```
%TEMP%\WordBarcodeStudio\
```

The document is kept in memory while the app runs; **Save DOCX** / **Export PDF**
write copies to the location you choose. On exit, the hidden Word instance is
quit and COM objects are released so no orphaned `WINWORD.EXE` processes remain.

---

## Error handling

| Situation | Message |
|-----------|---------|
| Word not installed | `Microsoft Word could not be found. This prototype requires the desktop version of Microsoft Word.` |
| Word automation failure | `Unable to communicate with Microsoft Word. Please close any stuck Word processes and try again.` |
| Invalid barcode data | `The provided data is not valid for the selected barcode type.` |
| Empty input | `Enter barcode data first.` |
| Save failure | `The document could not be saved. Check that the destination is writable.` |
| PDF export failure | `PDF export failed. Check that the destination is writable.` |

Raw COM stack traces are never shown to the user.

---

## Test checklist

### Test 1 — QR (basic)
- Data: `https://example.com`, Type: `QR`
- Expected: a valid QR code is generated by Word.

### Test 2 — CODE128
- Data: `HELLO-12345`, Type: `CODE128`
- Expected: a valid CODE128 barcode is generated by Word.

### Test 3 — QR switches
- Error correction: `H`, Scale: `150`
- Expected: the generated field code contains `\q H` and `\s 150`.

### Test 4 — Word auto-start
- Close Word completely, then click **Generate**.
- Expected: the app launches Word automatically.

### Test 5 — No Word
- Run on a machine without Word installed.
- Expected: a friendly error is shown, the app does not crash.

### Test 6 — Persisted DOCX
- Generate, **Save DOCX**, close the app, reopen the DOCX in Word.
- Expected: the document still contains the generated barcode field/result.

### Test 7 — PDF export
- Click **Export PDF**.
- Expected: the PDF contains the generated barcode.

---

## Troubleshooting

- **"Microsoft Word could not be found."**
  Install the desktop version of Microsoft Word. The Microsoft Store / "Word
  Online" web app is not sufficient — COM automation needs the desktop app.

- **"Unable to communicate with Microsoft Word."**
  Open Task Manager and end any stuck `WINWORD.EXE` processes, then try again.

- **Generation succeeds but no preview image appears**
  The preview is best-effort (Word must expose the barcode as an inline
  picture so it can be saved). Use **Open in Word** or **Export PDF** to verify
  the barcode — the critical proof is that **Word generated it**, not the
  in-app preview.

- **"Field generation failed."**
  The data may be invalid for the selected symbology (e.g. EAN-13 needs
  12–13 digits). Check the data, or switch the barcode type.

- **COM / interop build errors**
  This project uses late-bound COM and does **not** require the Word PIA.
  Build on Windows with the .NET 8 SDK; do not add a `Microsoft.Office.Interop.Word`
  reference unless you specifically want early binding.

- **Leftover WINWORD.EXE processes**
  The app quits Word on exit via `try/finally` cleanup. If you break into the
  debugger and stop abruptly, a stray process may remain — end it in Task
  Manager.

---

## Out of scope (prototype constraints)

No user accounts, cloud sync, databases, analytics, payments, licensing,
batch processing, AI, multi-user support, or advanced template management.
This prototype only proves:

```
Input → App → Word COM → DISPLAYBARCODE → Valid Word doc → QR/barcode
```
