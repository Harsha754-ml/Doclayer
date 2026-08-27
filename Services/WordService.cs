using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WordBarcodeStudio.Services;

/// <summary>
/// Controls Microsoft Word through COM automation and uses Word's native
/// DISPLAYBARCODE field as the barcode rendering engine.
/// Late binding (dynamic + ProgID) is used so the project compiles even when
/// the Word Primary Interop Assemblies are not installed.
/// </summary>
public class WordService : IDisposable
{
    private dynamic? _wordApp;
    private dynamic? _doc;
    private bool _disposed;

    public string TempDirectory { get; }

    public WordService()
    {
        TempDirectory = Path.Combine(Path.GetTempPath(), "WordBarcodeStudio");
        Directory.CreateDirectory(TempDirectory);
    }

    public bool IsWordRunning => _wordApp != null;

    public void EnsureWord(bool runInBackground)
    {
        if (_wordApp != null) return;

        var wordType = Type.GetTypeFromProgID("Word.Application");
        if (wordType == null)
        {
            throw new WordNotAvailableException(
                "Microsoft Word could not be found.\n\nThis prototype requires the desktop version of Microsoft Word.");
        }

        try
        {
            _wordApp = Activator.CreateInstance(wordType);
        }
        catch (Exception ex)
        {
            throw new WordAutomationException(
                "Unable to communicate with Microsoft Word.\nPlease close any stuck Word processes and try again.", ex);
        }

        _wordApp!.Visible = !runInBackground;
    }

    public BarcodeResult GenerateBarcode(string fieldCode, bool runInBackground)
    {
        EnsureWord(runInBackground);

        // Word only materializes the DISPLAYBARCODE graphic while it is visible, so
        // we must show it during field creation/update even in background mode, then
        // hide it again afterwards. Otherwise the rendered shape comes out blank.
        bool restoreHidden = runInBackground;

        try
        {
            if (restoreHidden) _wordApp!.Visible = true;

            CloseCurrentDoc();

            _doc = _wordApp!.Documents.Add();
            dynamic range = _doc.Content;
            dynamic fields = _doc.Fields;

            // Create an empty field, then set its code. This avoids passing the
            // code as the 3rd argument to Fields.Add, which is unreliable via
            // late-bound COM.
            dynamic field = fields.Add(range);
            field.Code.Text = fieldCode;
            field.Update();

            string docxPath = Path.Combine(TempDirectory, $"Barcode_{Guid.NewGuid():N}.docx");
            _doc.SaveAs2(docxPath, 16); // wdFormatDocumentDefault (.docx)

            return new BarcodeResult
            {
                FieldCode = fieldCode,
                DocxPath = docxPath
            };
        }
        catch (WordNotAvailableException)
        {
            throw;
        }
        catch (Exception ex)
        {
            throw new WordAutomationException(
                "Field generation failed.\nCheck the data and barcode type, then try again.", ex);
        }
        finally
        {
            if (restoreHidden) { try { _wordApp!.Visible = false; } catch { } }
        }
    }

    /// <summary>
    /// Pulls the rendered barcode out of Word and saves it as a PNG so the UI
    /// can show a real preview. Must be called on an STA thread (the UI thread)
    /// because it may read the clipboard.
    /// </summary>
    /// <summary>
    /// Pulls the rendered barcode out of Word and saves it as a PNG so the UI can
    /// show a real preview. Must be called on an STA thread (the UI thread).
    /// Word only materializes the barcode graphic when it is visible, so when the
    /// app is running Word in the background we briefly show Word, let it render,
    /// then hide it again.
    /// </summary>
    public async Task<string?> ExtractPreviewAsync()
    {
        if (_doc == null || _wordApp == null) return null;

#pragma warning disable CS8602
        bool wasHidden = false;
        try { wasHidden = !_wordApp.Visible; } catch { }
        if (wasHidden) { try { _wordApp.Visible = true; } catch { } }

        string? result = null;
        try
        {
            await Task.Delay(200); // give Word a moment to render the field result
            result = TryExtract();
        }
        finally
        {
            if (wasHidden) { try { _wordApp!.Visible = false; } catch { } }
        }
#pragma warning restore CS8602

        return result;
    }

    private string? TryExtract()
    {
        if (_doc == null) return null;

        // 1) Most reliable: let Word export the document to filtered HTML, which
        //    rasterizes the barcode as a real image file under a .files folder.
        var html = ExtractPreviewViaHtml();
        if (html != null) return html;

        // 2) Inline picture -> SaveAsPicture.
        try
        {
            dynamic inline = _doc.InlineShapes;
            if (inline.Count >= 1)
            {
                dynamic shape = inline[inline.Count];
                string p = NewPng();
                shape.SaveAsPicture(p);
                if (ValidImage(p)) return p;
            }
        }
        catch { }

        // 3) Floating shape -> CopyPicture -> clipboard.
        try
        {
            dynamic shapes = _doc.Shapes;
            if (shapes.Count >= 1)
            {
                dynamic shape = shapes[shapes.Count];
                shape.CopyPicture();
                var r = ReadClipboardPng();
                if (r != null) return r;
            }
        }
        catch { }

        // 4) Field -> select -> copy as picture -> clipboard (type-agnostic).
        try
        {
            dynamic fields = _doc.Fields;
            if (fields.Count >= 1)
            {
                dynamic field = fields[fields.Count];
                field.Select();
                dynamic sel = _wordApp!.Selection;
                sel.CopyAsPicture();
                var r = ReadClipboardPng();
                if (r != null) return r;
            }
        }
        catch { }

        return null;
    }

    /// <summary>
    /// Saves a throwaway copy of the document as filtered HTML. Word converts the
    /// DISPLAYBARCODE result into a raster image inside the companion .files folder,
    /// which we then copy out and use as the preview.
    /// </summary>
    private string? ExtractPreviewViaHtml()
    {
        try
        {
            string htmlPath = Path.Combine(TempDirectory, $"pv_{Guid.NewGuid():N}.htm");
            _doc!.SaveCopyAs(htmlPath, 10); // wdFormatFilteredHTML

            string? filesDir = Directory
                .GetDirectories(TempDirectory, "*.files")
                .OrderByDescending(d => Directory.GetLastWriteTimeUtc(d))
                .FirstOrDefault();

            if (filesDir == null || !Directory.Exists(filesDir)) return null;

            var img = Directory
                .GetFiles(filesDir, "*.*")
                .Where(f => new[] { ".png", ".gif", ".jpg", ".jpeg" }
                    .Any(e => f.EndsWith(e, StringComparison.OrdinalIgnoreCase)))
                .OrderByDescending(f => new FileInfo(f).Length)
                .FirstOrDefault();

            if (img == null) return null;

            string ext = Path.GetExtension(img);
            string outPath = Path.Combine(TempDirectory, $"preview_{Guid.NewGuid():N}{ext}");
            File.Copy(img, outPath);

            // Clean up the temporary HTML export.
            try { if (File.Exists(htmlPath)) File.Delete(htmlPath); } catch { }
            try { if (filesDir != null) Directory.Delete(filesDir, true); } catch { }

            return ValidImage(outPath) ? outPath : null;
        }
        catch
        {
            return null;
        }
    }

    private string NewPng() => Path.Combine(TempDirectory, $"preview_{Guid.NewGuid():N}.png");

    private static bool ValidImage(string p)
    {
        try
        {
            if (!File.Exists(p)) return false;
            var info = new FileInfo(p);
            // A real barcode image is well over a few hundred bytes.
            return info.Length > 200;
        }
        catch
        {
            return false;
        }
    }

    private string? ReadClipboardPng()
    {
        string p = NewPng();
        try
        {
            if (!Clipboard.ContainsData(DataFormats.Bitmap) &&
                !Clipboard.ContainsData(DataFormats.EnhancedMetafile))
            {
                return null;
            }

            // Direct bitmap is the easy case.
            if (Clipboard.ContainsData(DataFormats.Bitmap) &&
                Clipboard.GetData(DataFormats.Bitmap) is Bitmap bmp)
            {
                bmp.Save(p, ImageFormat.Png);
                if (ValidImage(p)) return p;
            }

            // Enhanced metafile: render onto a bitmap so PNG encoding is reliable.
            if (Clipboard.ContainsData(DataFormats.EnhancedMetafile) &&
                Clipboard.GetData(DataFormats.EnhancedMetafile) is Metafile mf)
            {
                int w = Math.Max(32, Math.Min(mf.Width, 4096));
                int h = Math.Max(32, Math.Min(mf.Height, 4096));
                using var outBmp = new Bitmap(w, h);
                using var g = Graphics.FromImage(outBmp);
                g.FillRectangle(Brushes.White, 0, 0, w, h);
                g.DrawImage(mf, 0, 0, w, h);
                outBmp.Save(p, ImageFormat.Png);
                if (ValidImage(p)) return p;
            }
        }
        catch
        {
            try { if (File.Exists(p)) File.Delete(p); } catch { }
        }

        try { if (File.Exists(p)) File.Delete(p); } catch { }
        return null;
    }

    public void SaveDocx(string destinationPath)
    {
        if (_doc == null) throw new WordAutomationException("Generate a barcode first.");

        try
        {
            _doc.SaveCopyAs(destinationPath);
        }
        catch (Exception ex)
        {
            throw new WordAutomationException(
                "The document could not be saved.\nCheck that the destination is writable.", ex);
        }
    }

    public void ExportPdf(string pdfPath)
    {
        if (_doc == null) throw new WordAutomationException("Generate a barcode first.");

        try
        {
            PdfService.ExportToPdf(_doc, pdfPath);
        }
        catch (Exception ex)
        {
            throw new WordAutomationException(
                "PDF export failed.\nCheck that the destination is writable.", ex);
        }
    }

    public void OpenInWord()
    {
        if (_doc == null) throw new WordAutomationException("Generate a barcode first.");

        try
        {
            // Save a throwaway copy and open it in the user's normal Word.
            string tempPath = Path.Combine(TempDirectory, $"Open_{Guid.NewGuid():N}.docx");
            _doc.SaveCopyAs(tempPath);
            using var process = Process.Start(new ProcessStartInfo(tempPath) { UseShellExecute = true });
            if (process == null) ShowWordFallback();
        }
        catch
        {
            // If shell open fails, reveal our Word instance with the document.
            ShowWordFallback();
        }
    }

    private void ShowWordFallback()
    {
        try { if (_wordApp != null) _wordApp.Visible = true; } catch { }
    }

    private void CloseCurrentDoc()
    {
        if (_doc != null)
        {
            try { _doc.Close(0); } // wdDoNotSaveChanges
            catch { }
            Release(_doc);
            _doc = null;
        }
    }

    private static void Release(object? o)
    {
        if (o != null)
        {
            try { Marshal.ReleaseComObject(o); } catch { }
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        try { CloseCurrentDoc(); } catch { }

        if (_wordApp != null)
        {
            try { _wordApp.Quit(0); } catch { }
            Release(_wordApp);
            _wordApp = null;
        }

        GC.Collect();
        GC.WaitForPendingFinalizers();
    }
}
