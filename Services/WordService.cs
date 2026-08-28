using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace WordBarcodeStudio.Services;

/// <summary>
/// Controls Microsoft Word through COM automation and uses Word's native
/// DISPLAYBARCODE field as the barcode rendering engine.
/// Late binding (dynamic + ProgID) is used so the project compiles even when
/// the Word Primary Interop Assemblies are not installed.
/// </summary>
public class WordService : IDisposable
{
    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool OpenClipboard(IntPtr hWndNewOwner);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool CloseClipboard();

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr GetClipboardData(uint uFormat);

    private const uint CF_BITMAP = 2;
    private const uint CF_ENHMETAFILE = 14;

    private dynamic? _wordApp;
    private dynamic? _doc;
    private string? _currentDocxPath;
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
                "Microsoft Word could not be found.\n\nThis application requires the desktop version of Microsoft Word.");
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
        // hide it again afterwards.
        bool restoreHidden = runInBackground;

        try
        {
            if (restoreHidden) _wordApp!.Visible = true;

            CloseCurrentDoc();

            _doc = _wordApp!.Documents.Add();
            dynamic range = _doc.Content;
            dynamic fields = _doc.Fields;

            // Create an empty field, then set its code.
            dynamic field = fields.Add(range);
            field.Code.Text = fieldCode;
            field.Update();

            try
            {
                _doc.ActiveWindow.View.ShowFieldCodes = false;
            }
            catch { }

            string docxPath = Path.Combine(TempDirectory, $"Barcode_{Guid.NewGuid():N}.docx");
            _doc.SaveAs2(docxPath, 16); // wdFormatDocumentDefault (.docx)
            _currentDocxPath = docxPath;

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

    public BarcodeResult GenerateMultipleBarcodes(IEnumerable<(string FieldCode, string Label)> items, bool runInBackground)
    {
        EnsureWord(runInBackground);
        bool restoreHidden = runInBackground;

        try
        {
            if (restoreHidden) _wordApp!.Visible = true;

            CloseCurrentDoc();

            _doc = _wordApp!.Documents.Add();
            dynamic fields = _doc.Fields;

            var itemList = items.ToList();
            for (int i = 0; i < itemList.Count; i++)
            {
                var item = itemList[i];
                dynamic endRange = _doc.Content;
                endRange.Collapse(0); // wdCollapseEnd = 0

                if (i > 0)
                {
                    endRange.InsertParagraphAfter();
                    endRange.Collapse(0);
                }

                if (!string.IsNullOrWhiteSpace(item.Label))
                {
                    endRange.InsertAfter(item.Label + "\n");
                    endRange.Collapse(0);
                }

                dynamic field = fields.Add(endRange);
                field.Code.Text = item.FieldCode;
                field.Update();
            }

            try
            {
                _doc.ActiveWindow.View.ShowFieldCodes = false;
            }
            catch { }

            string docxPath = Path.Combine(TempDirectory, $"MultiBarcode_{Guid.NewGuid():N}.docx");
            _doc.SaveAs2(docxPath, 16); // wdFormatDocumentDefault (.docx)
            _currentDocxPath = docxPath;

            return new BarcodeResult
            {
                FieldCode = string.Join("\n", itemList.Select(it => it.FieldCode)),
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
                "Multi-barcode generation failed.\nCheck data and barcode types, then try again.", ex);
        }
        finally
        {
            if (restoreHidden) { try { _wordApp!.Visible = false; } catch { } }
        }
    }

    /// <summary>
    /// Pulls the rendered barcode out of Word and saves it as a PNG so the UI
    /// can show a real preview. Must be called on an STA thread (the UI thread).
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
            await Task.Delay(300); // give Word time to render the field graphic

            // Attempt 1: Copy the field selection as picture
            try
            {
                dynamic fields = _doc.Fields;
                if (fields.Count >= 1)
                {
                    dynamic field = fields[1];
                    field.Select();
                    _wordApp.Selection.CopyAsPicture();
                    result = ReadClipboardPng();
                }
            }
            catch { }

            // Attempt 2: Copy whole document content
            if (result == null)
            {
                try
                {
                    _doc.Content.Select();
                    _wordApp.Selection.CopyAsPicture();
                    result = ReadClipboardPng();
                }
                catch { }
            }

            // Attempt 3: Standard copy
            if (result == null)
            {
                try
                {
                    _doc.Content.Select();
                    _wordApp.Selection.Copy();
                    result = ReadClipboardPng();
                }
                catch { }
            }
        }
        finally
        {
            if (wasHidden) { try { _wordApp!.Visible = false; } catch { } }
        }
#pragma warning restore CS8602

        return result;
    }

    private string NewPng() => Path.Combine(TempDirectory, $"preview_{Guid.NewGuid():N}.png");

    private static bool ValidImage(string p)
    {
        try
        {
            if (!File.Exists(p)) return false;
            var info = new FileInfo(p);
            return info.Length > 100;
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
            if (!OpenClipboard(IntPtr.Zero)) return null;

            try
            {
                IntPtr hEmf = GetClipboardData(CF_ENHMETAFILE);
                if (hEmf != IntPtr.Zero)
                {
                    using var mf = new Metafile(hEmf, false);
                    var header = mf.GetMetafileHeader();
                    int w = (int)header.Bounds.Width;
                    int h = (int)header.Bounds.Height;
                    if (w <= 0 || h <= 0)
                    {
                        w = Math.Max(250, mf.Width);
                        h = Math.Max(250, mf.Height);
                    }
                    if (w < 250 || h < 250)
                    {
                        float factor = Math.Max(350f / Math.Max(1, w), 350f / Math.Max(1, h));
                        w = Math.Max(100, (int)(w * factor));
                        h = Math.Max(100, (int)(h * factor));
                    }
                    using var bmp = new Bitmap(w, h);
                    using var g = Graphics.FromImage(bmp);
                    g.Clear(System.Drawing.Color.White);
                    g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                    g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                    g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                    g.DrawImage(mf, 0, 0, w, h);
                    bmp.Save(p, ImageFormat.Png);
                    if (ValidImage(p)) return p;
                }

                IntPtr hBmp = GetClipboardData(CF_BITMAP);
                if (hBmp != IntPtr.Zero)
                {
                    using var bmpSource = System.Drawing.Image.FromHbitmap(hBmp);
                    using var bmp = new Bitmap(bmpSource.Width, bmpSource.Height);
                    using var g = Graphics.FromImage(bmp);
                    g.Clear(System.Drawing.Color.White);
                    g.DrawImage(bmpSource, 0, 0, bmpSource.Width, bmpSource.Height);
                    bmp.Save(p, ImageFormat.Png);
                    if (ValidImage(p)) return p;
                }
            }
            finally
            {
                CloseClipboard();
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
            string tempPath = Path.Combine(TempDirectory, $"Open_{Guid.NewGuid():N}.docx");
            _doc.SaveCopyAs(tempPath);
            using var process = Process.Start(new ProcessStartInfo(tempPath) { UseShellExecute = true });
            if (process == null) ShowWordFallback();
        }
        catch
        {
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
