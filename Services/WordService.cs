using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
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

        try
        {
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
    }

    /// <summary>
    /// Pulls the rendered barcode out of Word and saves it as a PNG so the UI
    /// can show a real preview. Must be called on an STA thread (the UI thread)
    /// because it may read the clipboard.
    /// </summary>
    public string? ExtractPreview()
    {
        if (_doc == null || _wordApp == null) return null;

        // Method 1: the barcode usually renders as an inline picture.
        try
        {
            dynamic inline = _doc!.InlineShapes;
            if (inline.Count >= 1)
            {
                dynamic shape = inline[inline.Count];
                string p = Path.Combine(TempDirectory, $"preview_{Guid.NewGuid():N}.png");
                shape.SaveAsPicture(p);
                if (File.Exists(p) && new FileInfo(p).Length > 0) return p;
            }
        }
        catch { }

        // Method 2: copy the field result as a picture and read it from the clipboard.
        try
        {
            dynamic fields = _doc!.Fields;
            if (fields.Count >= 1)
            {
                dynamic field = fields[fields.Count];
                field.Select();
                dynamic sel = _wordApp!.Selection;
                sel.CopyAsPicture();
                return ReadClipboardPng();
            }
        }
        catch { }

        return null;
    }

    private string? ReadClipboardPng()
    {
        string p = Path.Combine(TempDirectory, $"preview_{Guid.NewGuid():N}.png");
        try
        {
            var data = Clipboard.GetDataObject();
            if (data == null) return null;

            if (data.GetDataPresent(DataFormats.EnhancedMetafile))
            {
                if (data.GetData(DataFormats.EnhancedMetafile) is Metafile mf)
                {
                    mf.Save(p, ImageFormat.Png);
                    return p;
                }
            }

            if (data.GetDataPresent(DataFormats.Bitmap))
            {
                if (data.GetData(DataFormats.Bitmap) is Bitmap bmp)
                {
                    bmp.Save(p, ImageFormat.Png);
                    return p;
                }
            }
        }
        catch { }

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
