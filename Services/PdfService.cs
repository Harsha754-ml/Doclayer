using System;

namespace WordBarcodeStudio.Services;

public static class PdfService
{
    // Exports a Word document to PDF using Word's own ExportAsFixedFormat.
    // 17 = wdExportFormatPDF
    public static void ExportToPdf(dynamic document, string filePath)
    {
        object format = 17;
        object missing = Type.Missing;

        document.ExportAsFixedFormat(
            filePath,
            format,
            missing, missing, missing, missing, missing,
            missing, missing, missing, missing, missing, missing, missing, missing);
    }
}
