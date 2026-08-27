using System;

namespace WordBarcodeStudio.Services;

public class WordNotAvailableException : Exception
{
    public WordNotAvailableException(string message) : base(message) { }
    public WordNotAvailableException(string message, Exception inner) : base(message, inner) { }
}

public class WordAutomationException : Exception
{
    public WordAutomationException(string message) : base(message) { }
    public WordAutomationException(string message, Exception inner) : base(message, inner) { }
}

public class BarcodeResult
{
    public string FieldCode { get; set; } = string.Empty;
    public string DocxPath { get; set; } = string.Empty;
}
