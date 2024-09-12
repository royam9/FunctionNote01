using System.Diagnostics;

namespace Services.Interfaces;

/// <summary>
/// Word 轉換成 PDF
/// </summary>
public interface IConvertWordToPdfService
{
    /// <summary>
    /// Word 轉換成 Pdf - 使用LibreOffice
    /// </summary>
    /// <param name="wordData">word數據</param>
    /// <returns></returns>
    Task<byte[]> ConvertWordToPdfAsync(byte[] wordData);
}
