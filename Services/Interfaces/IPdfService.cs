namespace Services.Interfaces;

public interface IPdfService
{
    /// <summary>
    /// Word 轉換成 Pdf
    /// </summary>
    /// <param name="wordData">word數據</param>
    /// <returns></returns>
    Task<byte[]> ConvertWordToPdfAsync(byte[] wordData);
    /// <summary>
    /// PDF密碼
    /// </summary>
    /// <param name="pdfBytes">PDF數據</param>
    /// <param name="password">密碼</param>
    /// <returns></returns>
    Task<byte[]> GeneratePdfWithPassword(byte[] pdfBytes, string password);
    /// <summary>
    /// 合併PDF
    /// </summary>
    /// <param name="pdfBytesList">PDF數據列表</param>
    /// <returns></returns>
    Task<byte[]> MergePdf(List<byte[]> pdfBytesList);
    /// <summary>
    /// PDF新增頁碼
    /// </summary>
    /// <param name="pdfBytes">PDF數據</param>
    /// <returns></returns>
    byte[] AddPageNumber(byte[] pdfBytes);

}
