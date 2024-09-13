using iTextSharp.text.pdf;
using System.Diagnostics;
using iTextSharp.text;
using Services.Interfaces;

namespace Services;

/// <summary>
/// PDF 相關
/// </summary>
public class PdfService : IPdfService
{
    /// <summary>
    /// 是否為 Linux環境
    /// </summary>
    private readonly bool _isLinux;
    /// <summary>
    /// LibreOffice執行檔位置
    /// </summary>
    private readonly string _libreOfficePath;

    public PdfService()
    {
        _isLinux = System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Linux);        
        _libreOfficePath = _isLinux ? "/lib/libreoffice/program/soffice" :
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "LibreOffice", "program", "soffice.exe");
    }

    /// <summary>
    /// Word 轉換成 Pdf
    /// </summary>
    /// <param name="wordData">word數據</param>
    /// <returns></returns>
    public async Task<byte[]> ConvertWordToPdfAsync(byte[] wordData)
    {
        var tempInputFile = Path.GetTempFileName();
        var tempOutputFile = Path.ChangeExtension(tempInputFile, ".pdf");

        try
        {
            await File.WriteAllBytesAsync(tempInputFile, wordData);

            using var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = _libreOfficePath,
                    Arguments = $"--headless --convert-to pdf \"{tempInputFile}\" --outdir \"{Path.GetDirectoryName(tempOutputFile)}\"",
                    UseShellExecute = false,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                }
            };

            process.Start();
            await process.WaitForExitAsync();

            if (process.ExitCode != 0)
            {
                var error = await process.StandardError.ReadToEndAsync();
                throw new Exception($"LibreOffice 轉換失敗。錯誤碼：{process.ExitCode}。錯誤信息：{error}");
            }

            return File.Exists(tempOutputFile) ? await File.ReadAllBytesAsync(tempOutputFile) :
                throw new Exception("無法找到生成的 PDF 文件。");
        }
        finally
        {
            File.Delete(tempInputFile);
            if (File.Exists(tempOutputFile)) File.Delete(tempOutputFile);
        }
    }

    /// <summary>
    /// PDF密碼
    /// </summary>
    /// <param name="pdfBytes">PDF數據</param>
    /// <param name="password">密碼</param>
    /// <returns></returns>
    public async Task<byte[]> GeneratePdfWithPassword(byte[] pdfBytes, string password)
    {
        // .NET 8 using
        // 將PDF數據放入Stream
        using var inputStream = new MemoryStream(pdfBytes);
        // 創一個Stream輸出
        using var outputStream = new MemoryStream();
        // PdfReader: 讀取與解析一個pdf Document
        using var reader = new PdfReader(inputStream);
        // PdfStamper: 開啟一個為已存在的PDF文件新增額外內容的程序
        using var stamper = new PdfStamper(reader, outputStream);

        stamper.SetEncryption(
            password != null ? System.Text.Encoding.UTF8.GetBytes(password) : null,
            password != null ? System.Text.Encoding.UTF8.GetBytes(password) : null,
            PdfWriter.ALLOW_PRINTING | PdfWriter.ALLOW_COPY,
            PdfWriter.ENCRYPTION_AES_128);

        stamper.Close();
        return outputStream.ToArray();
    }

    /// <summary>
    /// 合併PDF
    /// </summary>
    /// <param name="pdfBytesList">PDF數據列表</param>
    /// <returns></returns>
    /// Jill說因爲 pdf service 處理大量的 I/O，即使 func 内沒有使用 await 方法，register service 調用的時候如果以非同步方法呼叫會有比較好的 performance
    public async Task<byte[]> MergePdf(List<byte[]> pdfBytesList)
    {
        using var outputStream = new MemoryStream();
        using var document = new Document();
        // PdfCopy: 讀取一個reader或是Document的內容把它Copy到outputStream
        using var copy = new PdfCopy(document, outputStream);

        document.Open();

        foreach (var pdfBytes in pdfBytesList)
        {
            using var reader = new PdfReader(pdfBytes);
            for (int i = 1; i <= reader.NumberOfPages; i++)
            {
                // GetImportedPage: 從輸入的文件中抓取一個page
                copy.AddPage(copy.GetImportedPage(reader, i));
            }
            // 將reader寫進文件中並釋放記憶體
            copy.FreeReader(reader);
        }

        document.Close();
        return outputStream.ToArray();
    }

    /// <summary>
    /// PDF新增頁碼
    /// </summary>
    /// <param name="pdfBytes">PDF數據</param>
    /// <returns></returns>
    /// <remarks>https://todomato.blogspot.com/2019/07/pdf-pdf-itextsharp.html</remarks>
    public byte[] AddPageNumber(byte[] pdfBytes)
    {
        // 將數據寫入MemoryStream
        using var inputStream = new MemoryStream(pdfBytes);
        using var outputStream = new MemoryStream();
        using var reader = new PdfReader(inputStream);
        // 讀取一個文件，把修改的內容輸出到outputStream裡面
        using var stamper = new PdfStamper(reader, outputStream);

        int totalPages = reader.NumberOfPages;
        for (int i = 1; i <= totalPages; i++)
        {
            // GetOverContent() 取得第i頁的覆蓋層
            // PDF有三層
            // 背景層 背景等等
            // 內容層 內容等等
            // 覆蓋層 浮水印等等
            ColumnText.ShowTextAligned(stamper.GetOverContent(i),
                Element.ALIGN_CENTER, new Phrase($"{i}"),
                reader.GetPageSize(i).Width / 2, 30, 0);
        }

        stamper.Close();
        return outputStream.ToArray();
    }
}