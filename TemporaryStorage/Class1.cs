using System.Diagnostics;
using System.Text;

namespace TemporaryStorage;

public class Class1
{
    // 改成非同步之前的版本
    public byte[] ConvertUTF8HtmlToPdf(string htmlContent)
    {
        string wkhtmltopdfPath = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\..\Services\App\wkhtmltox\bin\wkhtmltopdf.exe"));

        // 確認 wkhtmltopdf.exe 是否存在
        if (!File.Exists(wkhtmltopdfPath))
        {
            throw new FileNotFoundException("The wkhtmltopdf executable was not found at the specified path.", wkhtmltopdfPath);
        }

        // 設置命令行參數，將 PDF 輸出到標準輸出流（stdout）
        string arguments = $"--page-size A4 --margin-top 20mm --margin-right 20mm --margin-bottom 20mm --margin-left 20mm --disable-smart-shrinking - -";

        // 創建進程啟動信息
        ProcessStartInfo processStartInfo = new ProcessStartInfo
        {
            FileName = wkhtmltopdfPath,
            Arguments = arguments,
            RedirectStandardInput = true,  // 重定向標準輸入
            RedirectStandardError = true,  // 重定向標準錯誤
            RedirectStandardOutput = true, // 重定向標準輸出
            UseShellExecute = false,
            CreateNoWindow = true
        };

        byte[] pdfBytes;

        using (Process process = Process.Start(processStartInfo))
        {
            // 確保使用 UTF-8 編碼寫入 HTML
            using (StreamWriter standardInput = new StreamWriter(process.StandardInput.BaseStream, Encoding.UTF8))
            {
                standardInput.Write(htmlContent);
            }

            // 使用 MemoryStream 來保存 PDF 數據
            using (MemoryStream memoryStream = new MemoryStream())
            {
                process.StandardOutput.BaseStream.CopyTo(memoryStream);
                pdfBytes = memoryStream.ToArray();
            }

            process.WaitForExit();
            if (process.ExitCode != 0)
            {
                string error = process.StandardError.ReadToEnd();
                throw new Exception($"wkhtmltopdf failed: {error}");
            }
        }

        return pdfBytes; // 返回生成的 PDF 數據
    }

    // Louis 之前提供的轉換程式碼
    /// <summary>
    /// Html 轉換成 PDF
    /// </summary>
    /// <returns></returns>
    public void ConvertHtmlToPdf(string htmlFilePath, string pdfOutputPath)
    {
        //var wkHtmlToPdfPath = Path.Combine(_hostingEnvironment.ContentRootPath, _wkHtmlToPdf);

        //string arguments = $"--encoding utf-8 --page-size A3 --margin-top 0 --margin-bottom 0 --margin-left 0 --margin-right 0 --disable-smart-shrinking --orientation Landscape \"{htmlFilePath}\" \"{pdfOutputPath}\"";

        //using var process = new Process();
        //process.StartInfo.FileName = wkHtmlToPdfPath;
        //process.StartInfo.Arguments = arguments;
        //process.StartInfo.UseShellExecute = false;
        //process.StartInfo.RedirectStandardOutput = true;
        //process.StartInfo.RedirectStandardError = true;
        //process.Start();
        //process.WaitForExit();
    }

    #region 想使用LibreOffice加密失敗
    /// <summary>
    /// 生成帶有密碼的pdf
    /// </summary>
    /// <param name="pdfBytes">最初沒有密碼的pdf數據</param>
    /// <returns></returns>
    //public async Task<byte[]> GeneratePdfWithPassword(byte[] pdfBytes, string password)
    //{
    //    var tempInputFile = Path.GetTempFileName();
    //    var tempOutputFile = Path.ChangeExtension(tempInputFile, ".pdf");

    //    try
    //    {
    //        await File.WriteAllBytesAsync(tempInputFile, pdfBytes);

    //        var startInfo = new ProcessStartInfo
    //        {
    //            FileName = _libreOfficePath,
    //            Arguments = $"--headless --convert-to pdf:writer_pdf_Export:{{\"EncryptFile\":{{\"type\":\"boolean\",\"value\":\"true\"}},\"DocumentOpenPassword\":{{\"type\":\"string\",\"value\":\"{password}\"}} {tempInputFile} --outdir \"{Path.GetDirectoryName(tempOutputFile)}\"",
    //            //Arguments = $@"--convert-to 'pdf:draw_pdf_Export:{{""EncryptFile"":{{""type"":""boolean"",""value"":""true""}},""DocumentOpenPassword"":{{""type"":""string"",""value"":""secret""}}}}' {tempInputFile}",
    //            UseShellExecute = false,
    //            RedirectStandardError = true,
    //            CreateNoWindow = true
    //        };

    //        using var process = new Process { StartInfo = startInfo };
    //        process.Start();

    //        var errorTask = process.StandardError.ReadToEndAsync();
    //        await process.WaitForExitAsync();

    //        if (process.ExitCode != 0)
    //        {
    //            var error = await errorTask;
    //            throw new Exception($"LibreOffice 轉換失敗。錯誤碼：{process.ExitCode}。錯誤信息：{error}");
    //        }

    //        if (File.Exists(tempOutputFile))
    //        {
    //            return await File.ReadAllBytesAsync(tempOutputFile);
    //        }
    //        else
    //        {
    //            throw new Exception("無法找到生成的 PDF 文件。");
    //        }
    //    }
    //    finally
    //    {
    //        if (File.Exists(tempInputFile)) File.Delete(tempInputFile);
    //        if (File.Exists(tempOutputFile)) File.Delete(tempOutputFile);
    //    }
    //}
    #endregion
}
