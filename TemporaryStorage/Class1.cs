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

}
