using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Text;

namespace Service;

public class ConvertHtmlToPdfService
{
    /// <summary>
    /// Html 轉換成 Pdf
    /// </summary>
    /// <param name="htmlContent">Html文字(UTF-8)</param>
    /// <return></return>
    public static byte[] ConvertUTF8HtmlToPdf(string htmlContent)
    {
        string wkhtmltopdfPath = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\..\Services\App\wkhtmltox\bin\wkhtmltopdf.exe"));

        // 確認 wkhtmltopdf.exe 是否存在
        if (!File.Exists(wkhtmltopdfPath))
        {
            throw new FileNotFoundException("The wkhtmltopdf executable was not found at the specified path.", wkhtmltopdfPath);
        }

        // 設置命令行參數，將 PDF 輸出到標準輸出流（stdout）
        string arguments = $"--page-size A4 --margin-top 20mm --margin-right 20mm --margin-bottom 20mm --margin-left 20mm --zoom 1.3 - -";

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

    /// <summary>
    /// Big5 Html 轉換成 Pdf
    /// </summary>
    /// <param name="htmlContent">Html文字(Big5)</param>
    /// <return></return>
    public static byte[] ConvertBig5HtmlToPdf(string htmlContent)
    {
        string wkhtmltopdfPath = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\..\Services\App\wkhtmltox\bin\wkhtmltopdf.exe"));

        // 確認 wkhtmltopdf.exe 是否存在
        if (!File.Exists(wkhtmltopdfPath))
        {
            throw new FileNotFoundException("The wkhtmltopdf executable was not found at the specified path.", wkhtmltopdfPath);
        }

        // 設置命令行參數，將 PDF 輸出到標準輸出流（stdout）
        string arguments = $"--page-size A4 --margin-top 20mm --margin-right 20mm --margin-bottom 20mm --margin-left 20mm --zoom 1.3 - -";

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
            using (StreamWriter standardInput = new StreamWriter(process.StandardInput.BaseStream, Encoding.GetEncoding("BIG5")))
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

    /// <summary>
    /// 替換Word另存新檔成html 裡面的參數
    /// </summary>
    /// <param name="htmlContent">參數</param>
    /// <remarks>參數樣式 {參數名稱}</remarks>
    /// <returns></returns>
    public static string ReplaceString(string htmlContent)
    {
        string pattern = @"\{[^{}]*>教室名稱<[^{}]*\}";

        // 使用正則來替換匹配的內容
        string result = Regex.Replace(htmlContent, pattern, match =>
        {
            // 刪除大括號
            string modified = match.Value.Trim('{', '}');

            // 替換 "教室名稱" 為 "音樂教室"
            modified = modified.Replace("教室名稱", "音樂教室");

            return modified;
        });

        return result;
    }
}
