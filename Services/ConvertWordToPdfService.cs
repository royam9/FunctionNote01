using Services.Interfaces;
using System.Diagnostics;

namespace Services;

public class ConvertWordToPdfService : IConvertWordToPdfService
{
    /// <summary>
    /// 是否為 Linux環境
    /// </summary>
    private readonly bool _isLinux;
    /// <summary>
    /// LibreOffice執行檔位置
    /// </summary>
    private readonly string _libreOfficePath;

    public ConvertWordToPdfService()
    {
        _isLinux = System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Linux);

        if (_isLinux)
        {
            _libreOfficePath = "/lib/libreoffice/program/soffice";
        }
        else
        {
            string programFilesPath = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            _libreOfficePath = Path.Combine(programFilesPath, "LibreOffice", "program", "soffice.exe");
        }
    }

    /// <summary>
    /// Word 轉換成 Pdf - 使用LibreOffice
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

            var startInfo = new ProcessStartInfo
            {
                FileName = _libreOfficePath,
                Arguments = $"--headless --writer --convert-to pdf \"{tempInputFile}\" --outdir \"{Path.GetDirectoryName(tempOutputFile)}\"",
                UseShellExecute = false,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            using var process = new Process { StartInfo = startInfo };
            process.Start();

            var errorTask = process.StandardError.ReadToEndAsync();
            await process.WaitForExitAsync();

            if (process.ExitCode != 0)
            {
                var error = await errorTask;
                throw new Exception($"LibreOffice 轉換失敗。錯誤碼：{process.ExitCode}。錯誤信息：{error}");
            }

            if (File.Exists(tempOutputFile))
            {
                return await File.ReadAllBytesAsync(tempOutputFile);
            }
            else
            {
                throw new Exception("無法找到生成的 PDF 文件。");
            }
        }
        finally
        {
            if (File.Exists(tempInputFile)) File.Delete(tempInputFile);
            if (File.Exists(tempOutputFile)) File.Delete(tempOutputFile);
        }
    }
}
