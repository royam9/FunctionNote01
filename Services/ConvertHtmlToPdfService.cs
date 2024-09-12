using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Text;
using Services.Interfaces;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.html.simpleparser;
using System.Collections;

namespace Services;

public class ConvertHtmlToPdfService : IConvertHtmlToPdfService
{
    /// <summary>
    /// 是否為 Linux環境
    /// </summary>
    private readonly bool _isLinux;

    public ConvertHtmlToPdfService()
    {
        _isLinux = System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Linux);
    }

    #region 使用 wkhtmltopdf
    /// <summary>
    /// Html 轉換成 Pdf
    /// </summary>
    /// <param name="htmlContent">Html文字(UTF-8)</param>
    /// <return></return>
    public byte[] ConvertUTF8HtmlToPdf(string htmlContent)
    {

        string wkhtmltopdfPath = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\..\Services\Resource\wkhtmltox\bin\wkhtmltopdf.exe"));

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
    public byte[] ConvertBig5HtmlToPdf(string htmlContent)
    {
        string wkhtmltopdfPath = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\..\Services\Resource\wkhtmltox\bin\wkhtmltopdf.exe"));

        // 確認 wkhtmltopdf.exe 是否存在
        if (!File.Exists(wkhtmltopdfPath))
        {
            throw new FileNotFoundException("The wkhtmltopdf executable was not found at the specified path.", wkhtmltopdfPath);
        }

        // 設置命令行參數，將 PDF 輸出到標準輸出流（stdout）
        // 把 --zoom 改成 --disable-smart-shrinking 禁止智能縮放
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
    /// 非同步 Html 轉換成 Pdf
    /// </summary>
    /// <param name="htmlContent">Html文字(UTF-8)</param>
    /// <return></return>
    public async Task<byte[]> HtmlToPdf(string htmlContent)
    {
        // ..\..\..\..\Services\Resource\wkhtmltox\bin\wkhtmltopdf.exe
        // 正式專案 ..\..\..\Resources\PdfPackage\wkhtmltox\bin\wkhtmltopdf.exe
        string wkhtmltopdfPath;
        if (_isLinux)
        {
            wkhtmltopdfPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Rotativa/Linux/wkhtmltopdf");
        }
        else
        {
            // 公司專案位置
            // wkhtmltopdfPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources/PdfPackage/wkhtmltox/bin/wkhtmltopdf.exe");

            // 目前專案位置
            wkhtmltopdfPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "../Services/Resources/PdfPackage/wkhtmltox/bin/wkhtmltopdf.exe");
            // 改採
            wkhtmltopdfPath = Path.Combine(Directory.GetCurrentDirectory(), "../Services/Resources/PdfPackage/wkhtmltox/bin/wkhtmltopdf.exe");
        }

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
                await standardInput.WriteAsync(htmlContent);
            }

            // 使用 MemoryStream 來保存 PDF 數據
            using (MemoryStream memoryStream = new MemoryStream())
            {
                await process.StandardOutput.BaseStream.CopyToAsync(memoryStream);
                pdfBytes = memoryStream.ToArray();
            }

            await process.WaitForExitAsync();
            if (process.ExitCode != 0)
            {
                string error = await process.StandardError.ReadToEndAsync();
                throw new Exception($"wkhtmltopdf failed: {error}");
            }
        }

        return pdfBytes; // 返回生成的 PDF 數據
    }
    #endregion

    /// <summary>
    /// 替換Word另存新檔成html 裡面的參數
    /// </summary>
    /// <param name="htmlContent">參數</param>
    /// <remarks>參數樣式 {參數名稱}</remarks>
    /// <returns></returns>
    public string ReplaceString(string htmlContent)
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

    #region 使用iTextSharp嘗試 要安裝iTextSharp.LGPLv2.Core https://github.com/wkhtmltopdf/wkhtmltopdf?tab=readme-ov-file
    public static MemoryStream ITextSharp_Document(string html)
    {
        #region 文件的1      
        StringReader sr = new StringReader(html);
        //步骤1
        Document document = new Document(PageSize.A4, 10f, 10f, 10f, 0f);

        MemoryStream stream = new MemoryStream();
        //步骤2
        PdfWriter.GetInstance(document, stream);
        //步骤3
        document.Open();

        //创建一个样式表
        StyleSheet styles = new StyleSheet();
        ////设置默认字体的属性
        //styles.LoadTagStyle(HtmlTags.BODY, "encoding", "Identity-H");
        //styles.LoadTagStyle(HtmlTags.BODY, HtmlTags.FONT, "Tahoma");
        //styles.LoadTagStyle(HtmlTags.BODY, "size", "16pt");

        //FontFactory.Register(@"C:\Windows\Fonts\tahoma.ttf");
        FontFactory.Register(@"C:\Windows\Fonts\kaiu.ttf", "標楷體");

        var unicodeFontProvider = FontFactoryImp.Instance;
        unicodeFontProvider.DefaultEmbedding = BaseFont.EMBEDDED;
        unicodeFontProvider.DefaultEncoding = BaseFont.IDENTITY_H;

        var props = new Hashtable
        {
            // { "img_provider", new MyImageFactory() },
            { "font_factory", unicodeFontProvider } //始终使用Unicode字体
        };

        //步骤4
        //var objects = HtmlWorker.ParseToList(sr, styles, props);
        var objects = HtmlWorker.ParseToList(sr, styles);
        foreach (IElement element in objects)
        {
            // BaseFont baseFont = BaseFont.CreateFont("STSong-Light", "UniGB-UCS2-H", BaseFont.NOT_EMBEDDED);
            BaseFont baseFont = BaseFont.CreateFont(@"C:\Windows\Fonts\kaiu.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
            //BaseFont baseFont = BaseFont.CreateFont(@"C:\Windows\Fonts\georgia.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);

            iTextSharp.text.Font font = new iTextSharp.text.Font(baseFont);

            Paragraph temp = element as Paragraph;
            if (temp != null)
            {
                var chuck = temp[0] as Chunk;
                if (chuck != null)
                {
                    chuck.Font = font;
                    //document.Add(temp);
                }
            }

            document.Add(element);
        }

        document.Close();
        return stream;
        #endregion 

        #region 文件的2
        //StringReader sr = new StringReader(html);

        //// 步驟1: 創建新的 Document
        //Document document = new Document(PageSize.A4, 10f, 10f, 10f, 0f);

        //MemoryStream stream = new MemoryStream();
        //PdfWriter writer = PdfWriter.GetInstance(document, stream);

        //// 步驟2: 開啟 Document 以便寫入
        //document.Open();

        //// 步驟3: 註冊字體
        ////FontFactory.Register(@"C:\Windows\Fonts\msjh.ttc", "Microsoft JhengHei"); // 微軟正黑體
        //FontFactory.Register(@"C:\Windows\Fonts\mingliu.ttc", "MingLiu");
        //FontFactory.Register(@"C:\Windows\Fonts\kaiu.ttf", "標楷體");

        //// 步驟4: 創建 StyleSheet 並設置 CSS
        //StyleSheet styles = new StyleSheet();
        //// 這裡可以根據需要配置 CSS 樣式

        //// 步驟5: 使用 HtmlWorker 解析 HTML
        //List<IElement> elements = HtmlWorker.ParseToList(sr, styles);

        //foreach (IElement element in elements)
        //{
        //    document.Add(element);
        //}

        //// 步驟6: 關閉 Document 並返回記憶體流
        //document.Close();
        //return stream;
        #endregion
    }

    public static void TextSharpToPDF(string html)
    {
        #region AI的
        string outputPath = @"C:\Users\TWJOIN\Desktop\Komo\轉換測試檔案\使用iTextSharp轉換.pdf";

        // 創建文件流以寫入PDF
        using (FileStream fs = new FileStream(outputPath, FileMode.Create, FileAccess.Write, FileShare.None))
        {
            // 創建Document對象
            Document document = new Document(PageSize.A4);

            // 創建PdfWriter對象
            PdfWriter.GetInstance(document, fs);

            // 開啟Document
            document.Open();

            // 設定中文字型
            //BaseFont baseFont = BaseFont.CreateFont(@"C:\Windows\Fonts\kaiu.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            //Font chineseFont = new Font(baseFont, 12); // 字型大小設置在 HTML 中

            //BaseFont baseFont2 = BaseFont.CreateFont(@"C:\Windows\Fonts\cambria.ttc", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            //Font chineseFont2 = new Font(baseFont2, 12);

            //BaseFont baseFont3 = BaseFont.CreateFont(@"C:\Windows\Fonts\georgia.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            //Font chineseFont3 = new Font(baseFont3, 12);

            // 使用FontFactory註冊字型
            FontFactory.Register(@"C:\Windows\Fonts\kaiu.ttf", "標楷體");
            //FontFactory.Register(@"C:\Windows\Fonts\cambria.ttc", "Cambria Math");
            FontFactory.Register(@"C:\Windows\Fonts\georgia.ttf", "Georgia");
            FontFactory.Register(@"C:\Windows\Fonts\mingliu.ttc", "新細明體");
            //FontFactory.Register(@"C:\Windows\Fonts\LiberationSerif - Regular.ttf", "Liberation Serif");
            FontFactory.Register(@"C:\Windows\Fonts\roman.fon", "roman");

            // 使用HTMLWorker來解析HTML字串
            using (StringReader sr = new StringReader(html))
            {
                HtmlWorker htmlWorker = new HtmlWorker(document);
                htmlWorker.Parse(sr);
            }

            // 關閉Document
            document.Close();
        }
        #endregion
    }

    public static string CreatePdfFile(string htmlContent)
    {
        string pattern = @"<span\s+lang=EN-US>\s*&nbsp;\s*\{學生姓名\}\s*</span>";
        htmlContent = Regex.Replace(htmlContent, pattern, "小明");

        return htmlContent;
    }

    public static void StackFlow(string htmlString)
    {
        //Create a byte array that will eventually hold our final PDF
        Byte[] bytes;

        //Boilerplate iTextSharp setup here
        //Create a stream that we can write to, in this case a MemoryStream
        using (var ms = new MemoryStream())
        {

            //Create an iTextSharp Document which is an abstraction of a PDF but **NOT** a PDF
            using (var doc = new Document())
            {

                //Create a writer that's bound to our PDF abstraction and our stream
                using (var writer = PdfWriter.GetInstance(doc, ms))
                {

                    //Open the document for writing
                    doc.Open();

                    BaseFont baseFont = BaseFont.CreateFont(@"C:\Windows\Fonts\kaiu.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                    FontFactory.Register(@"C:\Windows\Fonts\kaiu.ttf", "標楷體");
                    FontFactory.Register(@"C:\Windows\Fonts\georgia.ttf", "Georgia");

                    //Our sample HTML and CSS
                    var example_html = @"<p>This <em>is </em><span class=""headline"" style=""text-decoration: underline;"">some</span> <strong>sample <em> text</em></strong><span style=""color: red;"">!!!</span></p>";
                    var example_css = @".headline{font-size:200%}";

                    /**************************************************
                     * Example #1                                     *
                     *                                                *
                     * Use the built-in HTMLWorker to parse the HTML. *
                     * Only inline CSS is supported.                  *
                     * ************************************************/

                    //Create a new HTMLWorker bound to our document
                    using (var htmlWorker = new iTextSharp.text.html.simpleparser.HtmlWorker(doc))
                    {

                        //HTMLWorker doesn't read a string directly but instead needs a TextReader (which StringReader subclasses)
                        using (var sr = new StringReader(htmlString))
                        {

                            //Parse the HTML
                            htmlWorker.Parse(sr);
                        }
                    }

                    doc.Close();
                }
            }

            //After all of the PDF "stuff" above is done and closed but **before** we
            //close the MemoryStream, grab all of the active bytes from the stream
            bytes = ms.ToArray();
        }

        //Now we just need to do something with those bytes.
        //Here I'm writing them to disk but if you were in ASP.Net you might Response.BinaryWrite() them.
        //You could also write the bytes to a database in a varbinary() column (but please don't) or you
        //could pass them to another function for further PDF processing.
        var testFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "test.pdf");
        System.IO.File.WriteAllBytes(testFile, bytes);
    }

    public static void GeneratePdf(string htmlContent, string outputPath)
    {
        // 獲取執行目錄的路徑
        string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;

        // 設置 wkhtmltopdf 的路徑
        //string wkhtmltopdfPath = Path.Combine(baseDirectory, "App", "wkhtmltox", "bin", "wkhtmltopdf.exe");
        string wkhtmltopdfPath = @"C:\Users\TWJOIN\source\repos\kumon-dot-net-core-backend\Services\App\wkhtmltox\bin\wkhtmltopdf.exe";

        // 確認 wkhtmltopdf.exe 是否存在
        if (!File.Exists(wkhtmltopdfPath))
        {
            throw new FileNotFoundException("The wkhtmltopdf executable was not found at the specified path.", wkhtmltopdfPath);
        }

        // 創建 HTML 文件
        string htmlFilePath = Path.Combine(Path.GetTempPath(), "temp.html");
        File.WriteAllText(htmlFilePath, htmlContent);

        // 設置命令行參數
        //string arguments = $"\"{htmlFilePath}\" \"{outputPath}\"";
        string arguments = $"--page-size A4 --margin-top 20mm --margin-right 20mm --margin-bottom 20mm --margin-left 20mm \"{htmlFilePath}\" \"{outputPath}\"";

        // 創建進程啟動信息
        ProcessStartInfo processStartInfo = new ProcessStartInfo
        {
            FileName = wkhtmltopdfPath,
            Arguments = arguments,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };

        // 啟動進程
        using (Process process = Process.Start(processStartInfo))
        {
            process.WaitForExit();
            if (process.ExitCode != 0)
            {
                string error = process.StandardError.ReadToEnd();
                throw new Exception($"wkhtmltopdf failed: {error}");
            }
        }

        // 清理暫時的 HTML 文件
        File.Delete(htmlFilePath);
    }
    #endregion
}
