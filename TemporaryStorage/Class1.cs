using System.Diagnostics;
using System.Text;
using static System.Net.Mime.MediaTypeNames;

namespace TemporaryStorage;

public class Class1
{
    /// <summary>
    /// Html 轉換成 PDF - wkhtmltopdf -UTF8編碼
    /// </summary>
    /// <param name="htmlContent">Html文字</param>
    /// <returns></returns>
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
    /// Html 轉換成 PDF - wkhtmltopdf
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

    /// <summary>
    /// 編輯註冊入學及修改個資Word文件 - 寫法精簡與修正
    /// </summary>
    /// <param name="param">輸入參數</param>
    /// <param name="oldData">修改個資舊資料</param>
    /// <returns></returns>
    //private async Task<byte[]> UpdateRegisterAndModifyWordDocument(RegisterAndModifyWordDocumentModel param, ModifyStudentRegisterDataResponseModel? oldData = null)
    //{
    //    var templatePath = string.Empty;
    //    var todayDate = DateTime.Now;

    //    // 編輯Word文件類型
    //    switch (param.UpdateType)
    //    {
    //        case NotifyTypeEnum.Register:
    //            templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "Template", "註冊入學.docx");
    //            break;

    //        case NotifyTypeEnum.Modify:
    //            templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "Template", "修改個人資訊.docx");
    //            break;
    //    }

    //    // 讀取 Word 檔案
    //    byte[] templateBytes = await File.ReadAllBytesAsync(templatePath);

    //    using (MemoryStream ms = new())
    //    {
    //        // 將模板複製到 MemoryStream
    //        await new MemoryStream(templateBytes).CopyToAsync(ms);
    //        ms.Position = 0;

    //        using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(ms, true))
    //        {
    //            var body = wordDoc.MainDocumentPart!.Document.Body;

    //            // 讀取書籤
    //            var bookmarkStarts = wordDoc.MainDocumentPart?.RootElement?.Descendants<BookmarkStart>().ToList();

    //            #region 樣式設定
    //            // 使用第一個 bookmarkStart 作爲字形基礎 (文件標題)
    //            BookmarkStart? firstBookmatkStart = bookmarkStarts?.FirstOrDefault();

    //            // 一般樣式設定:預設字型樣式(size:12 字體:讀取該Paragraph設定)
    //            RunProperties? textStyleProperties = new(new FontSize { Val = "24" });
    //            RunProperties? titleProperties = new(new FontSize { Val = "44" }, new Bold { Val = true });

    //            // 取得該Paragraph的RunFonts
    //            var parentParagraph = firstBookmatkStart?.Parent as Paragraph;
    //            var runFonts = parentParagraph?.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<RunFonts>()?.CloneNode(true);
    //            var titleRunFonts = parentParagraph?.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<RunFonts>()?.CloneNode(true);

    //            if (runFonts != null)
    //            {
    //                textStyleProperties.AppendChild(runFonts);
    //                titleProperties.AppendChild(titleRunFonts);
    //            }
    //            #endregion

    //            for (int i = 0; i < bookmarkStarts?.Count; i++)
    //            {
    //                BookmarkStart? bookmarkStart = bookmarkStarts[i];
    //                var newRunProperty = textStyleProperties.CloneNode(true);

    //                // 替換文本
    //                string? bookmarkStartName = bookmarkStart.Name?.ToString();

    //                if ((bookmarkStartName?.Contains("StudentName") ?? false)
    //                    && bookmarkStartName != "OldStudentName")
    //                {
    //                    // StudentName1(標題)
    //                    if (bookmarkStartName == "StudentName1")
    //                    {
    //                        newRunProperty = titleProperties.CloneNode(true);
    //                        bookmarkStart.Parent?.InsertAfter(new Run(newRunProperty, new Text(param.StudentName ?? oldData.StudentName)), bookmarkStart);
    //                    }
    //                    else
    //                    {
    //                        bookmarkStart.Parent?.InsertAfter(new Run(newRunProperty, new Text(param.StudentName)), bookmarkStart);
    //                    }

    //                    continue;
    //                }

    //                if (bookmarkStartName?.Contains("StudentClassName") ?? false)
    //                {
    //                    bookmarkStart.Parent?.InsertAfter((new Run(newRunProperty, new Text(param.StudentClassName))), bookmarkStart);
    //                    continue;
    //                }

    //                switch (bookmarkStartName)
    //                {
    //                    #region 註冊/修改資訊
    //                    case "CurrentGrade":
    //                        bookmarkStart.Parent?.InsertAfter(new Run(newRunProperty, new Text(param.CurrentGrade.GetDescription())), bookmarkStart);
    //                        break;
    //                    case "EnrolledTime":
    //                        bookmarkStart.Parent?.InsertAfter(new Run(newRunProperty, new Text(param.EnrolledTime)), bookmarkStart);
    //                        break;
    //                    case "GenderType":
    //                        bookmarkStart.Parent?.InsertAfter(new Run(newRunProperty, new Text(param.GenderType.GetDescription())), bookmarkStart);
    //                        break;
    //                    case "RegisterDate":
    //                        bookmarkStart.Parent?.InsertAfter(new Run(newRunProperty, new Text(param.RegisterDate)), bookmarkStart);
    //                        break;
    //                    case "RepresentativeBirthdate":
    //                        bookmarkStart.Parent?.InsertAfter(new Run(newRunProperty, new Text(param.RepresentativeBirthdate)), bookmarkStart);
    //                        break;
    //                    case "RepresentativeContactAddress":
    //                        bookmarkStart.Parent?.InsertAfter(new Run(newRunProperty, new Text(param.RepresentativeContactAddress)), bookmarkStart);
    //                        break;
    //                    case "RepresentativeDistrictCode":
    //                        bookmarkStart.Parent?.InsertAfter(new Run(newRunProperty, new Text(param.RepresentativeDistrictCode)), bookmarkStart);
    //                        break;
    //                    case "RepresentativeEmail":
    //                        bookmarkStart.Parent?.InsertAfter(new Run(newRunProperty, new Text(param.RepresentativeEmail)), bookmarkStart);
    //                        break;
    //                    case "RepresentativeMobilePhone":
    //                        bookmarkStart.Parent?.InsertAfter(new Run(newRunProperty, new Text(param.RepresentativeMobilePhone)), bookmarkStart);
    //                        break;
    //                    case "RepresentativeName":
    //                        bookmarkStart.Parent?.InsertAfter(new Run(newRunProperty, new Text(param.RepresentativeName)), bookmarkStart);
    //                        break;
    //                    case "RepresentativeRelationshipType":
    //                        bookmarkStart.Parent?.InsertAfter(new Run(newRunProperty, new Text(param.RepresentativeRelationshipType.GetDescription())), bookmarkStart);
    //                        break;
    //                    case "RepresentativeTelephone":
    //                        bookmarkStart.Parent?.InsertAfter(new Run(newRunProperty, new Text(param.RepresentativeTelephone)), bookmarkStart);
    //                        break;
    //                    case "StudentBirthDate":
    //                        bookmarkStart.Parent?.InsertAfter(new Run(newRunProperty, new Text(param.StudentBirthDate)), bookmarkStart);
    //                        break;
    //                    case "StudentSchoolName":
    //                        bookmarkStart.Parent?.InsertAfter(new Run(newRunProperty, new Text(param.StudentSchoolName)), bookmarkStart);
    //                        break;
    //                    case "SubjectTypes":
    //                        // 學習科目要顯示的文字
    //                        string subjectList = string.Join("、",
    //                            param.SubjectTypes.Select(s => s.GetDescription()));

    //                        bookmarkStart.Parent?.InsertAfter(new Run(newRunProperty, new Text(subjectList)), bookmarkStart);
    //                        break;
    //                    case "TodayDate":
    //                        bookmarkStart.Parent?.InsertAfter(new Run(newRunProperty, new Text(todayDate.ToString(ConstString.DateStringFormat))), bookmarkStart);
    //                        break;
    //                    #endregion
    //                    #region 修改前資訊
    //                    case "OldRepresentativeName":
    //                        bookmarkStart.Parent?.InsertAfter(new Run(newRunProperty, new Text(oldData.ParentName)), bookmarkStart);
    //                        break;
    //                    case "OldRepresentativeMobilePhone":
    //                        bookmarkStart.Parent?.InsertAfter(new Run(newRunProperty, new Text(oldData.MobilePhoneNumber)), bookmarkStart);
    //                        break;
    //                    case "OldRepresentativeTelephone":
    //                        bookmarkStart.Parent?.InsertAfter(new Run(newRunProperty, new Text(oldData.PhoneNumber)), bookmarkStart);
    //                        break;
    //                    case "OldRepresentativeContactAddress":
    //                        bookmarkStart.Parent?.InsertAfter(new Run(newRunProperty, new Text(oldData.Address)), bookmarkStart);
    //                        break;
    //                    case "OldRepresentativeDistrictCode":
    //                        bookmarkStart.Parent?.InsertAfter(new Run(newRunProperty, new Text(oldData.PostalCode)), bookmarkStart);
    //                        break;
    //                    case "OldRepresentativeEmail":
    //                        bookmarkStart.Parent?.InsertAfter(new Run(newRunProperty, new Text(oldData.EmailAddress)), bookmarkStart);
    //                        break;
    //                    case "OldRepresentativeBirthdate":
    //                        bookmarkStart.Parent?
    //                        .InsertAfter(
    //                        new Run(newRunProperty,
    //                            new Text(
    //                                oldData?.ParentBirthDate?.ToString("yyyy-MM-dd") ?? string.Empty)),
    //                            bookmarkStart);
    //                        break;
    //                    case "OldStudentName":
    //                        bookmarkStart.Parent?.InsertAfter(new Run(newRunProperty, new Text(oldData.StudentName)), bookmarkStart);
    //                        break;
    //                    case "OldGenderType":
    //                        if (Enum.TryParse(oldData?.Gender, true, out GenderTypeEnum oldGenderType))
    //                        {
    //                            bookmarkStart.Parent?.InsertAfter(new Run(newRunProperty, new Text(oldGenderType.GetDescription())), bookmarkStart);
    //                        }
    //                        break;
    //                    case "OldStudentBirthDate":
    //                        bookmarkStart.Parent?
    //                        .InsertAfter(
    //                        new Run(newRunProperty,
    //                            new Text(oldData?.BirthDate?.ToString(ConstString.DateStringFormat) ?? string.Empty)),
    //                        bookmarkStart);
    //                        break;
    //                    case "OldCurrentGrade":
    //                        bookmarkStart.Parent?.InsertAfter(new Run(newRunProperty, new Text(oldData.SchoolGrade)), bookmarkStart);
    //                        break;
    //                    case "OldStudentSchoolName":
    //                        bookmarkStart.Parent?.InsertAfter(new Run(newRunProperty, new Text(oldData.StudentSchool)), bookmarkStart);
    //                        break;
    //                    case "OldRepresentativeRelationshipType":
    //                        bookmarkStart.Parent?.InsertAfter(new Run(newRunProperty, new Text(oldData.Relationship)), bookmarkStart);
    //                        break;
    //                    #endregion
    //                    default:
    //                        break;
    //                }
    //            }

    //            wordDoc.MainDocumentPart?.Document.Save();
    //        }
    //        return ms.ToArray();
    //    }
    //}


}
