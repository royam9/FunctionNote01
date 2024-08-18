using HtmlAgilityPack;
using Microsoft.AspNetCore.Http;
using Models.ApiModels;
using SixLabors.ImageSharp;
using static System.Net.Mime.MediaTypeNames;

namespace Services;

/// <summary>
/// Word文檔
/// </summary>
public class WordDocumentService 
{
    /// <summary>
    /// 編輯HTML文件
    /// </summary>
    /// <param name="param">參數</param>
    /// <returns></returns>
    public async Task<byte[]> UpdateHtmlDocument(WordDocumentRequestModel param)
    {
        // 獲取HTML檔案的路徑
        var templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "Template", "WordDocument.htm");
        var outputPath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.htm");

        // 讀取 HTML 檔案
        string htmlContent = await File.ReadAllTextAsync(templatePath);

        // 解析 HTML 文檔
        var document = new HtmlDocument();
        document.LoadHtml(htmlContent);

        // 替換文本
        ReplaceText(document, "學生姓名", param.StudentName);
        ReplaceText(document, "教室名稱", param.ClassName);
        ReplaceText(document, "{", "");
        ReplaceText(document, "}", "");

        // 插入圖片
        await InsertImagesIntoHtml(document, param.Images!);

        // 儲存修改後的 HTML 文件
        document.Save(outputPath);

        return await File.ReadAllBytesAsync(outputPath);
    }

    /// <summary>
    /// 插入文字到指定位置
    /// </summary>
    /// <param name="document">主文件</param>
    /// <param name="oldText">佔位符文字</param>
    /// <param name="newText">插入文字</param>
    private void ReplaceText(HtmlDocument document, string oldText, string newText)
    {
        var nodes = document.DocumentNode.SelectNodes("//span");
        if (nodes != null)
        {
            foreach (var node in nodes)
            {
                if (node.InnerText.Contains(oldText))
                {
                    node.InnerHtml = node.InnerHtml.Replace(oldText, newText);
                }
            }
        }
    }

    /// <summary>
    /// 插入圖片到指定位置
    /// </summary>
    /// <param name="document">主文件</param>
    /// <param name="images">上傳圖片</param>
    /// <returns></returns>
    private async Task InsertImagesIntoHtml(HtmlDocument document, IEnumerable<IFormFile> images)
    {
        var placeholderNodes = document.DocumentNode.SelectNodes("//span");
        if (placeholderNodes != null && images != null && images.Any())
        {
            foreach (var span in placeholderNodes)
            {
                if (span.InnerText.Contains("上傳圖片"))
                {
                    var imageHtmlList = new List<string>();

                    foreach (var imageFile in images)
                    {
                        var imageFileName = Path.GetFileNameWithoutExtension(imageFile.FileName);
                        var imageExtension = Path.GetExtension(imageFile.FileName);
                        var imagePath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}{imageExtension}");

                        using (var stream = new FileStream(imagePath, FileMode.Create))
                        {
                            await imageFile.CopyToAsync(stream);
                        }

                        // 獲取圖片尺寸
                        int imageWidth, imageHeight;
                        using (var image = SixLabors.ImageSharp.Image.Load(imagePath))
                        {
                            imageWidth = image.Width;
                            imageHeight = image.Height;
                        }

                        // 動態設置 MIME 類型
                        var mimeType = imageExtension.ToLower() switch
                        {
                            ".jpg" or ".jpeg" => "image/jpeg",
                            ".png" => "image/png",
                            ".gif" => "image/gif",
                            _ => "image/jpeg"
                        };

                        // 使用 base64 編碼將圖片嵌入 HTML
                        var imageBase64 = Convert.ToBase64String(await File.ReadAllBytesAsync(imagePath));
                        var imageHtml = $"<img src=\"data:{mimeType};base64,{imageBase64}\" style=\"max-width:100%; height:auto;\" />";

                        // 將圖片 HTML 添加到列表中
                        imageHtmlList.Add(imageHtml);
                    }

                    // 將所有圖片 HTML 連接成一個字串
                    var combinedImageHtml = string.Join(" ", imageHtmlList);

                    span.InnerHtml = span.InnerHtml.Replace("上傳圖片", combinedImageHtml);
                }
            }
        }
    }
}