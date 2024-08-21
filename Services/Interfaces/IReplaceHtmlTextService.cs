using HtmlAgilityPack;
using Microsoft.AspNetCore.Http;
using Models.ApiModels;

namespace Services.Interfaces;

public interface IReplaceHtmlTextService
{
    /// <summary>
    /// 編輯HTML文件
    /// </summary>
    /// <param name="param">參數</param>
    /// <returns></returns>
    Task<byte[]> UpdateHtmlDocument(WordDocumentRequestModel param);

    /// <summary>
    /// 插入文字到指定位置
    /// </summary>
    /// <param name="document">主文件</param>
    /// <param name="oldText">佔位符文字</param>
    /// <param name="newText">插入文字</param>
    void ReplaceText(HtmlDocument document, string oldText, string newText);

    /// <summary>
    /// 插入圖片到指定位置
    /// </summary>
    /// <param name="document">主文件</param>
    /// <param name="images">上傳圖片</param>
    /// <returns></returns>
    Task InsertImagesIntoHtml(HtmlDocument document, IEnumerable<IFormFile> images);
}
