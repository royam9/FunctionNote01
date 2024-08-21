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
}
