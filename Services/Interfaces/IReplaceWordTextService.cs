using Models.ApiModels;

namespace Services.Interfaces;

public interface IReplaceWordTextService
{
    /// <summary>
    /// 編輯Word檔
    /// </summary>
    /// <param name="model">參數</param>
    /// <returns></returns>
    Task<byte[]> UpdateWordDocument(WordDocumentRequestModel model);
    /// <summary>
    /// Word檔文字插補
    /// </summary>
    /// <param name="wordFilePath">Word檔路徑</param>
    /// <remarks>找尋Paragraph裡面的RunProperties</remarks>
    /// <returns></returns>
    Task<byte[]> UpdateWordDocument_SearchParagraphProperties(string wordFilePath);

}
