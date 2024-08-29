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
}
