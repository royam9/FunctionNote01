using Microsoft.AspNetCore.Http;

namespace Models.ApiModels;

/// <summary>
/// 原版
/// </summary>
public class WordDocumentRequestModel
{
    public string StudentName { get; set; } = null!;

    public string ClassName { get; set; } = null!;

    public IEnumerable<IFormFile> Images {  get; set; } = null!;
}

/// <summary>
/// 個資使用同意書 Request Model
/// </summary>
public class PersonalDataUsageConsentFormModel
{
    /// <summary>
    /// 學生資訊 model
    /// </summary>
    public List<StudentDataModel>? StudentDataModels { get; set; }
    ///<summary>
    /// 簽名
    /// </summary>
    public IFormFile? Sign { get; set; }
}

/// <summary>
/// 學生資訊 model
/// </summary>
public class StudentDataModel
{
    /// <summary>
    /// 學生姓名
    /// </summary>
    public string StudentName { get; set; } = string.Empty;
    /// <summary>
    /// 教室名稱
    /// </summary>
    public string ClassName { get; set; } = string.Empty;
}