using Microsoft.AspNetCore.Http;

namespace Models.ApiModels;

public class WordDocumentRequestModel
{
    public string StudentName { get; set; } = null!;

    public string ClassName { get; set; } = null!;

    public IEnumerable<IFormFile> Images {  get; set; } = null!;
}
