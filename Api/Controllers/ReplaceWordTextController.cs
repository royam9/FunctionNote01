using Microsoft.AspNetCore.Mvc;
using Services.Interfaces;

namespace Api.Controllers;

[ApiController]
[Route("[Controller]")]
public class ReplaceWordTextController : Controller
{
    private readonly IReplaceWordTextService _replaceWordTextService;

    public ReplaceWordTextController(IReplaceWordTextService replaceWordTextService)
    {
        _replaceWordTextService = replaceWordTextService;
    }

    /// <summary>
    /// 編輯Word文檔
    /// </summary>
    /// <remarks>使用ParagraphMarkRunProperties</remarks>
    /// <returns></returns>
    [HttpPost]
    [Route("ReplaceText")]
    public async Task<IActionResult> ReplaceText()
    {
        string filePath = @"C:\Users\TWJOIN\Desktop\Komo\轉換測試檔案\註冊入學設定標楷體.docx";

        var result = await _replaceWordTextService.UpdateWordDocument_SearchParagraphProperties(filePath);

        return File(result, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "ParagraphMarkRunProperties.docx");
    }
}
