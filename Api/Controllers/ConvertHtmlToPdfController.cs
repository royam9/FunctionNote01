using Microsoft.AspNetCore.Mvc;
using Services;
using System.Text;

namespace Api.Controllers;

[Route("api/[controller]")]
[ApiController]
public class ConvertHtmlToPdfController : Controller
{
    /// <summary>
    /// Html 轉換成 Pdf
    /// </summary>
    /// <param name="htmlContent">html文字(UTF-8編碼)</param>
    /// <returns></returns>
    [HttpPost]
    [Route("ConvertHtmlToPdf")]
    public IActionResult ConvertHtmlToPdf(string htmlContent)
    {
        // 先假設有 htmlstring
        htmlContent = System.IO.File.ReadAllText(@"C:\Users\TWJOIN\Desktop\Komo\已篩選的網頁\個資使用同意書_中英文_20240722_草案_已篩選_UTF8.htm");

        // 使用套件先轉換成html
        //string wordFilePath = @"C:\Users\TWJOIN\Desktop\Komo\個資使用同意書_中英文_20240722_草案.docx";
        //string htmlFilePath = @"C:\Users\TWJOIN\Desktop\Komo\ConvertHTML.html";
        //htmlContent = ConvertWordToPDFService.ConvertWordToHtml(wordFilePath, htmlFilePath);

        // 如果想要手動把big5編碼轉成utf8
        //byte[] utf8Bytes = Encoding.Convert(Encoding.GetEncoding("big5"), Encoding.UTF8, Encoding.GetEncoding("big5").GetBytes(htmlContent));
        //string utf8Content = Encoding.UTF8.GetString(utf8Bytes);

        // 如果replace一直失敗
        //string decodedHtml = System.Net.WebUtility.HtmlDecode(utf8Content);

        byte[] pdfbytes = ConvertHtmlToPdfService.ConvertUTF8HtmlToPdf(htmlContent);

        return File(pdfbytes, "application/pdf", "使用wkhtmltopdf轉換成pdf");
    }

    /// <summary>
    /// Big5 Html 轉換成 Pdf
    /// </summary>
    /// <param name="htmlContent">html文字(Big5編碼)</param>
    /// <returns></returns>
    [HttpPost]
    [Route("ConvertBig5HtmlToPdf")]
    public IActionResult ConvertBig5HtmlToPdf(string htmlContent)
    {
        // 要正確讀取要這段程式碼 可能要安裝 System.Text.Encoding.CodePages
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        
        // 然後這樣
        htmlContent = System.IO.File
            .ReadAllText(@"D:\hi\CodePractice\FunctionNote01\Api\Resource\Template\個資使用同意書_中英文_20240722_草案.html", 
            Encoding.GetEncoding("big5"));

        byte[] pdfbytes = ConvertHtmlToPdfService.ConvertBig5HtmlToPdf(htmlContent);

        return File(pdfbytes, "application/pdf", "使用wkhtmltopdf轉換成pdf");
    }
}
