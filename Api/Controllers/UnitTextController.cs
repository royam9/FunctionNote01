using Api.TestServices;
using DocumentFormat.OpenXml.Bibliography;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.AspNetCore.Mvc;
using PDF_Tests;

namespace Api.Controllers;

[ApiController]
[Route("[controller]")]
public class UnitTextController : Controller
{
    /// <summary>
    /// PDF加入頁碼
    /// </summary>
    /// <returns></returns>
    [HttpPost]
    [Route("AddPdfPageNum")]
    public async Task<IActionResult> AddPdfPageNum()
    {
        var bytes = await System.IO.File.ReadAllBytesAsync(@"C:\Users\TWJOIN\Desktop\Komo\頁碼加入範本.pdf");

        var result = AddPdfPageNumService.AddPageNumbers(bytes);

        await System.IO.File.WriteAllBytesAsync(@"C:\Users\TWJOIN\Desktop\Komo\轉換測試檔案\頁碼加入範本.pdf", result);

        return Ok();
    }

    /// <summary>
    /// 使用iTextSharp加入頁碼
    /// </summary>
    /// <returns></returns>
    [HttpPost]
    [Route("AddPdfPageNum_iTextSharp")]
    public async Task<IActionResult> AddPdfPageNum_iTextSharp()
    {
        var inputPdfBytes = await System.IO.File.ReadAllBytesAsync(@"C:\Users\TWJOIN\Desktop\Komo\頁碼加入範本.pdf");

        using (MemoryStream outputPdfStream = new MemoryStream())
        {
            // 讀取既有 PDF
            PdfReader pdfReader = new PdfReader(inputPdfBytes);

            // 創建文件和 PdfStamper，用來修改 PDF
            using (PdfStamper pdfStamper = new PdfStamper(pdfReader, outputPdfStream))
            {
                // 設置 BaseFont
                //BaseFont baseFont = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.WINANSI, BaseFont.EMBEDDED);
                //BaseFont baseFont = BaseFont.CreateFont(@"C:\Windows\Fonts\kaiu.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);

                ITextSharpAddPdfPageNumberService pageEventHelper = new ITextSharpAddPdfPageNumberService();
                pdfStamper.Writer.PageEvent = pageEventHelper;

                // 關閉 PdfStamper 並寫入新 PDF
                pdfStamper.Close();
            }

            // 返回新 PDF 的 byte[]
            await System.IO.File.WriteAllBytesAsync(@"C:\Users\TWJOIN\Desktop\Komo\轉換測試檔案\頁碼測試.pdf", outputPdfStream.ToArray());

            return Ok();
        }
    }

    [HttpPost]
    [Route("AddPdfPageNum_iText_v2")]
    public async Task<IActionResult> AddPdfPageNum_iText_v2()
    {
        //string inputPdfPath = @"path\to\your\input.pdf";
        //string outputPdfPath = @"path\to\your\output.pdf";

        //// 打開現有 PDF 文件
        //using (PdfReader reader = new PdfReader(inputPdfPath))
        //using (PdfWriter writer = new PdfWriter(outputPdfPath))
        //using (PdfDocument pdfDoc = new PdfDocument(reader, writer))
        //{
        //    Document document = new Document(pdfDoc);
        //    int numberOfPages = pdfDoc.GetNumberOfPages();

        //    // 遍歷每一頁
        //    for (int i = 1; i <= numberOfPages; i++)
        //    {
        //        // 獲取頁面大小
        //        PdfPage page = pdfDoc.GetPage(i);
        //        Rectangle pageSize = page.GetPageSize();

        //        // 創建 PdfCanvas 對象
        //        PdfCanvas canvas = new PdfCanvas(page);

        //        // 設置頁碼樣式
        //        canvas.BeginText();
        //        canvas.SetFontAndSize(PdfFontFactory.CreateFont(), 12);
        //        canvas.MoveText(pageSize.GetLeft(40), pageSize.GetBottom(30)); // 設置頁碼位置
        //        canvas.ShowText($"Page {i} of {numberOfPages}");
        //        canvas.EndText();
        //    }

        //    document.Close();
        //}

        return Ok();
    }

    /// <summary>
    /// 使用iTextSharp加入頁碼
    /// </summary>
    /// <returns></returns>
    [HttpPost]
    [Route("AddPdfPageNum_iTextSharp_SlackFlow")]
    public async Task<IActionResult> AddPdfPageNum_iTextSharp_SlackFlow()
    {
        TwoColumnHeaderFooter tool = new TwoColumnHeaderFooter();
        tool.Title = "Your Document Title";
        tool.HeaderLeft = "Left Header Text";
        tool.HeaderRight = "Right Header Text";
        tool.HeaderFont = FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL);
        tool.FooterFont = FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL);

        using (Stream inputPdfStream = new FileStream(@"C:\Users\TWJOIN\Desktop\Komo\頁碼加入範本.pdf", FileMode.Open, FileAccess.Read, FileShare.Read))
        using (Stream outputPdfStream = new FileStream(@"C:\Users\TWJOIN\Desktop\Komo\轉換測試檔案\頁碼測試.pdf", FileMode.Create, FileAccess.Write, FileShare.None))
        {
            PdfReader reader = new PdfReader(inputPdfStream);
            Document document = new Document(reader.GetPageSize(1));
            PdfWriter writer = PdfWriter.GetInstance(document, outputPdfStream);

            writer.PageEvent = new ITextSharpAddPdfPageNumberService();

            document.Open();

            for (int i = 1; i <= reader.NumberOfPages; i++)
            {
                document.NewPage();
                PdfImportedPage page = writer.GetImportedPage(reader, i);
                PdfContentByte cb = writer.DirectContent;
                cb.AddTemplate(page, 0, 0);
            }

            document.Close();
            writer.Close();
            reader.Close();
        }
        return Ok();
    }

    /// <summary>
    /// 使用iTextSharp加入頁碼_使用MemorySteam
    /// </summary>
    /// <returns></returns>
    [HttpPost]
    [Route("AddPdfPageNum_iTextSharp_SlackFlow_MemorySteam")]
    public async Task<IActionResult> AddPdfPageNum_iTextSharp_SlackFlow_MemorySteam()
    {
        var pdfBytes = await System.IO.File.ReadAllBytesAsync(@"C:\Users\TWJOIN\Desktop\Komo\頁碼加入範本.pdf");

        using(MemoryStream inputPdfStream = new MemoryStream(pdfBytes))
        using (MemoryStream outputPdfStream = new MemoryStream())
        {
            PdfReader reader = new PdfReader(inputPdfStream);
            Document document = new Document(reader.GetPageSize(1));
            PdfWriter writer = PdfWriter.GetInstance(document, outputPdfStream);

            writer.PageEvent = new ITextSharpAddPdfPageNumberService();

            document.Open();

            for (int i = 1; i <= reader.NumberOfPages; i++)
            {
                document.NewPage();
                PdfImportedPage page = writer.GetImportedPage(reader, i);
                PdfContentByte cb = writer.DirectContent;
                cb.AddTemplate(page, 0, 0);
            }

            document.Close();
            writer.Close();
            reader.Close();

            await System.IO.File.WriteAllBytesAsync(@"C:\Users\TWJOIN\Desktop\Komo\轉換測試檔案\頁碼測試.pdf", outputPdfStream.ToArray());
        }
        return Ok();
    }
}
