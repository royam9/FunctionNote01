using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace Api.TestServices;

public class AddPdfPageNumService
{
    public static byte[] AddPageNumbers(byte[] pdfBytes)
    {
        using (MemoryStream inputStream = new MemoryStream(pdfBytes))
        using (PdfDocument document = PdfReader.Open(inputStream, PdfDocumentOpenMode.Modify))
        {
            int count = document.PageCount;

            for (int idx = 0; idx < count; idx++)
            {
                PdfPage page = document.Pages[idx];
                // 繪圖功能
                XGraphics gfx = XGraphics.FromPdfPage(page);
                // 設定字型 字體大小
                XFont font = new XFont("Arial", 12);

                string pageNumber = $"{idx + 1}";
                // 計算字體大小
                XSize size = gfx.MeasureString(pageNumber, font);

                // 設定座標
                // 使用 XUnit 將cm轉換成point
                XUnit margin = XUnit.FromCentimeter(1.5);

                double x = (page.Width.Point - size.Width) / 2;
                double y = page.Height.Point - margin.Point;

                gfx.DrawString(pageNumber,
                    font,
                    XBrushes.Black, // 顏色
                    new XRect(x, y, size.Width, size.Height), //設定矩形座標、大小
                    XStringFormats.TopLeft); // 文本在矩形內的對齊方式
            }

            // Save到新的MemoryStream
            using (MemoryStream outputStream = new MemoryStream())
            {
                document.Save(outputStream, false);
                return outputStream.ToArray();
            }
        }
    }
}
