using iTextSharp.text;
using iTextSharp.text.pdf;

namespace Api.TestServices;

public class ITextSharpAddPdfPageNumberService : PdfPageEventHelper
{
    #region 測試1
    // This is the contentbyte object of the wirter
    PdfContentByte cb;

    // we will put the fianl number of pages in template
    //PdfTemplate template;

    // this is the BaseFont we are going to use for the header / footer
    BaseFont bf = null;

    // This keeps track of the creation time
    //DateTime PrintTime = DateTime.Now;

    public override void OnOpenDocument(PdfWriter writer, Document document)
    {
        try
        {
            //PrintTime = DateTime.Now;
            bf = BaseFont.CreateFont(BaseFont.HELVETICA,
                BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb = writer.DirectContent;
            //template = cb.CreateTemplate(50, 50);
        }
        catch (DocumentException de) { }
        catch (System.IO.IOException ioe) { }
    }

    // 當頁面結束
    public override void OnEndPage(PdfWriter writer, Document document)
    {
        base.OnEndPage(writer, document);

        // 當前是第幾頁
        int pageN = writer.PageNumber;
        // 要輸入的內容
        string text = pageN.ToString();
        // 輸入內容的長度
        float len = bf.GetWidthPoint(text, 12);
        // 文件頁面的大小
        Rectangle pageSize = document.PageSize;
        // 取的頁面中心點的座標(用於置中)
        float centerX = (pageSize.Left + pageSize.Right) / 2;

        // 設定顏色
        cb.SetRgbColorFill(0, 0, 0);
        // 開始寫字
        cb.BeginText();
        // 設定字體 和 字體大小
        cb.SetFontAndSize(bf, 12);

        //cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT,
        //    "Test",
        //    pageSize.GetRight(40),
        //    pageSize.GetTop(30),
        //    0);

        // ALIGN_CENTER 代表從文本以x座標為中心向左右生成
        cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER,
            text,
            centerX,
            pageSize.GetBottom(30),
            0);

        // SetTextMatrix 另一種生成方式
        //cb.SetTextMatrix(pageSize.GetLeft(80),
        //    pageSize.GetBottom(30));
        //cb.ShowText(text);
        cb.EndText();

        // 生成時間的，之後再研究
        //cb.AddTemplate(template, pageSize.GetLeft(40) + len,
        //    pageSize.GetBottom(30));
        //cb.BeginText();
        //cb.SetFontAndSize(bf, 50);
        //cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT,
        //    "Printed On" + PrintTime.ToString(),
        //    pageSize.GetRight(40),
        //    pageSize.GetBottom(40),
        //    0);
        //cb.EndText();
    }
    #endregion
    #region SlackFlow
    //// This is the contentbyte object of the writer
    //PdfContentByte cb;
    //// we will put the final number of pages in a template
    //PdfTemplate template;
    //// this is the BaseFont we are going to use for the header / footer
    //BaseFont bf = BaseFont.CreateFont(@"C:\Windows\Fonts\kaiu.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
    //// This keeps track of the creation time
    //DateTime PrintTime = DateTime.Now;
    //#region Properties
    //private string _Title;
    //public string Title
    //{
    //    get { return _Title; }
    //    set { _Title = value; }
    //}

    //private string _HeaderLeft;
    //public string HeaderLeft
    //{
    //    get { return _HeaderLeft; }
    //    set { _HeaderLeft = value; }
    //}
    //private string _HeaderRight;
    //public string HeaderRight
    //{
    //    get { return _HeaderRight; }
    //    set { _HeaderRight = value; }
    //}
    //private Font _HeaderFont;
    //public Font HeaderFont
    //{
    //    get { return _HeaderFont; }
    //    set { _HeaderFont = value; }
    //}
    //private Font _FooterFont;
    //public Font FooterFont
    //{
    //    get { return _FooterFont; }
    //    set { _FooterFont = value; }
    //}
    //#endregion
    //// we override the onOpenDocument method
    //public override void OnOpenDocument(PdfWriter writer, Document document)
    //{
    //    try
    //    {
    //        PrintTime = DateTime.Now;
    //        bf = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
    //        cb = writer.DirectContent;
    //        template = cb.CreateTemplate(50, 50);
    //    }
    //    catch (DocumentException de)
    //    {
    //    }
    //    catch (System.IO.IOException ioe)
    //    {
    //    }
    //}

    //public override void OnStartPage(PdfWriter writer, Document document)
    //{
    //    base.OnStartPage(writer, document);
    //    Rectangle pageSize = document.PageSize;
    //    if (Title != string.Empty)
    //    {
    //        cb.BeginText();
    //        cb.SetFontAndSize(bf, 15);
    //        cb.SetRgbColorFill(0, 0, 0);
    //        cb.SetTextMatrix(pageSize.GetLeft(40), pageSize.GetTop(40));
    //        cb.ShowText(Title);
    //        cb.EndText();
    //    }
    //    if (HeaderLeft + HeaderRight != string.Empty)
    //    {
    //        PdfPTable HeaderTable = new PdfPTable(2);
    //        HeaderTable.DefaultCell.VerticalAlignment = Element.ALIGN_MIDDLE;
    //        HeaderTable.TotalWidth = pageSize.Width - 80;
    //        HeaderTable.SetWidthPercentage(new float[] { 45, 45 }, pageSize);

    //        PdfPCell HeaderLeftCell = new PdfPCell(new Phrase(8, HeaderLeft, HeaderFont));
    //        HeaderLeftCell.Padding = 5;
    //        HeaderLeftCell.PaddingBottom = 8;
    //        HeaderLeftCell.BorderWidthRight = 0;
    //        HeaderTable.AddCell(HeaderLeftCell);
    //        PdfPCell HeaderRightCell = new PdfPCell(new Phrase(8, HeaderRight, HeaderFont));
    //        HeaderRightCell.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;
    //        HeaderRightCell.Padding = 5;
    //        HeaderRightCell.PaddingBottom = 8;
    //        HeaderRightCell.BorderWidthLeft = 0;
    //        HeaderTable.AddCell(HeaderRightCell);
    //        cb.SetRgbColorFill(0, 0, 0);
    //        HeaderTable.WriteSelectedRows(0, -1, pageSize.GetLeft(40), pageSize.GetTop(50), cb);
    //    }
    //}
    //public override void OnEndPage(PdfWriter writer, Document document)
    //{
    //    base.OnEndPage(writer, document);
    //    int pageN = writer.PageNumber;
    //    String text = "Page " + pageN + " of ";
    //    float len = bf.GetWidthPoint(text, 8);
    //    Rectangle pageSize = document.PageSize;
    //    cb.SetRgbColorFill(0, 0, 0);
    //    cb.BeginText();
    //    cb.SetFontAndSize(bf, 100);
    //    cb.SetTextMatrix(pageSize.GetLeft(40), pageSize.GetBottom(30));
    //    cb.ShowText(text);
    //    cb.EndText();
    //    cb.AddTemplate(template, pageSize.GetLeft(40) + len, pageSize.GetBottom(30));

    //    cb.BeginText();
    //    cb.SetFontAndSize(bf, 8);
    //    cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT,
    //    "Printed On " + PrintTime.ToString(),
    //    pageSize.GetRight(40),
    //    pageSize.GetBottom(30), 0);
    //    cb.EndText();
    //}
    //public override void OnCloseDocument(PdfWriter writer, Document document)
    //{
    //    base.OnCloseDocument(writer, document);
    //    template.BeginText();
    //    template.SetFontAndSize(bf, 100);
    //    template.SetTextMatrix(0, 0);
    //    template.ShowText("" + (writer.PageNumber - 1));
    //    template.EndText();
    //}
    #endregion
}
