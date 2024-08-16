using System.Xml.Linq;

namespace Service;

public class ConvertWordToHtmlService
{
    //public static string ConvertWordToHtml(string wordFilePath, string htmlFilePath)
    //{
    //    // 要安裝OpenXmlPowerTools

    //    byte[] byteArray = File.ReadAllBytes(wordFilePath);
    //    using (MemoryStream memoryStream = new MemoryStream())
    //    {
    //        memoryStream.Write(byteArray, 0, byteArray.Length);
    //        using (WordprocessingDocument doc = WordprocessingDocument.Open(memoryStream, true))
    //        {
    //            HtmlConverterSettings settings = new HtmlConverterSettings()
    //            {
    //                PageTitle = "Converted Document"
    //            };
    //            XElement html = HtmlConverter.ConvertToHtml(doc, settings);

    //            File.WriteAllText(htmlFilePath, html.ToStringNewLineOnAttributes());
    //            return html.ToStringNewLineOnAttributes();
    //        }
    //    }
    //}
}
