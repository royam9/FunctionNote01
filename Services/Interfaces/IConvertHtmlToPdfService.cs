namespace Services.Interfaces;

public interface IConvertHtmlToPdfService
{
    /// <summary>
    /// Html 轉換成 Pdf
    /// </summary>
    /// <param name="htmlContent">Html文字(UTF-8)</param>
    /// <return></return>
    byte[] ConvertUTF8HtmlToPdf(string htmlContent);


    /// <summary>
    /// Big5 Html 轉換成 Pdf
    /// </summary>
    /// <param name="htmlContent">Html文字(Big5)</param>
    /// <return></return>
    byte[] ConvertBig5HtmlToPdf(string htmlContent);

    /// <summary>
    /// 非同步 Html 轉換成 Pdf
    /// </summary>
    /// <param name="htmlContent">Html文字(UTF-8)</param>
    /// <return></return>
    Task<byte[]> HtmlToPdf(string htmlContent);

    /// <summary>
    /// 替換Word另存新檔成html 裡面的參數
    /// </summary>
    /// <param name="htmlContent">參數</param>
    /// <remarks>參數樣式 {參數名稱}</remarks>
    /// <returns></returns>
    string ReplaceString(string htmlContent);
}
