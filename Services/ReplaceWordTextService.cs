using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using Microsoft.AspNetCore.Http;
using Models.ApiModels;
using Services.Interfaces;
using SixLabors.ImageSharp.PixelFormats;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Models.ApiModels;
using Microsoft.AspNetCore.Http;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using SixLabors.ImageSharp.PixelFormats;
using SixLabors.ImageSharp;

namespace Services;

public class ReplaceWordTextService : IReplaceWordTextService
{
    /// <summary>
    /// 編輯Word襠
    /// </summary>
    /// <param name="param">參數</param>
    /// <returns></returns>
    public async Task<byte[]> UpdateWordDocument(WordDocumentRequestModel param)
    {
        // 獲取Word檔案的路徑
        var templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "WordDocument.docx");
        var outputPath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.docx");

        // 複製檔案到臨時檔案
        File.Copy(templatePath, outputPath, true);

        // 使用OpenXml來編輯Word檔
        using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(outputPath, true))
        {
            var body = wordDoc.MainDocumentPart!.Document.Body;

            // 替換文字佔位符
            ReplacePlaceholderText(body, "{學生姓名}", param.StudentName);
            ReplacePlaceholderText(body, "{教室名稱}", param.ClassName);

            // 插入圖片
            if (param.Images != null)
            {
                // 找到{圖片}佔位符所在的Run
                var placeholder = body.Descendants<Run>()
                    .FirstOrDefault(r => r.InnerText.Contains("{圖片}"));

                if (placeholder != null)
                {
                    // 替換佔位符為圖片
                    InsertImage(wordDoc, placeholder, param.Images);
                }
            }

            wordDoc.MainDocumentPart.Document.Save();
        }

        return await File.ReadAllBytesAsync(outputPath);
    }


    /// <summary>
    /// 插入文字到指定位置
    /// </summary>
    /// <param name="body">主文件</param>
    /// <param name="placeholder">佔位符文字</param>
    /// <param name="text">文字內容</param>
    private void ReplacePlaceholderText(Body body, string placeholder, string? text)
    {
        foreach (var textElement in body.Descendants<Text>())
        {
            if (textElement.Text.Contains(placeholder))
            {
                textElement.Text = textElement.Text.Replace(placeholder, text ?? string.Empty);
            }
        }
    }

    /// <summary>
    /// 插入圖片到指定位置
    /// </summary>
    /// <param name="wordDoc">主文件</param>
    /// <param name="placeholder">佔位符文字</param>
    /// <param name="imageFiles">圖片檔案</param>
    /// <returns></returns>
    private void InsertImage(WordprocessingDocument wordDoc, Run placeholder, IEnumerable<IFormFile> imageFiles)
    {
        foreach (var imageFile in imageFiles)
        {
            var imageType = imageFile.ContentType;
            var imagePart = wordDoc.MainDocumentPart!.AddNewPart<ImagePart>(imageType);

            // 獲取圖片流
            using (var stream = imageFile.OpenReadStream())
            {
                // 加載圖片
                using (var image = Image.Load<Rgba32>(stream))
                {
                    // 將流位置重置到開頭
                    stream.Position = 0;

                    // 將圖片數據寫入 ImagePart
                    imagePart.FeedData(stream);

                    // 計算圖片尺寸（以 EMUs 為單位）
                    long widthEmus = image.Width * 9525; // 1 px = 9525 EMUs
                    long heightEmus = image.Height * 9525;

                    // 獲取圖片 ID
                    string imageId = wordDoc.MainDocumentPart.GetIdOfPart(imagePart);

                    var element = new Drawing(
                        new DW.Inline(
                            new DW.Extent() { Cx = widthEmus, Cy = heightEmus },
                            new DW.DocProperties() { Id = (UInt32Value)1U, Name = imageFile.FileName },
                            new DW.NonVisualGraphicFrameDrawingProperties(
                                new A.GraphicFrameLocks() { NoChangeAspect = true }
                            ),
                            new A.Graphic(
                                new A.GraphicData(
                                    new Pic.Picture(
                                        new Pic.NonVisualPictureProperties(
                                            new Pic.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = imageFile.FileName },
                                            new Pic.NonVisualPictureDrawingProperties()
                                        ),
                                        new Pic.BlipFill(
                                            new A.Blip() { Embed = imageId },
                                            new A.Stretch(new A.FillRectangle())
                                        ),
                                        new Pic.ShapeProperties(
                                            new A.Transform2D(
                                                new A.Offset() { X = 0L, Y = 0L },
                                                new A.Extents() { Cx = widthEmus, Cy = heightEmus }
                                            ),
                                            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                                        )
                                    )
                                )
                                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                            )
                        )
                    );

                    placeholder.Parent!.InsertAfter(element, placeholder);
                }
            }
        }

        placeholder.Remove();
    }
}
