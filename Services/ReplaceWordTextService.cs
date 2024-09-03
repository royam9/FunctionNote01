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

    /// <summary>
    /// 插入圖片_v2
    /// </summary>
    /// <param name="worDocx"></param>
    /// <param name="insertPoint">你要插入的定位點</param>
    /// <param name="img"></param>
    /// <returns></returns>
    private async Task InsertImg_2(WordprocessingDocument worDocx, OpenXmlElement insertPoint, IFormFile img)
    {
        //MainDocumentPart mainPart = wordDoc.MainDocumentPart;
        //ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);

        //using (var imagefileStream = imageFile.OpenReadStream())
        //{
        //    imagePart.FeedData(imagefileStream);
        //}

        //string relationshipId = mainPart.GetIdOfPart(imagePart);

        //// Define the reference of the image.
        //var element =
        //     new Drawing(
        //         new DW.Inline(
        //             new DW.Extent() { Cx = 990000L, Cy = 792000L },
        //             new DW.EffectExtent()
        //             {
        //                 LeftEdge = 0L,
        //                 TopEdge = 0L,
        //                 RightEdge = 0L,
        //                 BottomEdge = 0L
        //             },
        //             new DW.DocProperties()
        //             {
        //                 Id = (UInt32Value)1U,
        //                 Name = "Picture 1"
        //             },
        //             new DW.NonVisualGraphicFrameDrawingProperties(
        //                 new A.GraphicFrameLocks() { NoChangeAspect = true }),
        //             new A.Graphic(
        //                 new A.GraphicData(
        //                     new PIC.Picture(
        //                         new PIC.NonVisualPictureProperties(
        //                             new PIC.NonVisualDrawingProperties()
        //                             {
        //                                 Id = (UInt32Value)0U,
        //                                 Name = "New Bitmap Image.jpg"
        //                             },
        //                             new PIC.NonVisualPictureDrawingProperties()),
        //                         new PIC.BlipFill(
        //                             new A.Blip(
        //                                 new A.BlipExtensionList(
        //                                     new A.BlipExtension()
        //                                     {
        //                                         Uri =
        //                                            "{28A0092B-C50C-407E-A947-70E740481C1C}"
        //                                     })
        //                             )
        //                             {
        //                                 Embed = relationshipId,
        //                                 CompressionState =
        //                                 A.BlipCompressionValues.Print
        //                             },
        //                             new A.Stretch(
        //                                 new A.FillRectangle())),
        //                         new PIC.ShapeProperties(
        //                             new A.Transform2D(
        //                                 new A.Offset() { X = 0L, Y = 0L },
        //                                 new A.Extents() { Cx = 990000L, Cy = 792000L }),
        //                             new A.PresetGeometry(
        //                                 new A.AdjustValueList()
        //                             )
        //                             { Preset = A.ShapeTypeValues.Rectangle }))
        //                 )
        //                 { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
        //         )
        //         {
        //             DistanceFromTop = (UInt32Value)0U,
        //             DistanceFromBottom = (UInt32Value)0U,
        //             DistanceFromLeft = (UInt32Value)0U,
        //             DistanceFromRight = (UInt32Value)0U,
        //             EditId = "50D07946"
        //         });

        //bookmarkStart.Parent.AppendChild(new Run(element));
    }

    /// <summary>
    /// Word檔文字插補
    /// </summary>
    /// <param name="wordFilePath">Word檔路徑</param>
    /// <remarks>找尋Paragraph裡面的RunProperties</remarks>
    /// <remark>結論 只是字型的話可以 但是其他的不會包含進去</remark>
    /// <returns></returns>
    public async Task<byte[]> UpdateWordDocument_SearchParagraphProperties(string wordFilePath)
    {
        byte[] templateDocumentBytes = await File.ReadAllBytesAsync(wordFilePath);

        // 將資料寫入內存流
        using (MemoryStream templateStream = new(templateDocumentBytes))
        {
            // 打開這個資料
            using(WordprocessingDocument wordDoc = WordprocessingDocument.Open(templateStream, true))
            {
                var body = wordDoc.MainDocumentPart!.Document.Body;

                //讀取書籤
                IDictionary<String, BookmarkStart> bookmarkMap = new Dictionary<String, BookmarkStart>();

                foreach (BookmarkStart bookmarkStart in wordDoc.MainDocumentPart.RootElement.Descendants<BookmarkStart>())
                {
                    bookmarkMap[bookmarkStart.Name] = bookmarkStart;
                }

                //根據書籤塞入資料
                foreach (BookmarkStart bookmarkStart in bookmarkMap.Values)
                {
                    // 找尋該Paragraph的下一個Run取得字型樣式
                    // 初始化變數來存儲找到的 Run 元素
                    Run nextRun = null;

                    // 使用迴圈遍歷 bookmarkStart 的後續兄弟節點
                    OpenXmlElement currentElement = bookmarkStart.NextSibling();

                    while (currentElement != null)
                    {
                        // 檢查當前元素是否為 Run
                        if (currentElement is Run)
                        {
                            nextRun = (Run)currentElement;
                            break;
                        }

                        // 繼續檢查下一個兄弟節點
                        currentElement = currentElement.NextSibling();
                    }

                    // 預設字型樣式(size:12 字體:讀取Paragraph設定)
                    var nextRunProperties = new RunProperties(new FontSize { Val = "24" });

                    // 取得該Paragraph的RunFonts
                    var ParentParagraph = bookmarkStart.Parent as Paragraph;
                    var isRunFontExist = ParentParagraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<RunFonts>() != null;

                    if (isRunFontExist)
                    {
                        var theRunFonts = ParentParagraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<RunFonts>().CloneNode(true);
                        nextRunProperties.AppendChild(theRunFonts);
                    }

                    // 如果Paragraph內找到下一個Run，使用下一個Run的字型設定
                    if (nextRun != null)
                    {
                        nextRunProperties = (RunProperties)nextRun.RunProperties.CloneNode(true);
                    }

                    #region 替換內容
                    if (bookmarkStart.Name == "CurrentGrade")
                    {
                        bookmarkStart.Parent.InsertAfter(new Run(nextRunProperties, new Text("呀")), bookmarkStart);
                    }

                    if (bookmarkStart.Name == "EnrolledTime")
                    {
                        bookmarkStart.Parent.InsertAfter(new Run(nextRunProperties, new Text("呀")), bookmarkStart);
                    }

                    if (bookmarkStart.Name == "GenderType")
                    {
                        bookmarkStart.Parent.InsertAfter(new Run(nextRunProperties, new Text("呀")), bookmarkStart);
                    }

                    if (bookmarkStart.Name == "RegisterDate")
                    {
                        bookmarkStart.Parent.InsertAfter(new Run(nextRunProperties, new Text("呀")), bookmarkStart);
                    }

                    if (bookmarkStart.Name == "RepresentativeBirthdate")
                    {
                        bookmarkStart.Parent.InsertAfter(new Run(nextRunProperties, new Text("呀")), bookmarkStart);
                    }

                    if (bookmarkStart.Name == "RepresentativeContactAddress")
                    {
                        bookmarkStart.Parent.InsertAfter(new Run(nextRunProperties, new Text("呀")), bookmarkStart);
                    }

                    if (bookmarkStart.Name == "RepresentativeDistrictCode")
                    {
                        bookmarkStart.Parent.InsertAfter(new Run(nextRunProperties, new Text("呀")), bookmarkStart);
                    }

                    if (bookmarkStart.Name == "RepresentativeEmail")
                    {
                        bookmarkStart.Parent.InsertAfter(new Run(nextRunProperties, new Text("呀")), bookmarkStart);
                    }

                    if (bookmarkStart.Name == "RepresentativeMobilePhone")
                    {
                        bookmarkStart.Parent.InsertAfter(new Run(nextRunProperties, new Text("呀")), bookmarkStart);
                    }

                    if (bookmarkStart.Name == "RepresentativeName")
                    {
                        bookmarkStart.Parent.InsertAfter(new Run(nextRunProperties, new Text("呀")), bookmarkStart);
                    }

                    if (bookmarkStart.Name == "RepresentativeRelationshipType")
                    {
                        bookmarkStart.Parent.InsertAfter(new Run(nextRunProperties, new Text("呀")), bookmarkStart);
                    }

                    if (bookmarkStart.Name == "RepresentativeTelephone")
                    {
                        bookmarkStart.Parent.InsertAfter(new Run(nextRunProperties, new Text("呀")), bookmarkStart);
                    }

                    if (bookmarkStart.Name == "StudentBirthDate")
                    {
                        bookmarkStart.Parent.InsertAfter(new Run(nextRunProperties, new Text("呀")), bookmarkStart);
                    }

                    for (var i = 1; i < 5; i++)
                    {
                        if (bookmarkStart.Name == $"StudentClassName{i}")
                        {
                            bookmarkStart.Parent.InsertAfter((new Run(nextRunProperties, new Text("呀"))), bookmarkStart);
                        }
                    }

                    for (var i = 1; i < 5; i++)
                    {
                        if (bookmarkStart.Name == $"StudentName{i}")
                        {
                            bookmarkStart.Parent.InsertAfter(new Run(nextRunProperties, new Text("呀")), bookmarkStart);
                        }
                    }

                    if (bookmarkStart.Name == "StudentSchoolName")
                    {
                        bookmarkStart.Parent.InsertAfter(new Run(nextRunProperties, new Text("呀")), bookmarkStart);
                    }

                    if (bookmarkStart.Name == "TodayDate")
                    {
                        bookmarkStart.Parent.InsertAfter(new Run(nextRunProperties, new Text(DateTime.Now.ToString("yyyy-MM-dd"))), bookmarkStart);
                    }
                    #endregion
                }
                wordDoc.MainDocumentPart.Document.Save();
            }
            return templateStream.ToArray();
        }
    }
}
