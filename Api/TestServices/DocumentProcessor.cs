using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using GemBox.Document;
using iTextSharp.text.pdf;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace Api.TestServices;

public class DocumentProcessor
{
    private const int ParagraphsPerDocument = 20;

    public DocumentProcessor()
    {
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");
    }

    public void ProcessDocument(string inputWordPath, string outputPdfPath)
    {
        // 步驟 1: 分割 Word 文檔
        var wordStreams = SplitWordDocument(inputWordPath);

        // 步驟 2: 轉換成 PDF
        var pdfStreams = ConvertWordToPdf(wordStreams);

        // 步驟 3: 合併 PDF
        MergePdfs(pdfStreams, outputPdfPath);

        // 清理臨時流
        foreach (var stream in wordStreams.Concat(pdfStreams))
        {
            stream.Dispose();
        }
    }

    private List<MemoryStream> SplitWordDocument(string inputPath)
    {
        var outputStreams = new List<MemoryStream>();

        using (var wordDoc = WordprocessingDocument.Open(inputPath, false))
        {
            var body = wordDoc.MainDocumentPart.Document.Body;
            var paragraphs = new List<DocumentFormat.OpenXml.Wordprocessing.Paragraph>();

            foreach (var paragraph in body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>())
            {
                paragraphs.Add((DocumentFormat.OpenXml.Wordprocessing.Paragraph)paragraph.CloneNode(true));

                if (paragraphs.Count == ParagraphsPerDocument)
                {
                    outputStreams.Add(CreateWordDocument(paragraphs));
                    paragraphs.Clear();
                }
            }

            if (paragraphs.Count > 0)
            {
                outputStreams.Add(CreateWordDocument(paragraphs));
            }
        }

        return outputStreams;
    }

    private MemoryStream CreateWordDocument(List<DocumentFormat.OpenXml.Wordprocessing.Paragraph> paragraphs)
    {
        var stream = new MemoryStream();
        using (var newDoc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var mainPart = newDoc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            foreach (var paragraph in paragraphs)
            {
                mainPart.Document.Body.AppendChild(paragraph.CloneNode(true));
            }

            mainPart.Document.Save();
        }
        stream.Position = 0;
        return stream;
    }

    private List<MemoryStream> ConvertWordToPdf(List<MemoryStream> wordStreams)
    {
        var pdfStreams = new List<MemoryStream>();

        foreach (var wordStream in wordStreams)
        {
            var pdfStream = new MemoryStream();
            var document = DocumentModel.Load(wordStream, GemBox.Document.LoadOptions.DocxDefault);
            document.Save(pdfStream, GemBox.Document.SaveOptions.PdfDefault);
            pdfStream.Position = 0;
            pdfStreams.Add(pdfStream);
        }

        return pdfStreams;
    }

    private void MergePdfs(List<MemoryStream> pdfStreams, string outputPath)
    {
        using (var outputDocument = new PdfSharp.Pdf.PdfDocument())
        {
            foreach (var pdfStream in pdfStreams)
            {
                using (var inputDocument = PdfSharp.Pdf.IO.PdfReader.Open(pdfStream, PdfDocumentOpenMode.Import))
                {
                    for (int i = 0; i < inputDocument.PageCount; i++)
                    {
                        outputDocument.AddPage(inputDocument.Pages[i]);
                    }
                }
            }

            outputDocument.Save(outputPath);
        }
    }
}
