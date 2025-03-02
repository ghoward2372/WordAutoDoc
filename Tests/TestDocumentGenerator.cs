using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;

namespace DocumentProcessor.Tests
{
    public class TestDocumentGenerator
    {
        public static void CreateTestDocument(string filePath)
        {
            try
            {
                using (var document = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
                {
                    // Add a main document part
                    var mainPart = document.AddMainDocumentPart();
                    mainPart.Document = new Document();
                    var body = mainPart.Document.AppendChild(new Body());

                    // Add test content with various tags
                    AddParagraph(body, "Test Document for Document Processor");
                    AddParagraph(body, string.Empty);
                    AddParagraph(body, "Grammar Test Paragraph:");
                    AddParagraph(body, "The cats and dog is running fast. We dont need no help with grammer. This sentense contains muliple mispelled words. The weather have been nice yesterday?");
                    AddParagraph(body, string.Empty);
                    AddParagraph(body, "Work Item Example:");
                    AddParagraph(body, "[[WorkItem:1234]]");
                    AddParagraph(body, string.Empty);
                    AddParagraph(body, "Query Results Example:");
                    AddParagraph(body, "[[QueryID:12345678-1234-1234-1234-123456789ABC]]");
                    AddParagraph(body, string.Empty);
                    AddParagraph(body, "Acronym Examples:");
                    AddParagraph(body, "The Application Programming Interface (API) is used for integration.");
                    AddParagraph(body, "The Graphical User Interface (GUI) provides user interaction.");
                    AddParagraph(body, string.Empty);
                    AddParagraph(body, "Acronym Table:");
                    AddParagraph(body, "[[AcronymTable:true]]");

                    mainPart.Document.Save();
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error creating test document: {ex.Message}", ex);
            }
        }

        private static void AddParagraph(Body body, string text)
        {
            var para = body.AppendChild(new Paragraph());
            var run = para.AppendChild(new Run());
            run.AppendChild(new Text(text));
        }
    }
}