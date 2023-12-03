using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using System.Threading.Tasks;
using System.Text.Json;
using DocumentFormat.OpenXml;

namespace DocTextConverter
{
    public static class ConvertFunction
    {
        [FunctionName("ConvertFunction")]
        public static async Task<IActionResult> Run(
    [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = null)] HttpRequest req,
    ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            try
            {
                string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
                JsonDocument data = JsonDocument.Parse(requestBody);

                string conversionType = data.RootElement.GetProperty("conversionType").GetString();
                string content = data.RootElement.GetProperty("content").GetString();

                if (conversionType == "docToText")
                {
                    string base64EncodedDoc = data.RootElement.GetProperty("content").GetString();
                    byte[] fileBytes = Convert.FromBase64String(base64EncodedDoc);
                    using (var memoryStream = new MemoryStream(fileBytes))
                    {
                        string textResult = ConvertDocToText(memoryStream);
                        return new OkObjectResult(textResult);
                    }
                }
                else if (conversionType == "textToDoc")
                {
                    string plainTextContent = data.RootElement.GetProperty("content").GetString();
                    byte[] docResult = ConvertTextToDoc(plainTextContent);
                    return new FileContentResult(docResult, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                    {
                        FileDownloadName = "converted.docx"
                    };
                }
                else
                {
                    log.LogError("Invalid conversion type.");
                    return new BadRequestObjectResult("Invalid conversion type.");
                }
            }
            catch (Exception ex)
            {
                log.LogError($"An error occurred: {ex.Message}");
                return new BadRequestObjectResult($"An error occurred: {ex.Message}");
            }
        }
        private static string ConvertDocToText(MemoryStream docStream)
        {
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(docStream, false))
            {
                StringBuilder textBuilder = new StringBuilder();
                Body body = wordDocument.MainDocumentPart.Document.Body;
                foreach (var para in body.Elements<Paragraph>())
                {
                    foreach (var run in para.Elements<Run>())
                    {
                        foreach (var text in run.Elements<Text>())
                        {
                            if (text.Space == SpaceProcessingModeValues.Preserve)
                            {
                                // Preserve spaces explicitly if set
                                textBuilder.Append(text.Text);
                            }
                            else
                            {
                                // For text without the preserve attribute, spaces are typically preserved by default
                                textBuilder.Append(text.Text);
                            }
                        }
                    }
                    // Add a line break after each paragraph to maintain paragraph separation
                    textBuilder.AppendLine();
                }
                return textBuilder.ToString().TrimEnd(); // TrimEnd to remove any trailing new lines
            }
        }

        private static byte[] ConvertTextToDoc(string textContent)
        {
            using (var outputStream = new MemoryStream())
            {
                using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(outputStream, WordprocessingDocumentType.Document, true))
                {
                    MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                    mainPart.Document = new Document(new Body());
                    AddFormattedContent(mainPart, textContent);
                    mainPart.Document.Save();
                }
                return outputStream.ToArray();
            }
        }

        private static void AddFormattedContent(MainDocumentPart mainPart, string decodedText)
        {
            Body body = mainPart.Document.Body;

            // Split the text into paragraphs using \n\n as the delimiter
            string[] paragraphs = decodedText.Split(new string[] { "\n\n" }, StringSplitOptions.None);

            foreach (string paragraphText in paragraphs)
            {
                Paragraph para = new Paragraph();

                if (!string.IsNullOrWhiteSpace(paragraphText))
                {
                    string[] lines = paragraphText.Split(new string[] { "\n" }, StringSplitOptions.None);

                    foreach (string line in lines)
                    {
                        ProcessLine(para, line);
                    }
                }

                body.Append(para); // Append the paragraph to the body
            }
        }

        private static void ProcessLine(Paragraph paragraph, string line)
        {
            Regex quoteRegex = new Regex("\"([^\"]*)\"");
            int lastIndex = 0;

            foreach (Match match in quoteRegex.Matches(line))
            {
                AppendTextToParagraph(paragraph, line.Substring(lastIndex, match.Index - lastIndex), false);
                AppendTextToParagraph(paragraph, match.Value, true);
                lastIndex = match.Index + match.Length;
            }

            AppendTextToParagraph(paragraph, line.Substring(lastIndex), false);
        }


        private static void AppendTextToParagraph(Paragraph paragraph, string text, bool italic)
        {
            Run run = new Run();
            RunProperties runProps = new RunProperties(new RunFonts() { Ascii = "Palatino Linotype" }, new FontSize() { Val = "22" }); // Font size 11

            if (italic)
            {
                runProps.Append(new Italic());
            }

            run.Append(runProps);
            run.Append(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
            paragraph.Append(run);
        }
    }
}