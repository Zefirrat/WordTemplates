using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdig;
using HtmlToOpenXml;
using Ugntu.WordTemplates.Core.Core.Engines;

namespace Ugntu.WordTemplates.Core.Engines
{
    public class OpenXmlEngine : IDocumentEngine
    {
        public Task<bool> Replace(string templateFilePath, string finalFileName, IDictionary<string,string> parameters)
        {
            var outputDir = Path.Combine(Directory.GetCurrentDirectory(), "TemplateOutput");
            Directory.CreateDirectory(outputDir);
            var outPath = Path.Combine(outputDir,
                Path.GetFileNameWithoutExtension(finalFileName) +
                $"_{DateTime.Now:yyyyMMddHHmmss}.docx");

            // Копируем шаблон, чтобы не портить оригинал
            File.Copy(templateFilePath, outPath, true);

            using (var doc = WordprocessingDocument.Open(outPath, true))
            {
                var mainPart = doc.MainDocumentPart;
                var body = mainPart.Document.Body;

                foreach (var kvp in parameters)
                {
                    string tag = $"#{kvp.Key}";

                    if (kvp.Key.StartsWith("markdown", StringComparison.OrdinalIgnoreCase))
                    {
                        // Конвертируем Markdown→HTML
                        var html = Markdown.ToHtml(kvp.Value);

                        // Ищем параграф с маркером
                        var paras = body.Descendants<Text>()
                                        .Where(t => t.Text.Contains(tag))
                                        .Select(t => t.Parent as Paragraph)
                                        .Where(p => p != null)
                                        .ToList();

                        foreach (var p in paras)
                        {
                            // удаляем маркер
                            var text = p.Descendants<Text>().First();
                            text.Text = text.Text.Replace(tag, "");

                            // вставляем HTML как Word-параграфы
                            var converter = new HtmlConverter(mainPart);
                            var newBlocks = converter.Parse(html);
                            foreach (var block in newBlocks)
                                p.InsertAfterSelf(block);
                        }
                    }
                    else
                    {
                        // Простая замена текста во всём документе
                        var texts = body.Descendants<Text>()
                                        .Where(t => t.Text.Contains(tag));
                        foreach (var t in texts)
                            t.Text = t.Text.Replace(tag, kvp.Value);
                    }
                }

                mainPart.Document.Save();
            }

            return Task.FromResult(true);
        }
    }
}
