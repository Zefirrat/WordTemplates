using Microsoft.Office.Interop.Word;
using Task = System.Threading.Tasks.Task;

namespace Ugntu.WordTemplates.Core.Core.Engines;

public class WordEngine : IDocumentEngine
{
    public Task<bool> Replace(string templateFilePath, string finalFileName, IDictionary<string, string> parameters)
    {
        Application WordApp = null;
        Document? WordDoc = null;
        try
        {
            WordApp = new Application();
            WordDoc = WordApp.Documents.Open(templateFilePath);
            var markdownToWordConverter = new MarkdownToWordConverter();

            foreach (var replacedWord in parameters)
            {
                if (replacedWord.Key.StartsWith("markdown"))
                {
                    markdownToWordConverter.ConvertMarkdownToWord(
                            replacedWord.Value,
                            WordApp,
                            WordDoc,
                            replacedWord.Key);
                }
                else
                {
                    FindAndReplace(WordApp, $"#{replacedWord.Key}", replacedWord.Value);
                }
            }


            WordDoc.SaveAs2(
                    System.IO.Path.Combine(
                            System.IO.Directory.GetCurrentDirectory(),
                            "TemplateOutput",
                            finalFileName.Replace(".template", "").Replace(
                                    ".doc",
                                    $"{DateTime.Now:yyyyMMddHHmmss}.doc")));
        }
        finally
        {
            WordDoc?.Close();
            WordApp?.Quit();
        }

        return Task.FromResult(true);
    }

    private void FindAndReplace(Microsoft.Office.Interop.Word.Application doc, object findText, object replaceWithText)
    {
        //options
        object matchCase = false;
        object matchWholeWord = true;
        object matchWildCards = false;
        object matchSoundsLike = false;
        object matchAllWordForms = false;
        object forward = true;
        object format = false;
        object matchKashida = false;
        object matchDiacritics = false;
        object matchAlefHamza = false;
        object matchControl = false;
        object read_only = false;
        object visible = true;
        object replace = 2;
        object wrap = 1;
        //execute find and replace
        doc.Selection.Find.Execute(
                ref findText,
                ref matchCase,
                ref matchWholeWord,
                ref matchWildCards,
                ref matchSoundsLike,
                ref matchAllWordForms,
                ref forward,
                ref wrap,
                ref format,
                ref replaceWithText,
                ref replace,
                ref matchKashida,
                ref matchDiacritics,
                ref matchAlefHamza,
                ref matchControl);
    }
}