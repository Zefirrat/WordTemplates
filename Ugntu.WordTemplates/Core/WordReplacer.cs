using IronWord;

namespace Ugntu.WordTemplates.Core;

public class WordReplacer : IWordReplacer
{
    public byte[] Replace(string fileName, IDictionary<string, string> replaceDictionary)
    {
        var filePath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Templates", fileName);
        
        var WordApp = new Microsoft.Office.Interop.Word.Application();
        var WordDoc = WordApp.Documents.Open(filePath);
        
        FindAndReplace(WordApp, replaceDictionary.First().Key, replaceDictionary.First().Value);
        WordDoc.SaveAs(System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "TeplateOutput",
            fileName.Replace(".template", "")));

        return Array.Empty<byte>();
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
        doc.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
            ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, ref replace,
            ref matchKashida ,ref matchDiacritics, ref matchAlefHamza, ref matchControl);                
    }
}