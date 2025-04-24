using IronWord;
using Microsoft.Office.Interop.Word;
using Ugntu.WordTemplates.Core.Engines;

namespace Ugntu.WordTemplates.Core;

public class TemplateReplacer(IDocumentEngine documentEngine) : ITemplateReplacer
{
    protected IList<TemplateBase> Templates = new List<TemplateBase>()
    {
            new ExplanatoryNoteTemplate(documentEngine)
    };

    public byte[] Replace(string templateName, IDictionary<string, string> replaceDictionary)
    {
        throw new NotImplementedException();
    }

    public string[] GetAvailableTemplates(string templateName)
    {
        return Templates.Select(t => t.Name).ToArray();
    }
}