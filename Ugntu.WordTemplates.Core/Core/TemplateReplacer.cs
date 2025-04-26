using Ugntu.WordTemplates.Core.Core.Engines;
using Ugntu.WordTemplates.Core.Core.TemplatesCore;

namespace Ugntu.WordTemplates.Core.Core;

public class TemplateReplacer(IDocumentEngine documentEngine) : ITemplateReplacer
{
    protected IList<TemplateBase> Templates = new List<TemplateBase>()
    {
            new ExplanatoryNoteTemplate(documentEngine)
    };

    public async Task<byte[]> Replace(string templateName, IDictionary<string, string> replaceDictionary)
    {
        return await Templates.Single(t => string.Equals(templateName, t.Name)).Replace(replaceDictionary);
    }

    public string[] GetAvailableTemplates()
    {
        return Templates.Select(t => t.Name).ToArray();
    }

    public TemplateParameter[] GetParameters(string templateName)
    {
        return Templates.Single(t => t.Name == templateName).GetAvailableParameters();
    }
}