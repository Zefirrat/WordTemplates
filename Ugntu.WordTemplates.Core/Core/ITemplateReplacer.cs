using Ugntu.WordTemplates.Core.Core.TemplatesCore;

namespace Ugntu.WordTemplates.Core.Core;

public interface ITemplateReplacer
{
    Task<byte[]> Replace(string templateName, IDictionary<string, string> replaceDictionary);
    string[] GetAvailableTemplates();
    TemplateParameter[] GetParameters(string templateName);
}