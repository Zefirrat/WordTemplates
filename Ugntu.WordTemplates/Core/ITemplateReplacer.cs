namespace Ugntu.WordTemplates.Core;

public interface ITemplateReplacer
{
    byte[] Replace(string templateName, IDictionary<string, string> replaceDictionary);
    string[] GetAvailableTemplates(string templateName);
}