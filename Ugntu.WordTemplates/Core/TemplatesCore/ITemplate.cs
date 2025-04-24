namespace Ugntu.WordTemplates.Core;

public interface ITemplate
{
    byte[] Replace(IDictionary<string, string> replaceDictionary);
    TemplateParameter[] GetAvailableParameters();
}