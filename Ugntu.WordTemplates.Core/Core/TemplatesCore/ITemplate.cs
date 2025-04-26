namespace Ugntu.WordTemplates.Core.Core.TemplatesCore;

public interface ITemplate
{
    Task<byte[]> Replace(IDictionary<string, string> replaceDictionary);
    TemplateParameter[] GetAvailableParameters();
}