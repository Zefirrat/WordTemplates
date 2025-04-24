using Microsoft.Office.Interop.Word;

namespace Ugntu.WordTemplates.Core;

public abstract class TemplateBase(string Name, string FileName) : ITemplate
{
    public string Name { get; }
    public string FileName { get; }

    public abstract IEnumerable<TemplateParameter> TemplateParameters { get; }

    public byte[] Replace(IDictionary<string, string> replaceDictionary)
    {
        var filePath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Templates", FileName);


        return Array.Empty<byte>();
    }

    public TemplateParameter[] GetAvailableParameters()
    {
        return TemplateParameters.ToArray();
    }
}