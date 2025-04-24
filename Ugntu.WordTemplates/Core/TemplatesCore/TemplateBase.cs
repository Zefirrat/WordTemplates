using Microsoft.Office.Interop.Word;
using Ugntu.WordTemplates.Core.Engines;

namespace Ugntu.WordTemplates.Core;

public abstract class TemplateBase(string Name, string FileName, IDocumentEngine documentEngine) : ITemplate
{
    public string Name { get; }
    public string FileName { get; }

    public abstract IEnumerable<TemplateParameter> TemplateParameters { get; }

    public byte[] Replace(IDictionary<string, string> replaceDictionary)
    {
        var filePath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Templates", FileName);

        bool success;
        try
        {
            success = documentEngine.Replace(filePath, FileName, replaceDictionary);
        }
        catch (Exception e)
        {
            Console.WriteLine(e);
            throw new Exception("Ошибка при формировании файла.", e);
        }

        if (!success)
            throw new Exception("Проблема при формировании файла.");

        return File.ReadAllBytes(filePath);
    }

    public TemplateParameter[] GetAvailableParameters()
    {
        return TemplateParameters.ToArray();
    }
}