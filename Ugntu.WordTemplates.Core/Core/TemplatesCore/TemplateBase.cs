using Ugntu.WordTemplates.Core.Core.Engines;

namespace Ugntu.WordTemplates.Core.Core.TemplatesCore;

public abstract class TemplateBase(string name, string fileName, IDocumentEngine documentEngine) : ITemplate
{
    public string Name { get; } = name;
    public string FileName { get; } = fileName;

    public abstract IEnumerable<TemplateParameter> TemplateParameters { get; }

    public async Task<byte[]> Replace(IDictionary<string, string> replaceDictionary)
    {
        var filePath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Templates", FileName);

        bool success;
        try
        {
            success = await documentEngine.Replace(filePath, FileName, replaceDictionary);
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