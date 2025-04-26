namespace Ugntu.WordTemplates.Core.Core.Engines;

public interface IDocumentEngine
{
    Task<bool> Replace(string templateFilePath, string finalFileName, IDictionary<string, string> parameters);
}