namespace Ugntu.WordTemplates.Core;

public interface IWordReplacer
{
    byte[] Replace(string fileName, IDictionary<string, string> replaceDictionary);
}