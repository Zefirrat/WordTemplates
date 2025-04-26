using Microsoft.AspNetCore.Mvc;
using Swashbuckle.AspNetCore.Annotations;
using Ugntu.WordTemplates.Core.Core;
using Ugntu.WordTemplates.Core.Core.TemplatesCore;

namespace Ugntu.WordTemplates.Api.Controllers;

[ApiController]
[Route("[controller]")]
public class Templates(ITemplateReplacer templateReplacer) : Controller
{
    [HttpGet("/")]
    [ProducesResponseType<string[]>(StatusCodes.Status200OK)]
    public IActionResult GetTemplates()
    {
        return Ok(templateReplacer.GetAvailableTemplates());
    }

    [HttpGet("/{templateName}/parameters")]
    [ProducesResponseType<TemplateParameter[]>(StatusCodes.Status200OK)]
    public IActionResult GetParameters(string templateName)
    {
        return Ok(templateReplacer.GetParameters(templateName));
    }

    [HttpPost("/{templateName}/replace")]
    [ProducesResponseType<FileResult>(StatusCodes.Status200OK)]
    [SwaggerResponse(StatusCodes.Status200OK, "File download", contentTypes:["application/octet-stream"])]
    public async Task<IActionResult> Replace(string templateName,
        [FromBody] IDictionary<string, string> replaceDictionary)
    {
        return File(await templateReplacer.Replace(templateName, replaceDictionary), "application/octet-stream",
            $"{templateName}.{DateTime.Now:yymmddhhMMss}.docx");
    }
}