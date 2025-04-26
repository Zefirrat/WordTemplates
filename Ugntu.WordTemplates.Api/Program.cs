using Ugntu.WordTemplates.Core.Core;
using Ugntu.WordTemplates.Core.Core.Engines;
using Ugntu.WordTemplates.Core.Engines;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen(c => c.EnableAnnotations());
builder.Services.AddControllers();

builder.Services.AddScoped<ITemplateReplacer, TemplateReplacer>();
builder.Services.AddScoped<IDocumentEngine, OpenXmlEngine>();

var app = builder.Build();

app.UseSwagger();
app.UseSwaggerUI();

app.MapControllers();

app.UseHttpsRedirection();

app.Run();