using System.Text.Json;
using SharpDocx;

var builder = WebApplication.CreateBuilder(args);
// Add CORS services.
builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowAll",
        builder =>
        {
            builder
                .AllowAnyOrigin()
                .AllowAnyMethod()
                .AllowAnyHeader();
        });
});
builder.Services.AddEndpointsApiExplorer(); 
var app = builder.Build();

var jsonOptions = new JsonSerializerOptions
{
    PropertyNameCaseInsensitive = true,
};

// Configure the HTTP request pipeline.
 
app.UseHttpsRedirection();
 
app.UseCors(b => b
    .AllowAnyOrigin()
    .AllowAnyMethod()
    .AllowAnyHeader());
app.MapPost("/pjtest_do", async (HttpRequest req) => 
{
    // Read the request body as a string
    var bodyString = await new StreamReader(req.Body).ReadToEndAsync();
    Console.WriteLine(bodyString);
    // Deserialize the string into a PjTest object
    var modules = JsonSerializer.Deserialize<List<ModuleTest>>(bodyString,jsonOptions)
                  ?? new List<ModuleTest>();
    var pjTest = new PjTest(modules);
    var document = DocumentFactory.Create("pjtest.docx",modules);
    document.Generate("gen.docx");
    for (var i = 0; i < modules.Count; i++)
    {
        
    }
    // Now you can use the deserialized 'pjTest' object
    // ...

    return Results.Ok();
});

app.Run();

[Serializable]
public record TestCase(string Name, string Operation, string Result);
// 测试项
// 描述输入/操作
// 期望+真实结果
[Serializable]
public record ModuleTest(string Name, List<TestCase> Cases)
{
    // 模块名
    // 全部测试用例
   
}
// 项目测试
[Serializable]
public record PjTest(List<ModuleTest> Modules)
{
    // 全部模块测试
    
}


