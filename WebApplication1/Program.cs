using System.Text.Json;
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
var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer(); 
var app = builder.Build();
var jsonOptions = new JsonSerializerOptions
{
    PropertyNameCaseInsensitive = true,
};
// Configure the HTTP request pipeline.
 
app.UseHttpsRedirection();
 
app.MapGet("/pjtest_do", async context =>
    {
        await context.Response.WriteAsync("Hello World");
    })
    .WithName("index")
    .WithOpenApi();

app.Run();
