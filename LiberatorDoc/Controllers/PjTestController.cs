using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using LiberatorDoc.DocOps;
using LiberatorDoc.Models;
using Microsoft.AspNetCore.Mvc;
using Paragraph = DocumentFormat.OpenXml.Drawing.Paragraph;

namespace LiberatorDoc.Controllers;

[ApiController]
[Route("[controller]")]
public class PjTestController : ControllerBase
{
    [HttpPost]
    public async Task<IActionResult> Post()
    {
        using (StreamReader reader = new StreamReader(Request.Body))
        {
            //读json
            var json = await reader.ReadToEndAsync();
            var modules = JsonSerializer.Deserialize<List<ModuleTest>>(json,Options.Json)
                          ?? new List<ModuleTest>();
            //处理
            using (MemoryStream memStream = new MemoryStream())
            {
                using (var wDoc = Docs.New(memStream))
                {
                    Process(wDoc,modules);
                }
                return File(memStream.ToArray(), 
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document", 
                    "系统测试.docx");
            }
        }
    }

    private static void Process(WordprocessingDocument wDoc, List<ModuleTest> modules)
    {
        var mainPart = wDoc.AddMainDocumentPart();
        mainPart.Document = new Document();
        Body body = new Body(); 
        var p1 = Headings.CreateHeading1("3 系统测试");
        body.Append(p1);
        for (var index = 0; index < modules.Count; index++)
        {
            var moduleTest = modules[index]; 
            
            body.Append(Headings.CreateHeading2($"3.{index + 1} {moduleTest.Name}模块的测试"));
            body.Append(DocBodies.CreateMainBody($"{moduleTest.Name}模块的测试表，如表3.{index + 1}所示。"));
            var tableData = moduleTest.Cases.Select((testCase, i) => new List<string>
                { 
                    $"{i + 1}",
                    testCase.Name,
                    testCase.Operation,
                    testCase.Result,
                    testCase.Result,
                    "一致"
                })
                .ToList();
            //测试三线表
            body.Append(Tables.CreateTableNameParagraph($"表3.{index + 1} {moduleTest.Name}模块的测试表"));
            body.Append(Tables.Create3LineTable(
                
                new[]
                {
                    new TableColumnProps(700,"编号",JustificationValues.Center,false),
                    new TableColumnProps(1550,"测试项",JustificationValues.Center,false),
                    new TableColumnProps(2270,"描述输入/操作",JustificationValues.Both,true),
                    new TableColumnProps(1700,"预计结果",JustificationValues.Both,true),
                    new TableColumnProps(1560,"实际结果",JustificationValues.Both,true),
                    new TableColumnProps(1275,"结果对比",JustificationValues.Center,false)
                } ,tableData
            )); 
        
        }
        mainPart.Document.AppendChild(body);
        mainPart.Document.Save();
    }
}