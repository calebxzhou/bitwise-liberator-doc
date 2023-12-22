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
public class DbTableController : ControllerBase
{
    [HttpPost]
    public async Task<IActionResult> Post()
    {
        using (StreamReader reader = new StreamReader(Request.Body))
        {
            //读json
            var json = await reader.ReadToEndAsync();
            var modules = JsonSerializer.Deserialize<List<DbTable>>(json,Options.Json)
                          ?? new List<DbTable>();
            //处理
            using (MemoryStream memStream = new MemoryStream())
            {
                using (var wDoc = Docs.New(memStream))
                {
                    Process(wDoc,modules);
                }
                return File(memStream.ToArray(), 
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document", 
                    "pjtest.docx");
            }
        }
    }

    private static void Process(WordprocessingDocument wDoc, List<DbTable> tables)
    {
        var mainPart = wDoc.AddMainDocumentPart();
        mainPart.Document = new Document();
        Body body = new Body(); 
        body.Append( DocHeadings.H2("1.3 数据库设计"));
        body.Append(DocBodies.Main("MySQL 是最流行的数据库之一，是一个免费开源的关系型数据库管理系统，具有方便小巧、运行速度快等特点。此小节从数据库的概念结构设计、逻辑结构设计、物理结构设计三个方面对数据表进行介绍。"));
        body.Append(DocHeadings.H4("1. 概念结构设计"));
        body.Append(DocBodies.Main("本系统数据库E-R图，如图1.3所示。"));
        body.Append(DocHeadings.H6("图1.3 本系统数据库E-R图"));
        body.Append(DocHeadings.H4("2. 逻辑结构设计"));
        body.Append(DocBodies.Main("由实体关系图转换关系模式，结果如下："));
        for (var i = 0; i < tables.Count; i++)
        {
            var table = tables[i];
            var p1 = DocBodies.Main($"（{i + 1}）{table.Name}（");
            var colNames = table.Columns
                .Where(col => !string.IsNullOrWhiteSpace(col.Name))
                .Select(col => col.Name)
                .ToList();
            for (var j = 0; j < colNames.Count; j++)
            {
                //第一个字段是主键 有下划线
                if (j == 0)
                {
                    var run = DocTexts.Underlined(DocConst.SimSun, DocConst.Size4S, colNames[j]);
                    p1.Append(run);
                }
                //往后的字段没有下划线
                else
                {
                    p1.Append(DocTexts.Normal(DocConst.SimSun, DocConst.Size4S,colNames[j]));
                }
                //除了最后一个字段，后面都有逗号
                if (j != colNames.Count - 1)
                {
                    p1.Append(DocTexts.Normal(DocConst.SimSun, DocConst.Size4S,"，"));
                }
            }

            p1.Append(DocTexts.Normal(DocConst.SimSun, DocConst.Size4S,"）。"));
            body.Append(p1);
        }
        body.Append(DocHeadings.H4("3. 物理结构设计"));
        for (var i = 0; i < tables.Count; i++)
        {
            var table = tables[i];
            var tableData = table.Columns.Select(((column,j) => new List<string>
            {
                column.Id,
                column.Type,
                column.Len.Length==0?"——":column.Len,
                "否",
                j==0 ? "是":"否",
                column.Name
            })).ToList();
            body.Append(DocBodies.Main(table.Name+"表"+table.Id+"，如表1."+(i+1)+"所示。"));
            //三线表
            body.Append(DocHeadings.H6($"表1.{i+1} {table.Name}表{table.Id}"));
            body.Append(DocTables.Create3LineTable(
                new[]
                {
                    new TableColumnProps(1985,"字段名",JustificationValues.Center,false),
                    new TableColumnProps(1560,"数据类型",JustificationValues.Center,false),
                    new TableColumnProps(1275,"长度",JustificationValues.Center,false),
                    new TableColumnProps(1275,"是否为空",JustificationValues.Center,false),
                    new TableColumnProps(1275,"是否主键",JustificationValues.Center,false),
                    new TableColumnProps(1700,"描述",JustificationValues.Center,false)
                } ,tableData
            )); 
        }
        mainPart.Document.AppendChild(body);
        mainPart.Document.Save();
    }
}