using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Tables;
using Microsoft.AspNetCore.Mvc;
using Document = Aspose.Words.Document;
using Paragraph = Aspose.Words.Paragraph;
using Table = Aspose.Words.Tables.Table;

namespace LiberatorDoc.Controllers;

[ApiController]
[Route("[controller]")]
public class DocUpController : ControllerBase
{
    private const long MaxFileSize = 24 * 1024 * 1024; // 24MB

    [HttpPost]
    public async Task<IActionResult> Post(IFormFile? file)
    {
        if (file == null || file.Length == 0)
        {
            return BadRequest("No file uploaded.");
        }

        if (file.Length > MaxFileSize)
        {
            return BadRequest(">24MB!");
        }

        using var stream = file.OpenReadStream();
        var tempFilePath = Path.GetTempFileName();
        using var fileStream = new FileStream(tempFilePath, FileMode.Create, FileAccess.Write);
        stream.CopyTo(fileStream);


        // Now the file's data is stored in memory in 'memoryStream'
        // You can process it as needed

        System.IO.File.Delete(tempFilePath);
        return Ok("File uploaded successfully.");
    }

    //优化表格（三线表边框 添加续表）
    public static void ProcessTableOpt(Stream fileIn,Stream fileOut)
    {
        var doc = new Document(fileIn);
        var nodes = doc.GetChildNodes(NodeType.Any, true).ToArray();
        for (var i = 0; i < nodes.Length; i++)
        {
            var node = nodes[i];
            if (node.NodeType == NodeType.Paragraph)
            {
                var para = node as Paragraph;
                var paraText = para.GetText() ?? "";
                string pattern = @"(表\d+\.\d+)\D*";
                var match = Regex.Match(paraText, pattern);
                if (!match.Success) continue;
                //  表名 （表x.x）
                var tableName = match.Groups[1].Value;
                if (i + 1 == nodes.Length)
                    continue;
                Node? nextNode = null;
                //表名的下一个节点必须是 表格
                for (var j = i; j < nodes.Length; j++)
                {
                    if (nodes[j].NodeType == NodeType.Table) 
                        nextNode = nodes[j];
                }
                if(nextNode == null)
                    continue;
                
                var table = nextNode as Table;
                var layout = new LayoutCollector(doc);
                var rows = table.Rows;
                
                for (var j = 0; j < rows.Count; j++)
                {
                    var row = rows[j];
                    var rowPageNow = layout.GetStartPageIndex(row);
                    //不看最后一行
                    if (j + 1 == rows.Count) continue;
                    var nextRow = rows[j + 1];
                    var rowPageNext = layout.GetStartPageIndex(nextRow);
                    Console.WriteLine(row.GetText()+rowPageNow+rowPageNext);
                    //找出 前后不在同一页的 表行
                    if (rowPageNext > rowPageNow)
                    {
                        //表头
                        var tableHeader = table.FirstRow.Clone(true) as Row;
                        //插入新的表头
                        row.ParentNode.InsertAfter(tableHeader, row);
                    }
                }
            }
        }

        doc.Save("test3.docx");
    }
}