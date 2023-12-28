using System.Diagnostics;
using System.Text.Json;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using LiberatorDoc.DocOps;
using LiberatorDoc.Dsl;
using LiberatorDoc.Models;
using Microsoft.AspNetCore.Mvc;

namespace LiberatorDoc.Controllers;

//上传word 返回表格信息
[ApiController]
[Route("[controller]")]
public class DocTablesOptController : ControllerBase
{
    public record Dto(string file, List<DocTableContinue> continues);
    //要添加续表的表和行
    [HttpPost]
    public async Task<IActionResult> Post()
    {
        using var reader = new StreamReader(Request.Body);
        var body = await reader.ReadToEndAsync();
        if (body.Length is 0 or > DocConst.MaxFileSize)
        {
            return BadRequest(">24MB || <0MB!");
        }

        var dto = JsonSerializer.Deserialize<Dto>(body);
        byte[] byteArray = Convert.FromBase64String(dto.file);
        using var  stream = new MemoryStream(byteArray);
        using var outs = new MemoryStream();
        Process(stream, dto.continues ,outs);
        return File(outs.ToArray(), 
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document", 
            "1.docx");
    }

    public static void Process(Stream stream, List<DocTableContinue> continues, Stream outs)
    {
        stream.CopyTo(outs);
        using var doc = WordprocessingDocument.Open(outs, true);
        var tables = doc.MainDocumentPart.Document.Body.Elements<Table>().ToArray();
        if (continues.Count != tables.Length)
        {
          //  throw new ArgumentException("必须为每个表格指定续表！");
        }
        for (var i = 0; i < continues.Count; i++)
        {
            var conti = continues[i];
            var table = tables[conti.tableIndex];
            if(conti.rowIndex<0)
                continue;
            var th = table.Elements<TableRow>().First().CloneNode(true);
            // Create a new table and copy the before rows from the old table.
            Table newTable1 = new Table();
            for (var j = 0; j < conti.rowIndex; j++)
            {
                var row = (TableRow)table.Elements<TableRow>().ElementAt(j).CloneNode(true);
                newTable1.Append(row);
            }
            // Insert the new table into the document.
            doc.MainDocumentPart.Document.Body.InsertAfter(newTable1, table);

            // Create a new paragraph with text "abc".
            Paragraph newParagraph = DocHeadings.H6("续"+conti.tableName).SetHorizontalAlign(JustificationValues.Right);

            // Insert the new paragraph into the document.
            doc.MainDocumentPart.Document.Body.InsertAfter(newParagraph, newTable1);

            // Create another new table and copy the remaining rows from the old table.
            Table newTable2 = new Table();
            newTable2.Append(th);
            for (int j =conti.rowIndex; j < table.Elements<TableRow>().Count(); j++)
            {
                TableRow row = (TableRow)table.Elements<TableRow>().ElementAt(j).CloneNode(true);
                newTable2.Append(row);
            }

            // Insert the second new table into the document.
            doc.MainDocumentPart.Document.Body.InsertAfter(newTable2, newParagraph);
            table.Remove();
            
        }
 
    }
}