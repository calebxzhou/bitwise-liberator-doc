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
    //要添加续表的表和行
    [HttpPost]
    public async Task<IActionResult> Post(IFormFile file,List<DocTableContinue> continues)
    {
        if (file.Length is 0 or > DocConst.MaxFileSize)
        {
            return BadRequest(">24MB || <0MB!");
        }

        await using var stream = file.OpenReadStream();
        
        using var outs = new MemoryStream();
        Process(stream, continues,outs);
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
            throw new ArgumentException("必须为每个表格指定续表！");
        }
        for (var i = 0; i < tables.Length; i++)
        {
            var table = tables[i];
            var conti = continues[i];
            if(conti.RowIndexForContinue<0)
                continue;
            var th = table.Elements<TableRow>().First().CloneNode(true);
            // Create a new table and copy the before rows from the old table.
            Table newTable1 = new Table();
            for (var j = 0; j < conti.RowIndexForContinue; j++)
            {
                var row = (TableRow)table.Elements<TableRow>().ElementAt(j).CloneNode(true);
                newTable1.Append(row);
            }
            // Insert the new table into the document.
            doc.MainDocumentPart.Document.Body.InsertAfter(newTable1, table);

            // Create a new paragraph with text "abc".
            Paragraph newParagraph = DocHeadings.H6("续"+conti.TableName).SetHorizontalAlign(JustificationValues.Right);

            // Insert the new paragraph into the document.
            doc.MainDocumentPart.Document.Body.InsertAfter(newParagraph, newTable1);

            // Create another new table and copy the remaining rows from the old table.
            Table newTable2 = new Table();
            newTable2.Append(th);
            for (int j =conti.RowIndexForContinue; j < table.Elements<TableRow>().Count(); j++)
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