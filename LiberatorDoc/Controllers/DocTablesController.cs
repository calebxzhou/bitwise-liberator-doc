using System.Diagnostics;
using DocumentFormat.OpenXml.Packaging;
using LiberatorDoc.DocOps;
using LiberatorDoc.Dsl;
using Microsoft.AspNetCore.Mvc;

namespace LiberatorDoc.Controllers;

//上传word 返回表格信息
[ApiController]
[Route("[controller]")]
public class DocTablesController : ControllerBase 
{
    [HttpPost]
    public async Task<IActionResult> Post(IFormFile file)
    {
        if (file.Length is 0 or > DocConst.MaxFileSize)
        {
            return BadRequest(">24MB || <0MB!");
        }

        await using var stream = file.OpenReadStream();
        using var doc = WordprocessingDocument.Open(stream, true);
        using var streamOut = new MemoryStream();
        Process(doc, streamOut);
        return File(streamOut.ToArray(), 
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document", 
            "1.docx");
    }

    public static void Process(WordprocessingDocument doc, Stream streamOut)
    {
        
    }
}