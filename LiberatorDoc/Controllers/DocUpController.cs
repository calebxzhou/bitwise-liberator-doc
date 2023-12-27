using LiberatorDoc.DocOps;
using LiberatorDoc.Dsl;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Office.Interop.Word;

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

    public static void Process(string filePath)
    {
        var wordApp = new Application();
        var doc = wordApp.Documents.Open(filePath);
        // Get all the text in the document
        string allText = doc.Content.Text;

// Print the text
        Console.WriteLine(allText);

// Close the Word document and quit the Word application
        doc.Close();
        wordApp.Quit();
    }

}