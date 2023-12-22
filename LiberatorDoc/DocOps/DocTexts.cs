using DocumentFormat.OpenXml.Wordprocessing;

namespace LiberatorDoc.DocOps;

public static class DocTexts
{
    public static Run Underlined(string font, int fontSize,string text)
    {
        Run run = new Run();
        Text t = new Text(text);
        run.Append(t);
        Underline underline = new Underline() { Val = UnderlineValues.Single };
        run.RunProperties = new RunProperties(underline);
        run.RunProperties.AppendFontProp(font, fontSize);
        return run;
    }
    public static Run Normal(string font, int fontSize,string text)
    {
        // Create a new run
        Run run = new Run();
        // Create a new text element
        Text t = new Text(text);
        // Add the text to the run
        run.Append(t);
        run.RunProperties = new RunProperties();
        run.RunProperties.AppendFontProp(font, fontSize);
        return run;
    }
}