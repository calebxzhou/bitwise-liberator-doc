using DocumentFormat.OpenXml.Wordprocessing;

namespace LiberatorDoc.DocOps;

public static class Paragraphs
{
    //设定段落行间距22磅 前后指定 
    public static void SetParagraphSpacing(this Paragraph paragraph,int beforeAfterSpace)
    {
        paragraph.Append(new ParagraphProperties(
            new SpacingBetweenLines()
            {
                Before = $"{beforeAfterSpace}", 
                After = $"{beforeAfterSpace}", 
                Line = $"{DocConst.LineSpacing}", 
                LineRule = LineSpacingRuleValues.Exact
            }
        ));
    }
    //设定段落前后间距1行 
    public static void SetParagraphBeforeAfterLines(this Paragraph paragraph)
    {
        paragraph.Append(new ParagraphProperties(
            new SpacingBetweenLines()
            {
                BeforeLines = 100,AfterLines = 100
            }
        ));
    }
    //设定段落对齐模式 水平
    public static void SetParagraphHorizontalAlign(this Paragraph paragraph, JustificationValues hAlign)
    {
        paragraph.Append(new ParagraphProperties(
            new Justification() { Val = hAlign }
        )); 
    }  
}