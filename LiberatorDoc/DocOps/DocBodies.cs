using DocumentFormat.OpenXml.Wordprocessing;
using static LiberatorDoc.DocOps.DocConst;

namespace LiberatorDoc.DocOps;

public class DocBodies
{
    ///正文： 
    public static Paragraph Main(string text)
    {
        var para = new Paragraph();
        //两端对齐
        para.SetParagraphHorizontalAlign(JustificationValues.Both);
        // 行间距为固定值22磅；
        para.SetParagraphSpacing(0);
        //首行缩进2字符
        text = ChineseSpace + ChineseSpace + text;
        var paraProps = new ParagraphProperties();
        // 小四号，宋体
        paraProps.Append(DocFonts.GetFontProp(SimSun,Size4S));
        para.Append(paraProps);
        var run = new Run();
        // 小四号，宋体
        run.Append(DocFonts.GetFontProp(SimSun,Size4S));
        run.Append(new Text(text));
        para.Append(run);
        return para;
    }
}