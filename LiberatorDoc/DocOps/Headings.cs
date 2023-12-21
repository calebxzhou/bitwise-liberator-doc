using DocumentFormat.OpenXml.Wordprocessing;
using static LiberatorDoc.DocOps.DocConst;

namespace LiberatorDoc.DocOps;

//标题
public static class Headings
{
//创建一级标题：  
    public static Paragraph CreateHeading1(string text)
    {
        Paragraph para = new Paragraph();
        //居中
        para.SetParagraphHorizontalAlign(JustificationValues.Center);
        //段前、段后均为1行，行间距为固定值22磅；
        para.SetParagraphSpacing(SpaceBeforeAfter1Line);
        para.SetParagraphBeforeAfterLines();
        ParagraphProperties paraProps = new ParagraphProperties();

        //设置为可折叠
        OutlineLevel outlineLevel1 = new OutlineLevel() { Val = 1 };
        paraProps.Append(outlineLevel1);
        //三号黑体
        paraProps.Append(Fonts.GetFontProp(SimHei, Size3));
        para.Append(paraProps);
        Run run = new Run();
        //三号黑体
        run.Append(Fonts.GetFontProp(SimHei, Size3));
        run.Append(new Text(text));
        para.Append(run);
        return para;
    }

//创建二级标题 
    public static Paragraph CreateHeading2(string text)
    {
        Paragraph para = new Paragraph();
        //居左
        para.SetParagraphHorizontalAlign(JustificationValues.Left);
        //段前、段后均为1行，行间距为固定值22磅；
        para.SetParagraphSpacing(SpaceBeforeAfter12);
        para.SetParagraphBeforeAfterLines();
        ParagraphProperties paraProps = new ParagraphProperties();
        //设置为可折叠 
        OutlineLevel outlineLevel1 = new OutlineLevel() { Val = 2 };
        paraProps.Append(outlineLevel1);
        //四号黑体
        paraProps.Append(Fonts.GetFontProp(SimHei, Size4));
        para.Append(paraProps);
        Run run = new Run();
        //四号黑体
        run.Append(Fonts.GetFontProp(SimHei, Size4));
        run.Append(new Text(text));
        para.Append(run);
        return para;
    }

//创建三级标题： 
    public static Paragraph CreateHeading3(string text)
    {
        Paragraph para = new Paragraph();
        //居左
        para.SetParagraphHorizontalAlign(JustificationValues.Left);
        //段前、段后均为6磅，行间距为固定值22磅；
        para.SetParagraphSpacing(SpaceBeforeAfter6);
        ParagraphProperties paraProps = new ParagraphProperties();
        // 小四号，黑体
        paraProps.Append(Fonts.GetFontProp(SimHei, Size4S));
        para.Append(paraProps);
        Run run = new Run();
        // 小四号，黑体
        run.Append(Fonts.GetFontProp(SimHei, Size4S));
        run.Append(new Text(text));
        para.Append(run);
        return para;
    }

//创建小标题： 
    public static Paragraph CreateHeadingS(string text)
    {
        Paragraph para = new Paragraph();
        //居左
        para.SetParagraphHorizontalAlign(JustificationValues.Left);
        // 行间距为固定值22磅；
        para.SetParagraphSpacing(0);
        //首行缩进2字符
        text = ChineseSpace + ChineseSpace + text;
        ParagraphProperties paraProps = new ParagraphProperties();
        // 小四号，宋体
        paraProps.Append(Fonts.GetFontProp(SimSun, Size4S));
        para.Append(paraProps);
        Run run = new Run();
        // 小四号，宋体
        run.Append(Fonts.GetFontProp(SimSun, Size4S));
        run.Append(new Text(text));
        para.Append(run);
        return para;
    }
}