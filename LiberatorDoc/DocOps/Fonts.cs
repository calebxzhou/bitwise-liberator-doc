using DocumentFormat.OpenXml.Wordprocessing;

namespace LiberatorDoc.DocOps;

public class Fonts
{
    
    /// <summary>
    /// 获取字体属性
    /// </summary>
    /// <param name="fontName">字体名</param>
    /// <param name="fontSize">字体尺寸</param>
    /// <returns>字体属性</returns>
    public static RunProperties GetFontProp(string fontName, int fontSize)
    {
        var paraRunProps = new RunProperties();
        paraRunProps.Append(new RunFonts() { Ascii = DocConst.TimesNewRoman, EastAsia = fontName, HighAnsi = fontName,Hint  = FontTypeHintValues.EastAsia});
        paraRunProps.Append(new FontSize() { Val = $"{fontSize}" });
        paraRunProps.Append(new FontSizeComplexScript() { Val = $"{fontSize}" });
        return paraRunProps;
    } 

}