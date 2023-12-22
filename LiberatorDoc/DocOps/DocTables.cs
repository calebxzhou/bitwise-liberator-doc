using DocumentFormat.OpenXml.Wordprocessing;
using static LiberatorDoc.DocOps.DocConst;

namespace LiberatorDoc.DocOps;

public record class TableColumnProps(int Width, string Header, JustificationValues HAlign, bool AddSpaceBefore)
{
    
}
public static class DocTables
{
    //三线表 表名+表格 段落
    public static Paragraph New(string tableName, TableColumnProps[] props, List<List<string>> datas)
    {
        var para = CreateTableNameParagraph(tableName);
        para.Append( Create3LineTable(props,datas));
        return para;
    }

    public static Table Create3LineTable(TableColumnProps[] props,List<List<string>> contents)
    {
        var table = new Table();
        var tableProperties = new TableProperties();
        //单元格间距
        var tableCellMarginDefault = new TableCellMarginDefault(new TopMargin()
            {
                Width = "0", Type = TableWidthUnitValues.Dxa
            }, new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa },
            new LeftMargin() { Width = "100", Type = TableWidthUnitValues.Dxa },
            new RightMargin() { Width = "100", Type = TableWidthUnitValues.Dxa });
        tableProperties.Append(tableCellMarginDefault);
        //设定表格宽度 6.3''
        var tableWidth = new TableWidth() { Width = "9072", Type = TableWidthUnitValues.Dxa };
        tableProperties.Append(tableWidth);

        table.Append(tableProperties);

        //绘制表头
        var headerRow = new TableRow();
        for (var index = 0; index < props.Length; index++)
        {
            var prop = props[index];
            var header = prop.Header;
            var headerCell = CreateTextTableCellAlign(header, JustificationValues.Center,
                TableVerticalAlignmentValues.Bottom, prop.Width);
            //设置边框 表头上1磅下0.75磅
            SetCellBorders(headerCell, 8, 6);
            headerRow.Append(headerCell);
        }

        table.Append(headerRow);

        //绘制表格主要内容
        for (var rowIndex = 0; rowIndex < contents.Count; rowIndex++)
        {
            var contentRow = contents[rowIndex];
            var row = new TableRow();
            for (var colIndex = 0; colIndex < contentRow.Count; colIndex++)
            {
                var prop = props[colIndex];
                var content = contentRow[colIndex];
                if (prop.AddSpaceBefore)
                {
                    content = ChineseSpace + content;
                }

                //单元格
                var cell = CreateTextTableCellAlign(content, prop.HAlign, TableVerticalAlignmentValues.Top,
                    prop.Width);
                row.Append(cell);
                //最后一行 设定底边框1磅
                if (rowIndex == contents.Count - 1)
                {
                    SetCellBorders(cell, 0, 8);
                }
            }

            table.Append(row);
        }

        return table;
    }

    /// <summary>
    /// 设置单元格边框
    /// </summary>
    /// <param name="cell">单元格</param>
    /// <param name="topSize">上边框 尺寸</param>
    /// <param name="bottomSize">下边框 尺寸</param>
    private static void SetCellBorders(this TableCell cell, uint topSize, uint bottomSize)
    {
        TableCellBorders borders = new TableCellBorders();
        if (bottomSize > 0)
        {
            var bottomBorder = new BottomBorder()
            {
                Val = BorderValues.Single,
                Size = bottomSize,
                Color = "auto"
            };

            borders.Append(bottomBorder);
        }
        if (topSize > 0)
        {
            var topBorder = new TopBorder()
            {
                Val = BorderValues.Single,
                Size = topSize,
                Color = "auto"
            };

            borders.Append(topBorder);
        }
        TableCellProperties cellProperties = cell.AppendChild(new TableCellProperties());
        cellProperties.Append(borders);
    }

    //创建单元格 文字，水平对齐，垂直对齐，宽度
    private static TableCell CreateTextTableCellAlign(string text,
        JustificationValues hAlign,
        TableVerticalAlignmentValues vAlign,
        int width)
    {
        var para = new Paragraph();
        //段落对齐 行间距22磅 前后0
        para.SetParagraphSpacing(0);
        para.SetParagraphHorizontalAlign(hAlign);
        //宋体五号字
        var paraProps = new ParagraphProperties();
        paraProps.Append(DocFonts.GetFontProp(SimSun, Size5));
        para.Append(paraProps);

        var run = new Run();
        run.Append(DocFonts.GetFontProp(SimSun, Size5));
        run.Append(new Text(text));
        para.Append(run);


        var cell = new TableCell();
        //设定宽度 对齐
        cell.Append(new TableCellProperties(
            new TableCellWidth { Width = $"{width}", Type = TableWidthUnitValues.Dxa },
            new TableCellVerticalAlignment { Val = vAlign },
            new Justification { Val = hAlign }));
        cell.Append(para);
        return cell;
    }
    //表名段落 （eg 表3.1 管理员端的测试表）
    public static Paragraph CreateTableNameParagraph(string tableName)
    {
        var para = new Paragraph();
        para.SetParagraphSpacing(0);
        para.SetParagraphHorizontalAlign(JustificationValues.Center);
        //黑体五号字
        var paraProps = new ParagraphProperties();
        paraProps.Append(DocFonts.GetFontProp(SimHei, Size5));
        para.Append(paraProps);

        var run = new Run();
        run.Append(DocFonts.GetFontProp(SimHei, Size5));
        run.Append(new Text(tableName));
        para.Append(run);

        return para;
    }
}