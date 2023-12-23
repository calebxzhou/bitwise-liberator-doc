using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using LiberatorDoc.DocOps;

namespace LiberatorDoc.Dsl;

public record DocDslRow(DocElementType Type, params string[] Tokens)
{
    //合并第一个空格后面的tokens为一个整体
    public string Merged = string.Join(" ", Tokens);
}

public static class DocDsl
{
    //th 表头 字+宽度+c居中l居左r居右b两端+s开头1空格
    /*
     h1 2 系统实现
     h2 2.1 系统框架
     p 地标旅游管理信息系统使用Django框架。工程目录结构图，如图2.1所示。
     h6 图2.1 工程目录结构图
     p middleware是修改Django或者response对象的钩子。浏览器从请求到响应的过程中，Django需要通过很多中间件来处理。middlewar包的说明表，如表2.1所示。
     h6 表2.1 middleware包的说明表
     th 文件名4536c 作用4536c
     tr auth.py	进行权限管理
     tr auth.py	进行权限管理
     tr auth.py	进行权限管理
     h3 test
     h4 test
     */
    public static void WriteDocxFromDsl(this MemoryStream stream, string dsl)
    {
        using var wDoc = Docs.New(stream);
        var mainPart = wDoc.AddMainDocumentPart();
        mainPart.Document = new Document();
        Body body = new Body(); 
        foreach (var element in CompileRowsToXmlElements(CompileDslToRows(dsl)))
        {
            body.Append(element);
        }
        mainPart.Document.AppendChild(body);
        mainPart.Document.Save();
    }
    //把dsl编译成row列表
    private static List<DocDslRow> CompileDslToRows(string dsl)
    {
        var dslRows = new List<DocDslRow>();
        var rows = dsl.Split("\n");
        foreach (var row in rows)
        {
            var tokens = row.Trim().Split(" ");
            //找对应的类型
            if (!Enum.TryParse(tokens[0], out DocElementType elementType)) continue;
            var dslRow = new DocDslRow(elementType, tokens.Skip(1).ToArray());
            dslRows.Add(dslRow);
        }
        return dslRows;
    }
    //把row列表编译成ooxml
    private static List<OpenXmlElement> CompileRowsToXmlElements(IReadOnlyList<DocDslRow> rows)
    {
        var elements = new List<OpenXmlElement>();
        for (var i = 0; i < rows.Count; i++)
        {
            var row = rows[i];
            var mergedContent = row.Merged;
            switch (row.Type)
            {
                case DocElementType.h1:
                    elements.Add(DocHeadings.H1(mergedContent));
                    break;
                case DocElementType.h2:
                    elements.Add(DocHeadings.H2(mergedContent));
                    break;
                case DocElementType.h3:
                    elements.Add(DocHeadings.H3(mergedContent));
                    break;
                case DocElementType.h4:
                    elements.Add(DocHeadings.H4(mergedContent));
                    break;
                case DocElementType.h6:
                    elements.Add(DocHeadings.H6(mergedContent));
                    break;
                case DocElementType.p:
                    elements.Add(DocBodies.Main(mergedContent));
                    break;
                case DocElementType.th:
                    i += HandleTable(rows, i, elements);
                    break;
                case DocElementType.img:
                    //TODO
                    break;
                case DocElementType.tr:
                    throw new InvalidDataException("编译DocDSL时，tr不应该被独立读取");
                default:
                    throw new InvalidDataException("无效的DocElementType");
            }
        }
        return elements;
    }
    //处理表格相关DSL，返回offset（index跳过后面的tr）
    private static int HandleTable(IReadOnlyList<DocDslRow> rows, int indexNow, List<OpenXmlElement> elements)
    {
        int offset = 0;
        //列属性
        var thRow = rows[indexNow];
        var colProps = (from thRowToken in thRow.Tokens
            let addSpace = thRowToken.EndsWith('s')
            let align = thRowToken[^2] switch
            {
                'l' => JustificationValues.Left,
                'r' => JustificationValues.Right,
                'b' => JustificationValues.Both,
                'c' => JustificationValues.Center,
                _ => JustificationValues.Center
            }
            let width = Regex.Matches(thRowToken, @"\d+")[0].Value
            let name = Regex.Matches(thRowToken, "[\u4e00-\u9fa5]")[0].Value
            select new TableColumnProps(Convert.ToInt32(width), name, align, addSpace)).ToArray();
        var tableData = new List<List<string>>();
        //遇到表头，就把下面所有的表行读出来
        while (rows[indexNow + 1].Type != DocElementType.tr)
        {
            var trRow = rows[indexNow];
            tableData.Add(trRow.Tokens.ToList());
            indexNow++;
            offset++;
        }
        elements.Add(DocTables.Create3LineTable(colProps,tableData));
        return offset;
    }
}
 