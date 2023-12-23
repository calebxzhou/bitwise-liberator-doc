using LiberatorDoc.Dsl;

namespace LiberatorDocTest;
public class Tests
{
    [SetUp]
    public void Setup()
    {
    }

    [Test]
    public void Test1()
    {
        
        MemoryStream memoryStream = new MemoryStream();
// Write to memoryStream...
        memoryStream.WriteDocxFromDsl(@"
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
");
        using (FileStream fileStream = new FileStream("test.docx", FileMode.Create, FileAccess.Write))
        {
            memoryStream.WriteTo(fileStream);
        }
        Assert.Pass();
    }
}