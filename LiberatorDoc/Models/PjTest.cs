namespace LiberatorDoc.Models;
// 测试项
// 描述输入/操作
// 期望+真实结果
[Serializable]
public class ModuleTest(string Name, List<TestCase> Cases)
{
    // 模块名
    // 全部测试用例

    public string Name { get; init; } = Name;
    public List<TestCase> Cases { get; init; } = Cases;

    public void Deconstruct(out string Name, out List<TestCase> Cases)
    {
        Name = this.Name;
        Cases = this.Cases;
    }
}
[Serializable]
public class TestCase(string Name, string Operation, string Result)
{
    public string Name { get; init; } = Name;
    public string Operation { get; init; } = Operation;
    public string Result { get; init; } = Result;

    public void Deconstruct(out string Name, out string Operation, out string Result)
    {
        Name = this.Name;
        Operation = this.Operation;
        Result = this.Result;
    }
}
