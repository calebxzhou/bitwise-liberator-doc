namespace LiberatorDoc.Models;


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
