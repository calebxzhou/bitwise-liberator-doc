namespace LiberatorDoc.Models;

public class DbTable
{
    public string Id { get; set; }
    public string Name { get; set; }
    public List<Column> Columns { get; set; }

    public DbTable(string id, string name, List<Column> columns)
    {
        Id = id;
        Name = name;
        Columns = columns;
    }
}

public class Column
{
    public string Id { get; set; }
    public string Name { get; set; }
    public string Type { get; set; }
    public string Len { get; set; }

    public Column(string id, string name, string type, string len)
    {
        Id = id;
        Name = name;
        Type = type;
        Len = len;
    }
}
