namespace Excel_Reader.Model;

internal class Employee
{
    public int ID { get; set; }
    public string Name { get; set; } = null!;
    public int Age { get; set; }
    public string? Department { get; set; }
}
