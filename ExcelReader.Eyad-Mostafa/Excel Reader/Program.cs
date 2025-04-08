namespace Excel_Reader;

internal class Program
{
    private static void Main(string[] args)
    {
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.WriteLine("=========================================");
        Console.WriteLine("           EMPLOYEE IMPORT TOOL          ");
        Console.WriteLine("=========================================");
        Console.ResetColor();

        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine("\n[INFO] Reading From the file...");
        Console.ResetColor();

        var employees = FileReader.ReadEmployees("Assets\\employees.xlsx");

        if (employees.Count == 0)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("[ERROR] No employees found");
            Console.ResetColor();
            return;
        }

        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine("[SUCCESS] Reading completed.");
        Console.ResetColor();


        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine("\n[INFO] Saving to the database...");
        Console.ResetColor();

        DatabaseManager.SaveEmployees(employees);

        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine("[SUCCESS] Saving completed.");
        Console.ResetColor();

        var databaseEmployees = DatabaseManager.GetEmployees();

        foreach (var employee in databaseEmployees)
        {
            Console.WriteLine($"ID: {employee.ID}, Name: {employee.Name}, Age: {employee.Age}, Department: {employee.Department}");
        }


        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.WriteLine("\n=============================================");
        Console.WriteLine("        PROCESS COMPLETED SUCCESSFULLY   ");
        Console.WriteLine("=============================================");
        Console.ResetColor();
    }

}
