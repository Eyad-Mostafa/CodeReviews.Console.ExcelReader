using Excel_Reader.Model;
using OfficeOpenXml;

namespace Excel_Reader;

internal class FileReader
{
    public static List<Employee> ReadEmployees(string path)
    {
        var employees = new List<Employee>();

        if (!File.Exists(path))
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"[ERROR] The file '{path}' was not found.");
            Console.ResetColor();
            return employees;
        }

        try
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(path));

            var worksheet = package.Workbook.Worksheets[0];
            if (worksheet == null || worksheet.Dimension == null)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[ERROR] The Excel file is empty or corrupt.");
                Console.ResetColor();
                return employees;
            }

            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("[INFO] Reading employees from Excel...");
            Console.ResetColor();

            for (var i = 2; i <= worksheet.Dimension.End.Row; i++)
            {
                if (worksheet.Cells[i, 1].Value == null || worksheet.Cells[i, 2].Value == null)
                {
                    Console.ForegroundColor = ConsoleColor.DarkYellow;
                    Console.WriteLine($"[WARNING] Skipping row {i} due to missing data.");
                    Console.ResetColor();
                    continue;
                }

                var employee = new Employee
                {
                    ID = int.TryParse(worksheet.Cells[i, 1].Value?.ToString(), out int id) ? id : 0,
                    Name = worksheet.Cells[i, 2].Value?.ToString() ?? "Unknown",
                    Age = int.TryParse(worksheet.Cells[i, 3].Value?.ToString(), out int age) ? age : 0,
                    Department = worksheet.Cells[i, 4].Value?.ToString() ?? "Unassigned"
                };

                employees.Add(employee);
            }

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"[SUCCESS] Successfully read {employees.Count} employees from Excel.");
            Console.ResetColor();
        }
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"[ERROR] An error occurred while reading the file: {ex.Message}");
            Console.ResetColor();
        }

        return employees;
    }
}
