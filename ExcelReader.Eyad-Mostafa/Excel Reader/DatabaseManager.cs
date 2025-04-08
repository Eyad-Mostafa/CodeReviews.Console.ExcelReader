using Excel_Reader.Model;
using Microsoft.Data.SqlClient;
using System.Data;

namespace Excel_Reader;

internal class DatabaseManager
{
    private const string ConnectionString = "Server=.\\MSSQLSERVER08;TrustServerCertificate=True;Integrated Security=True;";
    private const string DatabaseConnectionString = "Server=.\\MSSQLSERVER08;Database=Employees;TrustServerCertificate=True;Integrated Security=True;";

    public static void SaveEmployees(List<Employee> employees)
    {
        try
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("[INFO] Initializing database...");
            Console.ResetColor();

            using var connection = new SqlConnection(ConnectionString);
            connection.Open();

            DropAndCreateDatabase(connection);
            CreateEmployeeTable();

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("[SUCCESS] Database initialized successfully.");
            Console.ResetColor();

            InsertEmployees(employees);

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("[SUCCESS] Employees saved successfully.");
            Console.ResetColor();
        }
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"[ERROR] {ex.Message}");
            Console.ResetColor();
        }
    }

    private static void DropAndCreateDatabase(SqlConnection connection)
    {
        using var command = connection.CreateCommand();
        command.CommandText = @"
            DROP DATABASE IF EXISTS Employees;
            CREATE DATABASE Employees;";
        command.ExecuteNonQuery();
    }

    private static void CreateEmployeeTable()
    {
        using var connection = new SqlConnection(DatabaseConnectionString);
        connection.Open();

        using var command = connection.CreateCommand();
        command.CommandText = @"
            CREATE TABLE Employees (
                ID INT PRIMARY KEY,
                Name NVARCHAR(100),
                Age INT,
                Department NVARCHAR(50)
            );";
        command.ExecuteNonQuery();
    }

    private static void InsertEmployees(List<Employee> employees)
    {
        using var connection = new SqlConnection(DatabaseConnectionString);
        connection.Open();

        using var transaction = connection.BeginTransaction();
        using var command = connection.CreateCommand();
        command.Transaction = transaction;
        command.CommandText = "INSERT INTO Employees (ID, Name, Age, Department) VALUES (@ID, @Name, @Age, @Department)";

        command.Parameters.Add("@ID", SqlDbType.Int);
        command.Parameters.Add("@Name", SqlDbType.NVarChar, 100);
        command.Parameters.Add("@Age", SqlDbType.Int);
        command.Parameters.Add("@Department", SqlDbType.NVarChar, 50);

        try
        {
            foreach (var employee in employees)
            {
                command.Parameters["@ID"].Value = employee.ID;
                command.Parameters["@Name"].Value = employee.Name;
                command.Parameters["@Age"].Value = employee.Age;
                command.Parameters["@Department"].Value = employee.Department;
                command.ExecuteNonQuery();
            }

            transaction.Commit();
        }
        catch
        {
            transaction.Rollback();
            throw;
        }
    }

    public static List<Employee> GetEmployees()
    {
        var employees = new List<Employee>();

        using var connection = new SqlConnection(DatabaseConnectionString);
        connection.Open();

        using var command = connection.CreateCommand();
        command.CommandText = "SELECT ID, Name, Age, Department FROM Employees";

        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            var employee = new Employee
            {
                ID = reader.GetInt32(0),
                Name = reader.GetString(1),
                Age = reader.GetInt32(2),
                Department = reader.IsDBNull(3) ? null : reader.GetString(3)
            };
            employees.Add(employee);
        }

        return employees;
    }
}
