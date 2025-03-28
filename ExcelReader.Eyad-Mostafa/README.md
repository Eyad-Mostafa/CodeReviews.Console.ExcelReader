# 🚀 Excel to Database Importer

This C# application reads employee data from an **Excel (.xlsx) file** and imports it into a **SQL Server database**.
It uses the **EPPlus** library for Excel file handling and **ADO.NET** for database operations.

## 📂 Project Structure

- **FileReader.cs** → Reads employee data from an Excel file.
- **DatabaseManager.cs** → Creates the database and saves employees.
- **Program.cs** → Runs the entire process and provides console feedback.

---

## 🛠️ Technologies Used

- **C# .NET 8**
- **EPPlus** (for reading Excel files)
- **SQL Server**
- **ADO.NET** (for database interaction)

---

## 👅 Setup & Installation

### 1️⃣ **Clone the repository**

```sh
git clone https://github.com/Eyad-Mostafa/ExcelReader.git
cd ExcelReader
```

### 2️⃣ **Place the employees Excel file**

Ensure your **`employees.xlsx`** file is inside the following directory:

```sh
{YourProjectDirectory}\bin\Debug\net8.0\employees.xlsx
```

### 3️⃣ **Run the project**

- Open the project in **Visual Studio**.
- Set the **startup project** to `Excel_Reader`.
- Run the project (`F5`).

---

## 📝 Usage

1. **Reads the Excel file** and extracts employee data.
2. **Deletes and recreates the database** before inserting new data.
3. **Prints progress messages** with colors in the console.
4. **Displays the imported data** after saving it to the database.

---

## ⚠️ Important Notes

- The application **automatically deletes and recreates the database** each time it runs.
- Ensure **SQL Server is installed and running** on your machine.
- You **don't need user input**—the process is fully automated.
