using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using ClosedXML.Excel;
using Microsoft.EntityFrameworkCore;

public static class SapEamProcessor
{
    public static void ProcessSapEam()
    {
        // string filePath = @"\\mypenm0opsapp01\SmartTorque$\est1c\SAPEAM_EST1C_122624.xlsx"; // Change path accordingly
        string filePath = @"C:\Users\4033375\Desktop\csv test\SAPEAM_1.xlsx"; // Change path accordingly

        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        string extension = Path.GetExtension(filePath).ToLowerInvariant();
        switch (extension)
        {
            case ".xlsx":
            case ".xlsm":
                ProcessExcelFile(filePath);
                break;
            case ".csv":
                ProcessCsvFile(filePath);
                break;
            default:
                Console.WriteLine($"Unsupported file format: {extension}");
                break;
        }
    }

    private static void ProcessExcelFile(string filePath)
    {
        try
        {
            using var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheets.First();
            var headers = worksheet.Row(1).Cells().Select(c => c.GetString().Trim()).ToArray();

            ProcessFile(headers, rowNumber =>
            {
                var row = worksheet.Row(rowNumber);
                return headers.Select(header =>
                {
                    var cellIndex = Array.IndexOf(headers, header) + 1;
                    var cell = row.Cell(cellIndex);
                    return cell.IsEmpty() ? null : cell.GetString().Trim();
                }).ToArray();
            }, worksheet.LastRowUsed().RowNumber());
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing Excel file: {ex.Message}");
            if (ex.InnerException != null)
            {
                Console.WriteLine($"Inner Exception: {ex.InnerException.Message}");
            }
        }
    }

    private static void ProcessCsvFile(string filePath)
    {
        try
        {
            var lines = File.ReadAllLines(filePath);
            var headers = lines.First().Split(',').Select(h => h.Trim()).ToArray();

            ProcessFile(headers, rowNumber =>
            {
                var line = lines[rowNumber];
                return line.Split(',').Select(value => value.Trim()).ToArray();
            }, lines.Length - 1);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing CSV file: {ex.Message}");
            if (ex.InnerException != null)
            {
                Console.WriteLine($"Inner Exception: {ex.InnerException.Message}");
            }
        }
    }

    private static void ProcessFile(string[] headers, Func<int, string[]> getRowValues, int lastRow)
    {
        using var dbContext = new MasterDbContext();

        var dbRows = dbContext.SapEam.OrderBy(e => e.Id).ToList(); // Fetch rows from the database
        var dbRowCount = dbRows.Count;
        var excelRowCount = lastRow - 1; // Excel rows minus the header row

        // Process each row in the Excel file
        for (int rowNumber = 2; rowNumber <= lastRow; rowNumber++) // Start from row 2 (assuming row 1 is the header)
        {
            var values = getRowValues(rowNumber);
            int dbRowIndex = rowNumber - 2; // Match row index in the database

            if (dbRowIndex < dbRowCount)
            {
                // Update existing row in the database if values differ
                var dbRow = dbRows[dbRowIndex];
                foreach (var header in headers)
                {
                    var columnIndex = Array.IndexOf(headers, header);
                    var excelValue = values[columnIndex] ?? string.Empty; // Get Excel value
                    var property = dbRow.GetType().GetProperty(header.Replace(" ", "").Replace(".", ""));
                    if (property != null)
                    {
                        var dbValue = property.GetValue(dbRow)?.ToString()?.Trim() ?? string.Empty;

                        if (excelValue != dbValue)
                        {
                            property.SetValue(dbRow, string.IsNullOrWhiteSpace(excelValue) ? null : excelValue);
                            dbContext.Entry(dbRow).State = EntityState.Modified;
                            Console.WriteLine($"Row {rowNumber}, Column '{header}' updated: '{dbValue}' -> '{excelValue}'");
                        }
                    }
                }
            }
            else
            {
                // Add new row to the database
                var newInventory = new SapEam
                {
                    Equipment = GetValue(headers, values, "Equipment"),
                    Asset = GetValue(headers, values, "Asset"),
                    ManufSerialNo = GetValue(headers, values, "ManufSerialNo."),
                    SortField = GetValue(headers, values, "Sort Field"),
                    Description = GetValue(headers, values, "Description"),
                    Location = GetValue(headers, values, "Location"),
                    Room = GetValue(headers, values, "Room"),
                    WorkCenter = GetValue(headers, values, "Work center"),
                    Position = GetValue(headers, values, "Position"),
                    ObjectType = GetValue(headers, values, "Object Type"),
                    ModelNumber = GetValue(headers, values, "Model number"),
                    CostCenter = GetValue(headers, values, "Cost Center"),
                    CompanyCode = GetValue(headers, values, "Company Code"),
                    PlanningPlant = GetValue(headers, values, "Planning Plant"),
                    MaintPlant = GetValue(headers, values, "MaintPlant"),
                    PlantSection = GetValue(headers, values, "Plant Section"),
                    Manufacturer = GetValue(headers, values, "Manufacturer"),
                    ManufPartNo = GetValue(headers, values, "ManufPartNo."),
                    FunctionalLoc = GetValue(headers, values, "Functional Loc."),
                    ABCIndic = GetValue(headers, values, "ABC Indic."),
                    SubNumber = GetValue(headers, values, "Sub-number"),
                    ConstructYear = GetValue(headers, values, "ConstructYear"),
                    ConstructMth = GetValue(headers, values, "ConstructMth"),
                    GrossWeight = GetValue(headers, values, "Gross weight"),
                    AcquisitionDate = GetValue(headers, values, "Acquistion date"),
                    AcquisitionValue = GetValue(headers, values, "AcquistnValue"),
                    WeightUnit = GetValue(headers, values, "Weight unit"),
                    SizeDimensions = GetValue(headers, values, "Size/dimens."),
                    MainWorkCtr = GetValue(headers, values, "Main WorkCtr"),
                    ManufCountry = GetValue(headers, values, "ManufCountry"),
                    PlannerGroup = GetValue(headers, values, "Planner Group"),
                    SerialNumber = GetValue(headers, values, "Serial Number"),
                    CatalogProfile = GetValue(headers, values, "Catalog Profile"),
                    UserStatus = GetValue(headers, values, "User status"),
                    SystemStatus = GetValue(headers, values, "System status")
                };

                dbContext.SapEam.Add(newInventory);
                Console.WriteLine($"New row added: Row {rowNumber}");
            }
        }

        // Delete rows from the database if it has more rows than the Excel file
        if (dbRowCount > excelRowCount)
        {
            var rowsToDelete = dbRows.Skip(excelRowCount).ToList();
            dbContext.SapEam.RemoveRange(rowsToDelete);
            Console.WriteLine($"Deleted {rowsToDelete.Count} extra rows from the database.");
        }

        dbContext.SaveChanges();
        Console.WriteLine("File processed successfully with updates, additions, and deletions.");
    }

    private static string GetValue(string[] headers, string[] values, string columnName)
    {
        var index = Array.IndexOf(headers, columnName);
        return index >= 0 && index < values.Length ? values[index] ?? string.Empty : string.Empty;
    }
}
