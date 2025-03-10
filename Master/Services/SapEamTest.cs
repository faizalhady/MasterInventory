using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using ClosedXML.Excel;
using Microsoft.EntityFrameworkCore;
using Serilog;

public static class SapEamProcessorTest
{
    public static void Process()
    {
         
        Log.Logger = new LoggerConfiguration()
            .WriteTo.Console()
            .WriteTo.File(@"\\mypenm0opsapp01\SmartTorque$\est1c\Testing - Syed\SapLog.txt", rollingInterval: RollingInterval.Day)
            .CreateLogger();

        Log.Information("Application started.");

        string directoryPath = @"\\mypenm0opsapp01\SmartTorque$\est1c"; // Change path accordingly
        // string directoryPath = @"\\mypenm0opsapp01\SmartTorque$\est1c\Testing - Syed"; // testting path
        if (!Directory.Exists(directoryPath))
        {
            Console.WriteLine($"Directory not found: {directoryPath}");
            Log.Error($"Directory not found: {directoryPath}");
            return;
        }

        var files = Directory.GetFiles(directoryPath, "*.*")
            .Where(file => file.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                           file.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase) ||
                           file.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (!files.Any())
        {
            Console.WriteLine($"No Excel or CSV files found in the directory: {directoryPath}");
            return;
        }

        // Select the latest file based on creation or modification date
        var latestFile = files.OrderByDescending(File.GetLastWriteTime).First(); // Use GetCreationTime if needed

        Log.Information($"Processing latest file: {Path.GetFileName(latestFile)}");
        Console.WriteLine($"Processing latest file: {Path.GetFileName(latestFile)}");

        string extension = Path.GetExtension(latestFile).ToLowerInvariant();
        try
        {
            switch (extension)
            {
                case ".xlsx":
                case ".xlsm":
                    ProcessExcelFile(latestFile);
                    break;
                case ".csv":
                    ProcessCsvFile(latestFile);
                    break;
                default:
                    Log.Warning($"Unsupported file format: {extension}");
                    Console.WriteLine($"Unsupported file format: {extension}");
                    break;
            }

            Log.Information("Processing completed successfully.");
            Console.WriteLine("Processing completed successfully.");
        }
        catch (Exception ex)
        {
            Log.Error(ex, "An error occurred during processing.");
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    private static void ProcessExcelFile(string filePath)
    {
        try
        {
            using var workbook = new XLWorkbook(filePath);
            
           var possibleSheetNames = new List<string> { "EST1C", "eST1C", "est1c" };

            var worksheet = workbook.Worksheets.FirstOrDefault(ws => possibleSheetNames.Contains(ws.Name, StringComparer.OrdinalIgnoreCase));

            if (worksheet == null)
            {
                Console.WriteLine("Worksheet with names 'EST1C', 'eSt1c', or 'est1c' not found in the file.");
                Log.Error("Worksheet with names 'EST1C', 'eSt1c', or 'est1c' not found in the file.");
                return;
            }



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
            Log.Error(ex, $"Error processing CSV file: {filePath}");
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
            Log.Error(ex, $"Error processing CSV file: {filePath}");
            if (ex.InnerException != null)
            {
                Console.WriteLine($"Inner Exception: {ex.InnerException.Message}");
            }
        }
    }

   private static void ProcessFile(string[] headers, Func<int, string[]> getRowValues, int lastRow)
{
    using var dbContext = new MasterDbContext();

    // Generate a new GroupId by checking the latest in the ChangeLogs table
    int groupId = dbContext.ChangeLogs.Any() 
        ? dbContext.ChangeLogs.Max(c => c.GroupId) + 1 
        : 1;

    Log.Information($"Processing started. Group ID: {groupId}");

    // 🔹 Load existing logs (prevent duplicate tracking)
   var existingLogs = dbContext.ChangeLogs
    .Select(log => new ChangeLogs
    {
        GroupId = log.GroupId,
        ManufSerialNo = log.ManufSerialNo ?? string.Empty, // Handle NULL
        ColumnName = log.ColumnName ?? "N/A", // Handle NULL
        OldValue = log.OldValue ?? "N/A", // Handle NULL
        NewValue = log.NewValue ?? "N/A", // Handle NULL
        ChangeType = log.ChangeType ?? "Unknown", // Handle NULL
        Timestamp = log.Timestamp
    })
    .AsNoTracking()
    .ToList();
  

    var dbRows = dbContext.SapEam.OrderBy(e => e.Id).ToList();
    var dbKeys = dbRows.Select(r => r.ManufSerialNo).ToHashSet();

    var fileKeys = new HashSet<string>();
    for (int rowNumber = 2; rowNumber <= lastRow; rowNumber++)
    {
        var values = getRowValues(rowNumber);
        if (values.All(string.IsNullOrWhiteSpace)) continue;

        var keyIndex = Array.IndexOf(headers, "Manuf. SerialNr");
        if (keyIndex >= 0 && keyIndex < values.Length)
        {
            var fileKey = values[keyIndex]?.Trim();
            if (!string.IsNullOrWhiteSpace(fileKey))
            {
                fileKeys.Add(fileKey);
            }
        }
    }

    var deletedKeys = dbKeys.Except(fileKeys);

    if (!deletedKeys.Any() && !fileKeys.Except(dbKeys).Any())
{
    Console.WriteLine("✅ No changes detected. Database is already up to date.");
    Log.Information("✅ No changes detected. Database is already up to date.");
    return; // Exit early
}


    // 🔹 Log deleted rows
    foreach (var deletedKey in deletedKeys)
    {
        var dbRow = dbRows.First(r => r.ManufSerialNo == deletedKey);

        if (!existingLogs.Any(log => log.ManufSerialNo == dbRow.ManufSerialNo && log.ChangeType == "Delete"))
        {
            dbContext.ChangeLogs.Add(new ChangeLogs
            {
                GroupId = groupId,  // 🔥 Use same GroupId for grouping
                ManufSerialNo = dbRow.ManufSerialNo,
                ColumnName = "N/A",
                OldValue = "Row existed",
                NewValue = "Row deleted",
                ChangeType = "Delete",
                Timestamp = DateTime.UtcNow
            });
        }

        Console.WriteLine($"Deleted row with Serial Number: {deletedKey}");
        dbContext.SapEam.Remove(dbRow);
    }

    // 🔹 Process updates or inserts
    for (int rowNumber = 2; rowNumber <= lastRow; rowNumber++)
    {
        var values = getRowValues(rowNumber);
        if (values.All(string.IsNullOrWhiteSpace)) continue;

        var keyIndex = Array.IndexOf(headers, "Manuf. SerialNr");
        var fileKey = keyIndex >= 0 && keyIndex < values.Length ? values[keyIndex]?.Trim() : null;

        if (!string.IsNullOrWhiteSpace(fileKey) && dbKeys.Contains(fileKey))
        {
            var dbRow = dbRows.First(r => r.ManufSerialNo == fileKey);
            bool isUpdated = false;

           foreach (var header in headers)
{
    if (!headerToPropertyMap.TryGetValue(header, out var propertyName)) continue;

    var property = dbRow.GetType().GetProperty(propertyName);
    if (property == null) continue;

    var columnIndex = Array.IndexOf(headers, header);
    var excelValue = values[columnIndex] ?? string.Empty;
    var dbValue = property.GetValue(dbRow)?.ToString()?.Trim() ?? string.Empty;

    

    if (!string.Equals(excelValue, dbValue, StringComparison.OrdinalIgnoreCase))
    {
        Console.WriteLine($"✅ LOGGING CHANGE: {fileKey}, Column '{header}', Old: '{dbValue}', New: '{excelValue}'");
        Log.Information($"✅ LOGGING CHANGE: {fileKey}, Column '{header}', Old: '{dbValue}', New: '{excelValue}'");
        

        dbContext.ChangeLogs.Add(new ChangeLogs
        {
            GroupId = groupId,
            ManufSerialNo = dbRow.ManufSerialNo,
            ColumnName = propertyName,
            OldValue = dbValue,
            NewValue = string.IsNullOrWhiteSpace(excelValue) ? null : excelValue,
            ChangeType = "Update",
            Timestamp = DateTime.UtcNow
        });

        property.SetValue(dbRow, string.IsNullOrWhiteSpace(excelValue) ? null : excelValue);

        // 🔹 Force EF to recognize the change
        dbContext.Entry(dbRow).Property(propertyName).IsModified = true;

        isUpdated = true;
    }

}
    


            if (isUpdated)
            {
                dbContext.Entry(dbRow).State = EntityState.Modified;
            }
        }
        else
        {
            var newInventory = new SapEam();

            foreach (var header in headers)
            {
                if (!headerToPropertyMap.TryGetValue(header, out var propertyName)) continue;

                var property = newInventory.GetType().GetProperty(propertyName);
                if (property == null) continue;

                var columnIndex = Array.IndexOf(headers, header);
                var value = columnIndex >= 0 && columnIndex < values.Length ? values[columnIndex] : null;
                property.SetValue(newInventory, string.IsNullOrWhiteSpace(value) ? null : value.Trim());
            }

            dbContext.SapEam.Add(newInventory);
            dbContext.SaveChanges();

            if (!existingLogs.Any(log => log.ManufSerialNo == newInventory.ManufSerialNo && log.ChangeType == "Create"))
            {
                dbContext.ChangeLogs.Add(new ChangeLogs
                {
                    GroupId = groupId,  // 🔥 Group new rows under same GroupId
                    ManufSerialNo = newInventory.ManufSerialNo,
                    ColumnName = "N/A",
                    OldValue = newInventory.ManufSerialNo,
                    NewValue = newInventory.ManufSerialNo,
                    ChangeType = "Create",
                    Timestamp = DateTime.UtcNow
                });

                Console.WriteLine($"Inserted new row with Serial Number: {fileKey}");
            }
        }
    }

    dbContext.SaveChanges();
    Log.Information("Database synchronization completed.");
    Console.WriteLine("File processed successfully.");
}



    private static readonly Dictionary<string, string> headerToPropertyMap = new()
{
    { "Equipment", "Equipment" },
    { "Asset", "Asset" },
    { "Manuf. SerialNr", "ManufSerialNo" }, // Updated key to match the new header
    { "Sort Field", "SortField" },
    { "Description", "Description" },
    { "Location", "Location" },
    { "Room", "Room" },
    { "Work Center", "WorkCenter" }, // Updated key to match the new header
    { "Position", "Position" },
    { "Object Type", "ObjectType" },
    { "Model number", "ModelNumber" },
    { "Cost Center", "CostCenter" },
    { "Company Code", "CompanyCode" },
    { "Planning Plant", "PlanningPlant" },
    { "MaintPlant", "MaintPlant" },
    { "Plant Section", "PlantSection" },
    { "EquipCategory", "EquipCategory" },
    { "Manufacturer", "Manufacturer" },
    { "ManufPartNo.", "ManufPartNo" },
    { "Functional Loc.", "FunctionalLoc" },
    { "ABC Indic.", "ABCIndic" },
    { "Sub-number", "SubNumber" },
    { "ConstructYear", "ConstructYear" },
    { "ConstructMth", "ConstructMth" },
    { "Gross Weight", "GrossWeight" }, // Updated key to match the new header
    { "Acquistion date", "AcquisitionDate" },
    { "AcquistnValue", "AcquisitionValue" },
    { "Unit of Weight", "WeightUnit" }, // Updated key to match the new header
    { "Size/dimens.", "SizeDimensions" },
    { "Main WorkCtr", "MainWorkCtr" },
    { "Mfr Ctry/Reg", "ManufCountry" }, // Updated key to match the new header
    { "Planner Group", "PlannerGroup" },
    { "Serial Number", "SerialNumber" },
    { "Catalog Profile", "CatalogProfile" },
    { "User status", "UserStatus" },
    { "System status", "SystemStatus" }
};
}
