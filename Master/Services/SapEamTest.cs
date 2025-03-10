using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Text;
using ClosedXML.Excel;
using Microsoft.EntityFrameworkCore;
using Serilog;
using CratePilotSystemWS.Services.Email; // Ensure this namespace matches yours

public static class SapEamProcessorTest
{
    // Maps spreadsheet column headers to your SapEam model's property names
    private static readonly Dictionary<string, string> headerToPropertyMap = new()
    {
        { "Equipment", "Equipment" },
        { "Asset", "Asset" },
        { "Manuf. SerialNr", "ManufSerialNo" },
        { "Sort Field", "SortField" },
        { "Description", "Description" },
        { "Location", "Location" },
        { "Room", "Room" },
        { "Work Center", "WorkCenter" },
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
        { "Gross Weight", "GrossWeight" },
        { "Acquistion date", "AcquisitionDate" },
        { "AcquistnValue", "AcquisitionValue" },
        { "Unit of Weight", "WeightUnit" },
        { "Size/dimens.", "SizeDimensions" },
        { "Main WorkCtr", "MainWorkCtr" },
        { "Mfr Ctry/Reg", "ManufCountry" },
        { "Planner Group", "PlannerGroup" },
        { "Serial Number", "SerialNumber" },
        { "Catalog Profile", "CatalogProfile" },
        { "User status", "UserStatus" },
        { "System status", "SystemStatus" }
    };

    public static void Process()
    {
        Log.Logger = new LoggerConfiguration()
            .WriteTo.Console()
            .WriteTo.File(@"\\mypenm0opsapp01\SmartTorque$\est1c\Testing - Syed\SapLog.txt", rollingInterval: RollingInterval.Day)
            .CreateLogger();

        Log.Information("Application started.");

        string directoryPath = @"\\mypenm0opsapp01\SmartTorque$\est1c\"; 
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

        // Select the latest file based on last write time
        var latestFile = files.OrderByDescending(File.GetLastWriteTime).First();

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
            var worksheet = workbook.Worksheets
                .FirstOrDefault(ws => possibleSheetNames.Contains(ws.Name, StringComparer.OrdinalIgnoreCase));

            if (worksheet == null)
            {
                Console.WriteLine("Worksheet 'EST1C/eST1C/est1c' not found in the file.");
                Log.Error("Worksheet 'EST1C/eST1C/est1c' not found in the file.");
                return;
            }

            var headers = worksheet.Row(1).Cells().Select(c => c.GetString().Trim()).ToArray();

            ProcessFile(
                headers,
                rowNumber =>
                {
                    var row = worksheet.Row(rowNumber);
                    return headers.Select(header =>
                    {
                        var cellIndex = Array.IndexOf(headers, header) + 1;
                        var cell = row.Cell(cellIndex);
                        return cell.IsEmpty() ? null : cell.GetString().Trim();
                    }).ToArray();
                },
                worksheet.LastRowUsed().RowNumber()
            );
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing Excel file: {ex.Message}");
            Log.Error(ex, $"Error processing Excel file: {filePath}");
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

            ProcessFile(
                headers,
                rowNumber =>
                {
                    var line = lines[rowNumber];
                    return line.Split(',').Select(value => value.Trim()).ToArray();
                },
                lines.Length - 1
            );
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

    /// <summary>
    /// Processes file rows to perform Create/Update/Delete operations,
    /// then builds and emails an HTML report with the CRUD logs.
    /// </summary>
    private static void ProcessFile(string[] headers, Func<int, string[]> getRowValues, int lastRow)
    {
        using var dbContext = new MasterDbContext();

        // Build a lookup of MainWorkCtr from existing records (for deleted rows)
        var dbRows = dbContext.SapEam.OrderBy(e => e.Id).ToList();
        var initialMainWorkCtrLookup = dbRows.ToDictionary(r => r.ManufSerialNo ?? "", r => r.MainWorkCtr);
        var dbKeys = dbRows.Select(r => r.ManufSerialNo).ToHashSet();

        // Collect keys from file rows
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

        // Generate a new GroupId for this processing run
        int groupId = dbContext.ChangeLogs.Any()
            ? dbContext.ChangeLogs.Max(c => c.GroupId) + 1
            : 1;

        Log.Information($"Processing started. Group ID: {groupId}");

        // -- Process Deleted Rows --
        var deletedKeys = dbKeys.Except(fileKeys);
        foreach (var deletedKey in deletedKeys)
        {
            var dbRow = dbRows.First(r => r.ManufSerialNo == deletedKey);

            // Always log the deletion
            dbContext.ChangeLogs.Add(new ChangeLogs
            {
                GroupId = groupId,
                ManufSerialNo = dbRow.ManufSerialNo,
                ColumnName = "N/A",
                OldValue = "Row existed",
                NewValue = "Row deleted",
                ChangeType = "Delete",
                Timestamp = DateTime.UtcNow
            });

            Console.WriteLine($"Deleted row with Serial Number: {deletedKey}");
            dbContext.SapEam.Remove(dbRow);
        }

        // -- Process Create or Update Rows --
        for (int rowNumber = 2; rowNumber <= lastRow; rowNumber++)
        {
            var values = getRowValues(rowNumber);
            if (values.All(string.IsNullOrWhiteSpace)) continue;

            var keyIndex = Array.IndexOf(headers, "Manuf. SerialNr");
            var fileKey = (keyIndex >= 0 && keyIndex < values.Length)
                ? values[keyIndex]?.Trim()
                : null;

            if (string.IsNullOrWhiteSpace(fileKey)) continue;

            if (dbKeys.Contains(fileKey))
            {
                // UPDATE
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
                // CREATE
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

                // Always log the creation
                dbContext.ChangeLogs.Add(new ChangeLogs
                {
                    GroupId = groupId,
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

        // Save all changes (updates, creations, deletions, and logs)
        dbContext.SaveChanges();
        Log.Information("Database synchronization completed.");
        Console.WriteLine("File processed successfully.");

        // After changes, build a lookup for current records (for Create and Update)
        var currentMainWorkCtrLookup = dbContext.SapEam.ToDictionary(r => r.ManufSerialNo ?? "", r => r.MainWorkCtr);

        // Retrieve all CRUD logs for this run, ordered by time
        var crudLogs = dbContext.ChangeLogs
            .Where(log => log.GroupId == groupId)
            .OrderBy(l => l.Timestamp)
            .ToList();

        // If no logs were recorded, send a simple up-to-date email.
        if (!crudLogs.Any())
        {
            var noChangesResult = EmailServices.SendCrudLogEmail("The database is already up to date.")
                                               .GetAwaiter().GetResult();
            Console.WriteLine(noChangesResult.msg);
            return;
        }

        // Group logs by type
        var createdLogs = crudLogs.Where(x => x.ChangeType == "Create").ToList();
        var updatedLogs = crudLogs.Where(x => x.ChangeType == "Update").ToList();
        var deletedLogs = crudLogs.Where(x => x.ChangeType == "Delete").ToList();

        // --- Build HTML Tables with "Main WorkCtr" column ---
        string BuildCreateTable(List<ChangeLogs> logs)
        {
            if (!logs.Any()) return string.Empty;
            var sb = new StringBuilder();
            sb.AppendLine("<h3 style='color: #2c3e50; margin-top:30px;'>Created Items</h3>");
            sb.AppendLine("<table style='border-collapse: collapse; width: 100%;'>");
            sb.AppendLine("<thead>");
            sb.AppendLine("<tr>");
            sb.AppendLine("<th style='border:1px solid #ccc; padding:5px;'>Serial Number</th>");
            sb.AppendLine("<th style='border:1px solid #ccc; padding:5px;'>Main WorkCtr</th>");
            sb.AppendLine("</tr>");
            sb.AppendLine("</thead>");
            sb.AppendLine("<tbody>");
            foreach (var log in logs)
            {
                // For created items, look up in current records
                string mainWorkCtr = currentMainWorkCtrLookup.TryGetValue(log.ManufSerialNo, out var value) ? value : "N/A";
                sb.AppendLine("<tr>");
                sb.AppendLine($"<td style='border:1px solid #ccc; padding:5px;'>{log.ManufSerialNo}</td>");
                sb.AppendLine($"<td style='border:1px solid #ccc; padding:5px;'>{mainWorkCtr}</td>");
                sb.AppendLine("</tr>");
            }
            sb.AppendLine("</tbody>");
            sb.AppendLine("</table>");
            return sb.ToString();
        }

        string BuildUpdateTable(List<ChangeLogs> logs)
        {
            if (!logs.Any()) return string.Empty;
            var sb = new StringBuilder();
            sb.AppendLine("<h3 style='color: #2c3e50; margin-top:30px;'>Updated Items</h3>");
            sb.AppendLine("<table style='border-collapse: collapse; width: 100%;'>");
            sb.AppendLine("<thead>");
            sb.AppendLine("<tr>");
            sb.AppendLine("<th style='border:1px solid #ccc; padding:5px;'>Serial</th>");
            sb.AppendLine("<th style='border:1px solid #ccc; padding:5px;'>Column</th>");
            sb.AppendLine("<th style='border:1px solid #ccc; padding:5px;'>Old Value</th>");
            sb.AppendLine("<th style='border:1px solid #ccc; padding:5px;'>New Value</th>");
            sb.AppendLine("<th style='border:1px solid #ccc; padding:5px;'>Main WorkCtr</th>");
            sb.AppendLine("</tr>");
            sb.AppendLine("</thead>");
            sb.AppendLine("<tbody>");
            foreach (var log in logs)
            {
                // For updates, use current records
                string mainWorkCtr = currentMainWorkCtrLookup.TryGetValue(log.ManufSerialNo, out var value) ? value : "N/A";
                sb.AppendLine("<tr>");
                sb.AppendLine($"<td style='border:1px solid #ccc; padding:5px;'>{log.ManufSerialNo}</td>");
                sb.AppendLine($"<td style='border:1px solid #ccc; padding:5px;'>{log.ColumnName}</td>");
                sb.AppendLine($"<td style='border:1px solid #ccc; padding:5px;'>{log.OldValue}</td>");
                sb.AppendLine($"<td style='border:1px solid #ccc; padding:5px;'>{log.NewValue}</td>");
                sb.AppendLine($"<td style='border:1px solid #ccc; padding:5px;'>{mainWorkCtr}</td>");
                sb.AppendLine("</tr>");
            }
            sb.AppendLine("</tbody>");
            sb.AppendLine("</table>");
            return sb.ToString();
        }

        string BuildDeleteTable(List<ChangeLogs> logs)
        {
            if (!logs.Any()) return string.Empty;
            var sb = new StringBuilder();
            sb.AppendLine("<h3 style='color: #2c3e50; margin-top:30px;'>Deleted Items</h3>");
            sb.AppendLine("<table style='border-collapse: collapse; width: 100%;'>");
            sb.AppendLine("<thead>");
            sb.AppendLine("<tr>");
            sb.AppendLine("<th style='border:1px solid #ccc; padding:5px;'>Serial Number</th>");
            sb.AppendLine("<th style='border:1px solid #ccc; padding:5px;'>Main WorkCtr</th>");
            sb.AppendLine("</tr>");
            sb.AppendLine("</thead>");
            sb.AppendLine("<tbody>");
            foreach (var log in logs)
            {
                // For deleted items, look up in the initial dictionary
                string mainWorkCtr = initialMainWorkCtrLookup.TryGetValue(log.ManufSerialNo, out var value) ? value : "N/A";
                sb.AppendLine("<tr>");
                sb.AppendLine($"<td style='border:1px solid #ccc; padding:5px;'>{log.ManufSerialNo}</td>");
                sb.AppendLine($"<td style='border:1px solid #ccc; padding:5px;'>{mainWorkCtr}</td>");
                sb.AppendLine("</tr>");
            }
            sb.AppendLine("</tbody>");
            sb.AppendLine("</table>");
            return sb.ToString();
        }

        // Combine all tables into one HTML content string.
        var htmlBuilder = new StringBuilder();
        htmlBuilder.AppendLine("<div style='font-size:14px;'>");
        htmlBuilder.AppendLine(BuildCreateTable(createdLogs));
        htmlBuilder.AppendLine(BuildUpdateTable(updatedLogs));
        htmlBuilder.AppendLine(BuildDeleteTable(deletedLogs));
        htmlBuilder.AppendLine("</div>");

        var emailContent = htmlBuilder.ToString();
        var emailResult = EmailServices.SendCrudLogEmail(emailContent).GetAwaiter().GetResult();
        Console.WriteLine(emailResult.msg);
    }
}
