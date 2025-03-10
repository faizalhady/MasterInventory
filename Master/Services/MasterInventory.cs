// using System;
// using System.IO;
// using System.Linq;
// using System.Collections.Generic;
// using ClosedXML.Excel;
// using Microsoft.EntityFrameworkCore;

// public static class Est1cDbProcessor
// {
//     public static void ProcessMasterInventory()
//     {
//         // File path
//         string filePath = @"C:\Users\4033375\Desktop\eST1C\Docs\Excels\est1c chanop.xlsx"; // Change path accordingly

//         // Check if the file exists
//         if (!File.Exists(filePath))
//         {
//             Console.WriteLine($"File not found: {filePath}");
//             return;
//         }

//         // Determine file extension
//         string extension = Path.GetExtension(filePath).ToLowerInvariant();
//         switch (extension)
//         {
//             case ".xlsx":
//             case ".xlsm":
//                 ProcessExcelFile(filePath);
//                 break;
//             case ".csv":
//                 ProcessCsvFile(filePath);
//                 break;
//             default:
//                 Console.WriteLine($"Unsupported file format: {extension}");
//                 break;
//         }
//     }

//     private static void ProcessExcelFile(string filePath)
//     {
//         try
//         {
//             using var workbook = new XLWorkbook(filePath);
//             var worksheet = workbook.Worksheets.First();
//             var headers = worksheet.Row(1).Cells().Select(c => c.GetString().Trim()).ToArray();

//             ProcessFile(headers, rowNumber =>
//             {
//                 var row = worksheet.Row(rowNumber);
//                 return headers.Select(header =>
//                 {
//                     var cellIndex = Array.IndexOf(headers, header) + 1;
//                     var cell = row.Cell(cellIndex);
//                     return cell.IsEmpty() ? null : cell.GetString().Trim();
//                 }).ToArray();
//             }, worksheet.LastRowUsed().RowNumber());
//         }
//         catch (Exception ex)
//         {
//             Console.WriteLine($"Error processing Excel file: {ex.Message}");
//             if (ex.InnerException != null)
//             {
//                 Console.WriteLine($"Inner Exception: {ex.InnerException.Message}");
//             }
//         }
//     }

//     private static void ProcessCsvFile(string filePath)
//     {
//         try
//         {
//             var lines = File.ReadAllLines(filePath);
//             var headers = lines.First().Split(',').Select(h => h.Trim()).ToArray();

//             ProcessFile(headers, rowNumber =>
//             {
//                 var line = lines[rowNumber];
//                 return line.Split(',').Select(value => value.Trim()).ToArray();
//             }, lines.Length - 1);
//         }
//         catch (Exception ex)
//         {
//             Console.WriteLine($"Error processing CSV file: {ex.Message}");
//             if (ex.InnerException != null)
//             {
//                 Console.WriteLine($"Inner Exception: {ex.InnerException.Message}");
//             }
//         }
//     }

//     private static void ProcessFile(string[] headers, Func<int, string[]> getRowValues, int lastRow)
//     {
//         using var dbContext = new MasterDbContext();

//         // Load database rows into a dictionary for efficient access
//         var dbRows = dbContext.MasterInventory.ToDictionary(e => e.Id);

//         // Track IDs from the Excel file
//         var excelIds = new HashSet<int>();

//         for (int rowNumber = 1; rowNumber <= lastRow; rowNumber++)
//         {
//             var values = getRowValues(rowNumber);

//             // Match "No." column with database "Id"
//             var excelNo = GetValue(headers, values, "No.");
//             if (!int.TryParse(excelNo, out int excelId))
//             {
//                 Console.WriteLine($"Skipping row {rowNumber}: Invalid 'No.' value.");
//                 continue;
//             }

//             excelIds.Add(excelId); // Add to Excel ID list

//             if (dbRows.TryGetValue(excelId, out var dbRow))
//             {
//                 // Compare and update existing row
//                 bool rowUpdated = false;

//                 foreach (var header in headers)
//                 {
//                     if (header == "No.") continue; // Skip the "No." column

//                     var columnIndex = Array.IndexOf(headers, header);
//                     var excelValue = values[columnIndex] ?? string.Empty; // Get Excel value
//                     var dbValue = dbRow.GetType().GetProperty(header.Replace(" ", ""))?.GetValue(dbRow)?.ToString()?.Trim() ?? string.Empty; // Get DB value

//                     // Compare normalized values
//                     if (excelValue != dbValue)
//                     {
//                         var property = dbRow.GetType().GetProperty(header.Replace(" ", ""));
//                         if (property != null)
//                         {
//                             property.SetValue(dbRow, string.IsNullOrWhiteSpace(excelValue) ? null : excelValue);
//                             rowUpdated = true;
//                             Console.WriteLine($"Row {rowNumber}, Column '{header}' updated: '{dbValue}' -> '{excelValue}'");
//                         }
//                     }
//                 }

//                 if (rowUpdated)
//                 {
//                     dbContext.Entry(dbRow).State = EntityState.Modified;
//                 }
//             }
//             else
//             {
//                 // Add new row to the database
//                 var newInventory = new MasterInventory
//                 {
//                     Asset = GetValue(headers, values, "Asset"),
//                     Description = GetValue(headers, values, "Description"),
//                     TorqueSerialNumber = GetValue(headers, values, "Torque Serial Number"),
//                     TorqueModel = GetValue(headers, values, "Torque Model"),
//                     TorqueRange = GetValue(headers, values, "Torque Range"),
//                     ControllerSerialNumber = GetValue(headers, values, "Controller Serial Number"),
//                     Brand = GetValue(headers, values, "Brand"),
//                     Manufacturer = GetValue(headers, values, "Manufacturer"),
//                     Supplier = GetValue(headers, values, "Supplier"),
//                     EquipmentID = GetValue(headers, values, "Equipment ID"),
//                     RegisterSAPEAM = GetValue(headers, values, "SAP EAM Number"),
//                     FunctionalGroup = GetValue(headers, values, "Functional Group"),
//                     ProcessGroup = GetValue(headers, values, "Process Group"),
//                     Sector = GetValue(headers, values, "Sector"),
//                     Plant = GetValue(headers, values, "Plant"),
//                     Workcell = GetValue(headers, values, "Workcell"),
//                     CostCenter = GetValue(headers, values, "Cost Center"),
//                     NetBookValue = GetValue(headers, values, "Net Book Value"),
//                     UserStatus = GetValue(headers, values, "User Status"),
//                     KeyInDate = GetValue(headers, values, "Key In Date"),
//                     Owner = GetValue(headers, values, "Owner"),
//                     Comment = GetValue(headers, values, "Comment")
//                 };

//                 dbContext.MasterInventory.Add(newInventory);
//                 Console.WriteLine($"New row added: {excelNo}");
//             }
//         }

//         // Identify rows to delete
//         var idsToDelete = dbRows.Keys.Except(excelIds).ToList();
//         if (idsToDelete.Any())
//         {
//             var rowsToDelete = dbContext.MasterInventory.Where(e => idsToDelete.Contains(e.Id)).ToList();
//             dbContext.MasterInventory.RemoveRange(rowsToDelete);
//             Console.WriteLine($"Deleted {rowsToDelete.Count} rows from the database that were not in the Excel file.");
//         }

//         dbContext.SaveChanges();
//         Console.WriteLine("File processed successfully with updates and deletions.");
//     }

//     private static string GetValue(string[] headers, string[] values, string columnName)
//     {
//         var index = Array.IndexOf(headers, columnName);
//         return index >= 0 && index < values.Length ? values[index] ?? string.Empty : string.Empty;
//     }
// }
// nk mai