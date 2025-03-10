// using System;
// using System.IO;
// using System.Linq;
// using System.Collections.Generic;
// using ClosedXML.Excel;
// using Microsoft.EntityFrameworkCore;
// {

// public static class CycleCountProcessor
//     public static void ProcessCycleCount()
//     {
//         // File path
//         string filePath = @"C:\Users\4033375\Desktop\eST1C\Docs\Excels\SAP EAM_Dec 2024_Cycle Count.xlsx";

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
//         }
//     }

//     private static void ProcessFile(string[] headers, Func<int, string[]> getRowValues, int lastRow)
//     {
//         using var dbContext = new MasterDbContext();

//         // Load database rows into a dictionary for efficient access
//         var dbRows = dbContext.CycleCount.ToDictionary(e => e.Id);

//         // Track IDs from Excel
//         var excelIds = new HashSet<int>();

//         for (int rowNumber = 2; rowNumber <= lastRow; rowNumber++) // Skip header row
//         {
//             var values = getRowValues(rowNumber);

//             // Match database row using row number
//             var excelId = rowNumber - 1; // Adjust row number to match DB ID (row 2 in Excel = ID 1 in DB)
//             excelIds.Add(excelId);

//             if (dbRows.TryGetValue(excelId, out var dbRow))
//             {
//                 // Update existing row
//                 bool rowUpdated = false;

//                 foreach (var header in headers)
//                 {
//                     var columnIndex = Array.IndexOf(headers, header);
//                     var excelValue = values[columnIndex] ?? string.Empty;
//                     var property = dbRow.GetType().GetProperty(header.Replace(" ", "").Replace("'", ""));
//                     if (property != null)
//                     {
//                         var dbValue = property.GetValue(dbRow)?.ToString()?.Trim() ?? string.Empty;

//                         if (excelValue != dbValue)
//                         {
//                             property.SetValue(dbRow, string.IsNullOrWhiteSpace(excelValue) ? null : excelValue);
//                             rowUpdated = true;
//                             Console.WriteLine($"Row {excelId}, Column '{header}' updated: '{dbValue}' -> '{excelValue}'");
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
//                 // Add new row to database
//                 var newRow = new CycleCount
//                 {
//                     ManufSerialNo = GetValue(headers, values, "ManufSerialNo."),
//                     MainWorkCtr = GetValue(headers, values, "Main WorkCtr"),
//                     January2024 = GetValue(headers, values, "January'2024"),
//                     February2024 = GetValue(headers, values, "February'2024"),
//                     March2024 = GetValue(headers, values, "March'2024"),
//                     April2024 = GetValue(headers, values, "April'2024"),
//                     May2024 = GetValue(headers, values, "May'2024"),
//                     June2024 = GetValue(headers, values, "June'2024"),
//                     July2024 = GetValue(headers, values, "July'2024"),
//                     August2024 = GetValue(headers, values, "August'2024"),
//                     September2024 = GetValue(headers, values, "September'2024"),
//                     October2024 = GetValue(headers, values, "October'2024"),
//                     November2024 = GetValue(headers, values, "November'2024"),
//                     December2024 = GetValue(headers, values, "December'2024"),
//                     Sum = GetValue(headers, values, "Sum")
//                 };

//                 dbContext.CycleCount.Add(newRow);
//                 Console.WriteLine($"New row added: Row {excelId}");
//             }
//         }

//         // Identify rows to delete
//         var idsToDelete = dbRows.Keys.Except(excelIds).ToList();
//         if (idsToDelete.Any())
//         {
//             var rowsToDelete = dbContext.CycleCount.Where(e => idsToDelete.Contains(e.Id)).ToList();
//             dbContext.CycleCount.RemoveRange(rowsToDelete);
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
