using EFCore.BulkExtensions;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.IO;

namespace Master.Services
{
    public static class CsvFileProcessor2
    {
        private const int batchSize = 1000;

        public static void CompareAndProcessCsvFile()
        {
            string? filePath = @"C:\Users\4033375\Desktop\eST1C\Docs\Excels\master_inventory\Inventory_Database_11222024.csv";
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                Console.WriteLine($"File does not exist or invalid: {filePath}");
                return;
            }

            Console.WriteLine($"Processing file and comparing with database: {filePath}");

            try
            {
                using (var context = new MasterDbContext())
                {
                    var dbData = context.MasterInventory.ToList(); // Load all database rows
                    var dbRowKeys = new HashSet<string>(context.MasterInventory.Select(x => x.Asset)); // Use Asset as the unique key for comparison
                    var updates = new List<MasterInventory>();
                    var newRows = new List<MasterInventory>();

                    using (var reader = new StreamReader(filePath))
                    {
                        string? line = reader.ReadLine();

                        // Read headers
                        if (line == null)
                        {
                            Console.WriteLine("The file is empty.");
                            return;
                        }
                        var headers = line.Split(',');

                        while ((line = reader.ReadLine()) != null)
                        {
                            var values = line.Split(',');

                            // Build a unique key from the Asset column or other fields
                            var uniqueKey = GetValue(headers, values, "Asset");
                            if (string.IsNullOrWhiteSpace(uniqueKey))
                            {
                                Console.WriteLine("Skipping row with empty Asset value.");
                                continue;
                            }

                            // Check if this row already exists in the database
                            var existingRow = dbData.FirstOrDefault(x => x.Asset == uniqueKey);
                            if (existingRow != null) // Update existing rows
                            {
                                bool hasChanges = false;

                                for (int colIndex = 0; colIndex < headers.Length; colIndex++)
                                {
                                    var columnName = headers[colIndex].Trim();
                                    var newValue = values[colIndex].Trim();
                                    var property = typeof(MasterInventory).GetProperty(columnName);

                                    if (property == null) continue;

                                    var existingValue = property.GetValue(existingRow)?.ToString() ?? "";

                                    if (existingValue != newValue) // Update if values differ
                                    {
                                        hasChanges = true;
                                        property.SetValue(existingRow, string.IsNullOrWhiteSpace(newValue) ? null : newValue);
                                        Console.WriteLine($"Row with Asset '{uniqueKey}', Column '{columnName}' updated: '{existingValue}' -> '{newValue}'");
                                    }
                                }

                                if (hasChanges)
                                {
                                    updates.Add(existingRow); // Add to update batch
                                }
                            }
                            else // Add new rows
                            {
                                var newRow = new MasterInventory();
                                for (int colIndex = 0; colIndex < headers.Length; colIndex++)
                                {
                                    var columnName = headers[colIndex].Trim();
                                    var value = values[colIndex].Trim();
                                    var property = typeof(MasterInventory).GetProperty(columnName);

                                    if (property != null)
                                    {
                                        property.SetValue(newRow, string.IsNullOrWhiteSpace(value) ? null : value);
                                    }
                                }

                                // Check if the new row is truly unique
                                if (!dbRowKeys.Contains(uniqueKey))
                                {
                                    newRows.Add(newRow);
                                    dbRowKeys.Add(uniqueKey);
                                    Console.WriteLine($"New row added: Asset '{uniqueKey}'");
                                }
                            }
                        }

                        // Bulk update modified rows
                        if (updates.Count > 0)
                        {
                            context.BulkUpdate(updates);
                            Console.WriteLine($"Database updated with {updates.Count} modified rows.");
                        }

                        // Bulk insert new rows
                        if (newRows.Count > 0)
                        {
                            context.BulkInsert(newRows);
                            Console.WriteLine($"Database added with {newRows.Count} new rows.");
                        }
                        else
                        {
                            Console.WriteLine("No new rows to add.");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file {filePath}: {ex.Message}");
            }
        }

        private static string GetValue(string[] headers, string[] values, string headerName)
        {
            var index = Array.IndexOf(headers, headerName);
            return index >= 0 && index < values.Length ? values[index].Trim() : string.Empty;
        }
    }
}
