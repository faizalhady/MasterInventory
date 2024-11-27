using EFCore.BulkExtensions;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace Master.Services
{
    public static class CsvFileProcessor
    {
        private static int rowCount = 0;
        private const int batchSize = 1000;
        private static List<MasterInventory> dataBatch = new List<MasterInventory>();

        public static void ProcessCsvFile()
        {
            string? filePath = @"\\mypenm0opsapp01\SmartTorque$\eST1C\Inventory Database\Inventory_Database_11222024.csv";
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                Console.WriteLine($"File does not exist or invalid: {filePath}");
                return;
            }

            Console.WriteLine($"Processing file: {filePath}");

            try
            {
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

                    // Read data rows
                    while ((line = reader.ReadLine()) != null)
                    {
                        try
                        {
                            var values = line.Split(',');

                            var masterInventory = new MasterInventory
                            {
                                Asset = GetValue(headers, values, "Asset"),
                                Description = GetValue(headers, values, "Description"),
                                TorqueSerialNumber = GetValue(headers, values, "Torque Serial Number"),
                                TorqueModel = GetValue(headers, values, "Torque Model"),
                                TorqueRange = GetValue(headers, values, "Torque Range"),
                                ControllerSerialNumber = GetValue(headers, values, "Controller Serial Number"),
                                Brand = GetValue(headers, values, "Brand"),
                                Manufacturer = GetValue(headers, values, "Manufacturer"),
                                Supplier = GetValue(headers, values, "Supplier"),
                                EquipmentID = GetValue(headers, values, "Equipment ID"),
                                RegisterSAPEAM = GetValue(headers, values, "Register SAP EAM"),
                                FunctionalGroup = GetValue(headers, values, "Functional Group"),
                                AssemblyGroup = GetValue(headers, values, "Assembly Group"),
                                Sector = GetValue(headers, values, "Sector"),
                                Plant = GetValue(headers, values, "Plant"),
                                Workcell = GetValue(headers, values, "Workcell"),
                                CostCenter = GetValue(headers, values, "Cost Center"),
                                NetBookValue = GetValue(headers, values, "Net Book Value"),
                                UserStatus = GetValue(headers, values, "User Status"),
                                KeyInDate = DateTime.TryParse(GetValue(headers, values, "Key In Date"), out var keyInDate) ? keyInDate : (DateTime?)null,
                                Owner = GetValue(headers, values, "Owner"),
                                Comment = GetValue(headers, values, "Comment")
                            };

                            AddToBatch(masterInventory);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error processing row: {ex.Message}");
                        }
                    }
                }

                FlushBatch();
                Console.WriteLine($"Processing complete. Total rows processed: {rowCount}");
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

        private static void AddToBatch(MasterInventory masterInventory)
        {
            dataBatch.Add(masterInventory);

            if (dataBatch.Count >= batchSize)
            {
                FlushBatch();
            }
        }

        private static void FlushBatch()
        {
            if (dataBatch.Count > 0)
            {
                using (var context = new MasterDbContext())
                {
                    context.BulkInsertOrUpdate(dataBatch); // Bulk insert or update into MasterInventory table
                }

                rowCount += dataBatch.Count;
                dataBatch.Clear();
            }
        }
    }
}
