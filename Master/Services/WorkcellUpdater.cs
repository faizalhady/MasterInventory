using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.EntityFrameworkCore;
using EFCore.BulkExtensions;

public static class WorkcellUpdater
{
    public static void UpdateWorkcellColumn()
    {
        Console.WriteLine("🔹 Starting Workcell Update Process...");

        using var dbContext = new MasterDbContext();

        try
        {
            // Test database connection
            dbContext.WcMapping.FirstOrDefault();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Database connection failed: {ex.Message}");
            return;
        }

        // Step 1: Fetch Workcell Mapping from the database
        var workcellMapping = dbContext.WcMapping.ToList();
        var workcellDictionary = workcellMapping.ToDictionary(w => w.MainWorkCtr, w => w.Workcell, StringComparer.OrdinalIgnoreCase);

        // Step 2: Fetch all records from SapEam to verify mappings
        var allRecords = dbContext.SapEam.ToList();
        var recordsToUpdate = new List<SapEam>();
        var newWorkcellEntries = new List<WcMapping>();

        foreach (var record in allRecords)
        {
            if (workcellDictionary.TryGetValue(record.MainWorkCtr, out var correctWorkcell))
            {
                // If Workcell does not match the mapping, update it
                if (record.Workcell != correctWorkcell)
                {
                    Console.WriteLine($"⚠️ Incorrect Workcell Found: {record.MainWorkCtr} -> {record.Workcell}, should be {correctWorkcell}");
                    record.Workcell = correctWorkcell;
                    recordsToUpdate.Add(record);
                }
            }
            else
            {
                // Handle missing mappings
                string newWorkcell = $"UNKNOWN_{record.MainWorkCtr}";
                Console.WriteLine($"🚨 Missing Mapping: Adding {record.MainWorkCtr} -> {newWorkcell}");

                var newMapping = new WcMapping { MainWorkCtr = record.MainWorkCtr, Workcell = newWorkcell };
                newWorkcellEntries.Add(newMapping);

                // Add to dictionary for future lookups
                workcellDictionary[record.MainWorkCtr] = newWorkcell;

                // Update SapEam with new Workcell mapping
                record.Workcell = newWorkcell;
                recordsToUpdate.Add(record);
            }
        }

        // Step 3: Bulk insert new workcell mappings
        if (newWorkcellEntries.Count > 0)
        {
            try
            {
                dbContext.BulkInsert(newWorkcellEntries);
                Console.WriteLine($"✅ Inserted {newWorkcellEntries.Count} new Workcell mappings.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Failed to insert new Workcell mappings: {ex.Message}");
            }
        }

        // Step 4: Bulk update incorrect Workcell records in SapEam
        if (recordsToUpdate.Count > 0)
        {
            try
            {
                dbContext.BulkUpdate(recordsToUpdate);
                Console.WriteLine($"✅ Updated {recordsToUpdate.Count} incorrect Workcell values.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Bulk Update Failed: {ex.Message}");
            }
        }

        dbContext.SaveChanges();

        // ✅ Summary Log
        Console.WriteLine($"✅ Process Completed: {newWorkcellEntries.Count} new mappings added, {recordsToUpdate.Count} Workcell values updated.");
    }
}
