using Microsoft.EntityFrameworkCore;
using System.ComponentModel.DataAnnotations;  // ✅ For [Key] attribute
using System.ComponentModel.DataAnnotations.Schema;  // ✅ For [Column] attribute


public class MasterDbContext : DbContext
{
    public DbSet<MasterInventory> MasterInventory { get; set; }
    public DbSet<SapEam> SapEam { get; set; }
    public DbSet<SapEamTest> SapEamTest { get; set; }
    public DbSet<FaSummary> FaSummary { get; set; }
    public DbSet<CycleCount> CycleCount { get; set; } 
    public DbSet<ChangeLogs> ChangeLogs { get; set; } 
    public DbSet<WcMapping> WcMapping { get; set; } 



    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
    {
        // Replace with your actual SQL Server connection string
        // optionsBuilder.UseSqlServer(@"Server=mypenm0iesvr01;Database=Testing;Trusted_Connection=True;TrustServerCertificate=True;");
        optionsBuilder.UseSqlServer(@"Server=mypenm0iesvr01;Database=Testing;Trusted_Connection=True;TrustServerCertificate=True;");

    }


}

public class MasterInventory
{
    public int Id { get; set; }
    public string? Asset { get; set; }
    public string Description { get; set; } = string.Empty;
    public string? TorqueSerialNumber { get; set; }
    public string? TorqueModel { get; set; }
    public string? TorqueRange { get; set; }
    public string? ControllerSerialNumber { get; set; }
    public string? Brand { get; set; }
    public string? Manufacturer { get; set; }
    public string? Supplier { get; set; }
    public string? EquipmentID { get; set; }
    public string? RegisterSAPEAM { get; set; }
    
    public string? FunctionalGroup { get; set; }
    public string? ProcessGroup { get; set; }
    public string? Sector { get; set; }
    public string? Plant { get; set; }
    public string? Workcell { get; set; }
    public string? CostCenter { get; set; }
    public string? NetBookValue { get; set; }
    public string? UserStatus { get; set; }
    public string? KeyInDate { get; set; }
    public string? Owner { get; set; }
    public string? Comment { get; set; }
}

public class SapEamInventory
{
    public int Id { get; set; }
    public string Equipment { get; set; } = string.Empty;
    public string? Asset { get; set; }
    public string? ManufSerialNo { get; set; }
    public string? SortField { get; set; }
    public string? WorkCenter { get; set; }
    public string? ModelNumber { get; set; }
    public string? CostCenter { get; set; }
    public string? MainWorkCtr { get; set; } // Main WorkCtr
    public string? ManufCountry { get; set; } // ManufCountry
    public string? PlannerGroup { get; set; }
    public string? SerialNumber { get; set; } // Serial Number
    public string? CatalogProfile { get; set; } // Catalog Profile
    public string? UserStatus { get; set; }
    public string? SystemStatus { get; set; }
    public DateTime CreatedAt { get; set; } = DateTime.Now;
    public DateTime UpdatedAt { get; set; } = DateTime.Now;
}



public class FaSummary
{
    public int Id { get; set; }
    public string Asset { get; set; }
    public string AssetDescription { get; set; }
    public string SerialNumber { get; set; }
    public string CostCtr { get; set; }
    public string CostCtrDescription { get; set; }
    public string Workcell { get; set; }
    public string CapDate { get; set; }
    public string AcquVal { get; set; }
    public string NetBookVal { get; set; }
    public string TypeName { get; set; }
    public string Dept { get; set; }
    public string Plant { get; set; }
    public string Supplier { get; set; }
}

public class CycleCount
{
    public int Id { get; set; } // Primary Key

    public string? ManufSerialNo { get; set; }
    public string? MainWorkCtr { get; set; }
    public string? January2024 { get; set; }
    public string? February2024 { get; set; }
    public string? March2024 { get; set; }
    public string? April2024 { get; set; }
    public string? May2024 { get; set; }
    public string? June2024 { get; set; }
    public string? July2024 { get; set; }
    public string? August2024 { get; set; }
    public string? September2024 { get; set; }
    public string? October2024 { get; set; }
    public string? November2024 { get; set; }
    public string? December2024 { get; set; }
    public string? Sum { get; set; }
}



public class SapEam
{
    public int? Id { get; set; } // Primary key, if needed
    public string? Equipment { get; set; }
    public string? Asset { get; set; }
    public string? ManufSerialNo { get; set; }
    public string? SortField { get; set; }
    public string? Description { get; set; }
    public string? Location { get; set; }
    public string? Room { get; set; }
    public string? WorkCenter { get; set; }
    public string? Position { get; set; }
    public string? ObjectType { get; set; }
    public string? ModelNumber { get; set; }
    public string? CostCenter { get; set; }
    public string? CompanyCode { get; set; }
    public string? PlanningPlant { get; set; }
    public string? MaintPlant { get; set; }
    public string? PlantSection { get; set; }
    public string? Manufacturer { get; set; }
    public string? ManufPartNo { get; set; }
    public string? FunctionalLoc { get; set; }
    public string? ABCIndic { get; set; }
    public string? SubNumber { get; set; }
    public string? ConstructYear { get; set; }
    public string? ConstructMth { get; set; }
    public string? GrossWeight { get; set; }
    public string? AcquisitionDate { get; set; }
    public string? AcquisitionValue { get; set; }
    public string? WeightUnit { get; set; }
    public string? SizeDimensions { get; set; }
    public string? MainWorkCtr { get; set; }
    public string? ManufCountry { get; set; }
    public string? PlannerGroup { get; set; }
    public string? SerialNumber { get; set; }
    public string? CatalogProfile { get; set; }
    public string? UserStatus { get; set; }
    public string? SystemStatus { get; set; }
    public string? Workcell { get; set; }
}
public class SapEamTest
{
    public int? Id { get; set; } // Primary key, if needed
    public string? Equipment { get; set; }
    public string? Asset { get; set; }
    public string? ManufSerialNo { get; set; }
    public string? SortField { get; set; }
    public string? Description { get; set; }
    public string? Location { get; set; }
    public string? Room { get; set; }
    public string? WorkCenter { get; set; }
    public string? Position { get; set; }
    public string? ObjectType { get; set; }
    public string? ModelNumber { get; set; }
    public string? CostCenter { get; set; }
    public string? CompanyCode { get; set; }
    public string? PlanningPlant { get; set; }
    public string? MaintPlant { get; set; }
    public string? PlantSection { get; set; }
    public string? Manufacturer { get; set; }
    public string? ManufPartNo { get; set; }
    public string? FunctionalLoc { get; set; }
    public string? ABCIndic { get; set; }
    public string? SubNumber { get; set; }
    public string? ConstructYear { get; set; }
    public string? ConstructMth { get; set; }
    public string? GrossWeight { get; set; }
    public string? AcquisitionDate { get; set; }
    public string? AcquisitionValue { get; set; }
    public string? WeightUnit { get; set; }
    public string? SizeDimensions { get; set; }
    public string? MainWorkCtr { get; set; }
    public string? ManufCountry { get; set; }
    public string? PlannerGroup { get; set; }
    public string? SerialNumber { get; set; }
    public string? CatalogProfile { get; set; }
    public string? UserStatus { get; set; }
    public string? SystemStatus { get; set; }
    public string? Workcell { get; set; }
}


public class ChangeLogs
{
    public int Id { get; set; }  // Auto-incremented primary key
    public int GroupId { get; set; }  // Used for grouping
    public string ManufSerialNo { get; set; }
    public string ColumnName { get; set; }
    public string OldValue { get; set; }
    public string NewValue { get; set; }
    public string ChangeType { get; set; }  // 'Create', 'Update', 'Delete'
    public DateTime Timestamp { get; set; } = DateTime.UtcNow;
}



public class WcMapping
{
    [Key]  // ✅ Mark as Primary Key
    [Column(TypeName = "nvarchar(50)")] // Adjust size accordingly
    public string MainWorkCtr { get; set; } = string.Empty;  

    [Column(TypeName = "nvarchar(100)")]  // Adjust size accordingly
    public string Workcell { get; set; } = string.Empty;  
}