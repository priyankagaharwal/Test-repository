namespace ExcelImport.Entities
{
    using System;
    using System.ComponentModel.DataAnnotations;
    using System.Data.Entity;

    public class ExcelImportEntities : DbContext
    {
       
        public ExcelImportEntities()
             : base("ExcelImport")
        {
            Database.SetInitializer(new CreateDatabaseIfNotExists<ExcelImportEntities>());
        }

        public DbSet<CompanyDetail> CompanyDetails { get; set; }
       
    }

    public partial class CompanyDetail
    {
        [Key]
        public int CompId { get; set; }

        public string CompName { get; set; }

        public string GSTIN { get; set; }

        public DateTime? StartDate { get; set; }

        public DateTime? EndDate { get; set; }

        public string TurnOverAmount { get; set; }

        public string EmailId { get; set; }

        public string ContactNo { get; set; }

        public int? IsValidData { get; set; }
    }
}