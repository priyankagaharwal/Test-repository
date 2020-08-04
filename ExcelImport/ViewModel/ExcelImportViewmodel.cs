using ExcelImport.Entities;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace ExcelImport.ViewModel
{
    public class ExcelImportViewmodel
    {
        [Required(ErrorMessage = "Please upload an Excel Document")]
        [RegularExpression(@"^.*\.(xlsx|xls)$", ErrorMessage = "Please Upload Excel File only")]
        public HttpPostedFileBase Document { get; set; }
    }
    public class CompanyListViewModel
    {
       public List<CompanyDetail> ValidDataList { get; set; }

       public List<CompanyDetail> DuplicateDataList { get; set; }

        public List<CompanyDetail> ErrorDataList { get; set; }
    }
    public class CompnyDetailsViewModel 
    {
        public int CompId { get; set; }

        [Required(ErrorMessage = "Company Name Is Required")]
        [Display(Name = "Comp Name")]
        public string CompName { get; set; }

        [Required(ErrorMessage = "GSTIN Is Required")]
        [Display(Name = "GSTIN")]
        [RegularExpression(@"\d{2}[A-Z]{5}\d{4}[A-Z]{1}[A-Z\d]{1}[Z]{1}[A-Z\d]{1}", ErrorMessage = "Please Enter A Valid GSTIN")]
        public string GSTIN { get; set; }

        [Required(ErrorMessage = "Start Date Is Required")]
        [Display(Name = "Start Date")]
        [DisplayFormat(ApplyFormatInEditMode = true, DataFormatString = "{0:dd/MM/yyyy}")]
        public DateTime? StartDate { get; set; }

        [Required(ErrorMessage = "End Date Is Required")]
        [Display(Name = "End Date")]
        [DisplayFormat(ApplyFormatInEditMode = true, DataFormatString = "{0:dd/MM/yyyy}")]
        public DateTime? EndDate { get; set; }

        [Required(ErrorMessage = "Turn Over Amount Is Required")]
        [Display(Name = "Turn Over")]
        [Range(0, Double.PositiveInfinity, ErrorMessage = "TurnOver Amount Must be positive ")]
        public string TurnOverAmount { get; set; }

        [Required(ErrorMessage = "Email Id Is Required")]
        [RegularExpression(@"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", ErrorMessage = "Please Enter A Valid Email ID")]
        public string EmailId { get; set; }

        [Required(ErrorMessage = "Contact No Is Required")]
        [Display(Name = "Contact No")]
        [RegularExpression(@"^(?:(?:\+|0{0,2})91(\s*[\ -]\s*)?|[0]?)?[789]\d{9}|(\d[ -]?){10}\d$", ErrorMessage = "Please Enter A Valid Contact No")]
        public string ContactNo { get; set; }

        public int? IsValidData { get; set; }

       
    }
    
}