using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ExcelImport.Entities
{
    public class CompnyDetailsMetaData 
    {
       
        [Display(Name = "Comp Name")]
        public string CompName { get; set; }

        
        [Display(Name = "GSTIN")]
        public string GSTIN { get; set; }

        
        [Display(Name = "Start Date")]
        [DisplayFormat(ApplyFormatInEditMode = true, DataFormatString = "{0:dd/MM/yyyy}")]
        public DateTime? StartDate { get; set; }

        
        [Display(Name = "End Date")]
        [DisplayFormat(ApplyFormatInEditMode = true, DataFormatString = "{0:dd/MM/yyyy}")]
        public DateTime? EndDate { get; set; }

        
        [Display(Name = "Turn Over")]
        public string TurnOverAmount { get; set; }

        
        public string EmailId { get; set; }

        
        [Display(Name = "Contact No")]
        public string ContactNo { get; set; }

    }
    }
