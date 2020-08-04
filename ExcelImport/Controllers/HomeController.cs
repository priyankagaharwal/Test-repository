using Excel_Import.Models;
using ExcelImport.Entities;
using ExcelImport.ViewModel;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ExcelImport.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            ExcelImportEntities dbContext = new ExcelImportEntities();
            CompanyListViewModel comp = new CompanyListViewModel();
            comp.ValidDataList = dbContext.CompanyDetails.Where(c => c.IsValidData == (int)Excel_Import.Common.Constants.IsValidData.Valid).ToList();
            comp.DuplicateDataList = dbContext.CompanyDetails.Where(c => c.IsValidData == (int)Excel_Import.Common.Constants.IsValidData.Duplicate).ToList();
            comp.ErrorDataList = dbContext.CompanyDetails.Where(c => c.IsValidData == (int)Excel_Import.Common.Constants.IsValidData.Error).ToList();
            return View(comp);
        }
        public ActionResult UploadExcel()
        {
            return PartialView();
        }

        [HttpPost]
        public ActionResult UploadExcel(ExcelImportViewmodel model)
        {
            string targetpath = Server.MapPath("~/App_Data/");
            model.Document.SaveAs(targetpath + Path.GetFileName(model.Document.FileName));
            string pathToExcelFile = targetpath + Path.GetFileName(model.Document.FileName);
            try
            {
                BL_ManageExcelData bl = new BL_ManageExcelData(model.Document);
                bl.ReadExcel(pathToExcelFile);
                if (bl.Result)
                {
                    bl.SaveCompanyDetails();
                }
                System.IO.File.Delete(pathToExcelFile);
                return Json(new { Result = bl.Result, Message = bl.Message }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                System.IO.File.Delete(pathToExcelFile);
                return Json(new { Result = false, Message = "Failed To upload Excel !" }, JsonRequestBehavior.AllowGet);
            }
           
        }

        public ActionResult EditCompDetails(int compId)
        {
            ExcelImportEntities dbContext = new ExcelImportEntities();
            CompnyDetailsViewModel comp = new CompnyDetailsViewModel();
            CompanyDetail compd = dbContext.CompanyDetails.Find(compId);
            if(compd != null)
            {
                comp.CompName = compd.CompName;
                comp.GSTIN = compd.GSTIN;
                comp.StartDate = compd.StartDate;
                comp.EndDate = compd.EndDate;
                comp.TurnOverAmount = compd.TurnOverAmount;
                comp.EmailId = compd.EmailId;
                comp.ContactNo = compd.ContactNo;
                comp.CompId = compd.CompId;
            }
            
            return PartialView(comp);
        }

        [HttpPost]
        public ActionResult EditCompDetails(CompnyDetailsViewModel model)
        {

            if (ModelState.IsValid)
            {
                BL_ManageExcelData bl = new BL_ManageExcelData();
                bl.EditCompDetails(model);
                return Json(new { Result = bl.Result, Message = bl.Message }, JsonRequestBehavior.AllowGet);
            }
            else
            {
                return Json(new { Result = false, Message = "Problem Occured While Editing Company Details!" }, JsonRequestBehavior.AllowGet);
            }
        }

      
    }
}