using ExcelImport.Entities;
using ExcelImport.ViewModel;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Web;

namespace Excel_Import.Models
{

    public class BL_ManageExcelData
    {
        #region Properties
        public List<CompnyDetailsViewModel> CompRecords { get; set; }

        public HttpPostedFileBase uploadedExcel { get; set; }

        public bool Result { get; set; }

        public string Message { get; set; }
        #endregion

        #region Constructor
        public BL_ManageExcelData()
        {
          
        }
        public BL_ManageExcelData(HttpPostedFileBase Doc)
        {
            uploadedExcel = Doc;
        }
        #endregion

        #region ReadExcel
        /// <summary>
        /// Reads All data from uploade Excel
        /// </summary>
        /// <returns></returns>
        public void ReadExcel(string pathToExcelFile)
        {
            
            ExcelManager excelManager = null;
            try
            {
                
                CompRecords = new List<CompnyDetailsViewModel>();
                excelManager = new ExcelManager(true);
                excelManager.xl_workbook = excelManager.xl_app.Workbooks.Open(pathToExcelFile, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                excelManager.xl_worksheet = (Worksheet)excelManager.xl_workbook.Sheets.Cast<Worksheet>().FirstOrDefault();
                if (excelManager.xl_worksheet != null)
                {
                    excelManager.xl_worksheet.Activate();
                    excelManager.xl_range = excelManager.xl_worksheet.UsedRange;
                    int rowCount = excelManager.xl_range.Rows.Count;
                    int colCount = excelManager.xl_range.Columns.Count;
                    #region Read Excel
                    for (int i = 2; i <= rowCount; i++)
                    {
                        if (string.IsNullOrEmpty((excelManager.xl_worksheet.Cells[i, 2] as Range).Value) && string.IsNullOrEmpty((excelManager.xl_worksheet.Cells[i, 3] as Range).Value) && string.IsNullOrEmpty((excelManager.xl_worksheet.Cells[i + 1, 2] as Range).Value) &&
                            string.IsNullOrEmpty((excelManager.xl_worksheet.Cells[i + 1, 3] as Range).Value))
                            break;

                        if (!string.IsNullOrEmpty((excelManager.xl_worksheet.Cells[i, 2] as Range).Value) && !string.IsNullOrEmpty((excelManager.xl_worksheet.Cells[i, 3] as Range).Value))
                        {
                            CompnyDetailsViewModel CompDetails = new CompnyDetailsViewModel();
                            //Company Name
                            CompDetails.CompName = string.IsNullOrEmpty((excelManager.xl_worksheet.Cells[i, 2] as Range).Value) ? string.Empty : (excelManager.xl_worksheet.Cells[i, 2] as Range).Value.ToString();

                            //GSTIN
                            CompDetails.GSTIN = string.IsNullOrEmpty((excelManager.xl_worksheet.Cells[i, 3] as Range).Value) ? string.Empty : (excelManager.xl_worksheet.Cells[i, 3] as Range).Value.ToString();

                            //Start Date
                            CompDetails.StartDate =  (excelManager.xl_worksheet.Cells[i, 4] as Range)?.Value ?? null;

                            //End Date
                            CompDetails.EndDate = (excelManager.xl_worksheet.Cells[i, 5] as Range)?.Value ?? null;

                            //TurnOver Date
                            CompDetails.TurnOverAmount =  (excelManager.xl_worksheet.Cells[i,6] as Range)?.Value?.ToString();

                            //Contact Email
                            CompDetails.EmailId =  (excelManager.xl_worksheet.Cells[i, 7] as Range)?.Value?.ToString();

                            //Contact Contact No
                            CompDetails.ContactNo = (excelManager.xl_worksheet.Cells[i,8] as Range)?.Value?.ToString();

                            CompRecords.Add(CompDetails);
                        }
                    }
                    #endregion
                }
                excelManager.xl_range = null;
                excelManager.xl_workbook.Save();
                excelManager.xl_workbook.Close(true);

                excelManager.ReleaseExcelObjects();
                excelManager = null;
                GC.Collect();
                Result = true;
                Message = "Data Saved Successfully";
            }
            catch (Exception ex)
            {
                excelManager.ReleaseExcelObjects();
                excelManager = null;
                GC.Collect();
                Result = false;
                Message = "Failed to Read Excel";
            }
        }
        #endregion

        #region Save Data in Database
        public void SaveCompanyDetails()
        {
            try
            {
                ExcelImportEntities dbContext = new ExcelImportEntities();
                List<string> ExistingGSTN = dbContext.CompanyDetails?.ToList()?.Select(c =>c.GSTIN).ToList();
                List<string> ExistingEmailIDs = dbContext.CompanyDetails?.ToList()?.Select(c => c.EmailId).ToList();
                List<string> ExistingContactNo = dbContext.CompanyDetails?.ToList()?.Select(c => c.ContactNo).ToList();
                CompRecords.All(c =>
                {
                    ValidationContext vc = new ValidationContext(c);
                    ICollection<ValidationResult> results = new List<ValidationResult>();
                    bool isValid = Validator.TryValidateObject(c, vc, results, true);
                    CompanyDetail comp = new CompanyDetail();//I have created another view model for giving validation attribute because we have to also save data with error
                    if (!isValid)
                        comp.IsValidData = (int)Common.Constants.IsValidData.Error;
                    else if((ExistingGSTN?.Contains(c.GSTIN) ?? false )|| (ExistingEmailIDs?.Contains(c.EmailId) ?? false )||( ExistingContactNo?.Contains(c.ContactNo) ?? false))
                        comp.IsValidData = (int)Common.Constants.IsValidData.Duplicate;
                    else
                        comp.IsValidData = (int)Common.Constants.IsValidData.Valid;
                    
                    comp.CompName = c.CompName;
                    comp.GSTIN = c.GSTIN;
                    comp.StartDate = c.StartDate;
                    comp.EndDate = c.EndDate;
                    comp.TurnOverAmount = c.TurnOverAmount;
                    comp.EmailId = c.EmailId;
                    comp.ContactNo = c.ContactNo;
                    dbContext.CompanyDetails.Add(comp);
                    dbContext.SaveChanges();
                    return true;
                });
                Result = true;
                Message = "Data Saved Successfully";
            }
            catch (Exception ex)
            {
                Result = false;
                Message = "Failed to save data";
            }
        }
        #endregion

        public void EditCompDetails(CompnyDetailsViewModel model)
        {
            try
            {
                ExcelImportEntities dbContext = new ExcelImportEntities();
                List<string> ExistingGSTN = dbContext.CompanyDetails?.ToList()?.Where(c =>c.CompId != model.CompId)?.Select(c => c.GSTIN).ToList();
                List<string> ExistingEmailIDs = dbContext.CompanyDetails?.ToList()?.Where(c => c.CompId != model.CompId)?.Select(c => c.EmailId).ToList();
                List<string> ExistingContactNo = dbContext.CompanyDetails?.ToList()?.Where(c => c.CompId != model.CompId)?.Select(c => c.ContactNo).ToList();
                CompanyDetail comp = dbContext.CompanyDetails.Find(model.CompId);
                if(comp != null)
                {
                    if ((ExistingGSTN?.Contains(model.GSTIN) ?? false) || (ExistingEmailIDs?.Contains(model.EmailId) ?? false) || (ExistingContactNo?.Contains(model.ContactNo) ?? false))
                        comp.IsValidData = (int)Common.Constants.IsValidData.Duplicate;
                    else
                        comp.IsValidData = (int)Common.Constants.IsValidData.Valid;

                    comp.CompName = model.CompName;
                    comp.StartDate = model.StartDate;
                    comp.EndDate = model.EndDate;
                    comp.GSTIN = model.GSTIN;
                    comp.TurnOverAmount = model.TurnOverAmount;
                    comp.EmailId = model.EmailId;
                    comp.ContactNo = model.ContactNo;
                    dbContext.CompanyDetails.Attach(comp);
                    dbContext.Entry(comp).State = EntityState.Modified;
                    dbContext.SaveChanges();
                    Result = true;
                    Message = "Company Details Edited successfully !";
                }
                Result = false;
                Message = "Problem Occured While Editing Compant Details!";

            }
            catch (Exception ex)
            {
                Result = false;
                Message = "Problem Occured While Editing Compant Details!";
            }

        }


        #region ExcelManager Class
        public class ExcelManager
        {
            public Type ty { get; set; } = Type.GetTypeFromProgID("Excel.Application");
            public Application xl_app { get; set; }
            public Workbook xl_workbook { get; set; }
            public Worksheet xl_worksheet { get; set; }
            public Range xl_range { get; set; }
            public object misValue { get; set; } = System.Reflection.Missing.Value;

            public ExcelManager()
            {

            }
            public ExcelManager(bool IsNew)
            {
                if (IsNew)
                {

                    xl_app = (Application)Activator.CreateInstance(ty);
                    xl_app.Visible = true;
                    xl_app.DisplayAlerts = false;
                }
            }



            public void ReleaseExcelObjects()
            {
                releaseObject(xl_range);
                releaseObject(xl_worksheet);
                releaseObject(xl_workbook);
                if (xl_app != null)
                {
                    try
                    {
                        xl_app.Quit();
                    }
                    catch (Exception ex)
                    {
                        // Machine_Global.Global.WriteLog("Error While Quit() Xl_App", ex);
                    }
                }
                releaseObject(xl_app);
            }

            public void releaseObject(object obj)
            {
                try
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                    obj = null;
                }
                catch (Exception ex)
                {
                    obj = null;
                    // MessageBox.Show("Unable to release the Object " + ex.ToString());
                }
                finally
                {
                    GC.Collect();
                }
            }
        }
        #endregion
    }
}