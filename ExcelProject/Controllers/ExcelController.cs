using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Model;

namespace ExcelProject.Controllers
{
    public class ExcelController : Controller
    {
        // GET: Excel
        public ActionResult Index()
        {
            


            return View();
        }

        public ActionResult UploadExcelFile()
        {
            //Delete previous data in database or use unique key
            //Save excel file from form to an object
            HttpPostedFileBase parametersTemplate = Request.Files["muhSpecialFileName"];
            //Save excel file to server
            string path = Server.MapPath("~/Content/Files/Upload" + parametersTemplate.FileName);
            if (System.IO.File.Exists(path))
            {
                System.IO.File.Delete(path);
                parametersTemplate.SaveAs(path);
            }
            parametersTemplate.SaveAs(path);
            
            //Save new parameters to database
            var excelData = new ExcelData(path); // link to other project
            var fileInput = excelData.getData("MuhSheetOne");
            var calculationList = new List<CalculationModel>();

            foreach (var row in fileInput)
            {
                var parameterDataToAdd = new CalculationModel()
                {
                    Parameter1 = float.Parse(row["Heading 1"].ToString()),
                    Parameter2 = float.Parse(row["Heading 2"].ToString()),
                    
                    Result = ((float.Parse(row["Heading 1"].ToString())) * float.Parse(row["Heading 2"].ToString()))
                };
                calculationList.Add(parameterDataToAdd);
            }

            var testFlag = calculationList;



            return RedirectToAction("Index");
        }
    }
}