using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using ExcelProcessor.Models;
using System.Diagnostics;

namespace ExcelProcessor.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }
        [HttpPost]
        public ActionResult ImportExcel()
        {
            var httpFileCollection = Request.Files["file"];
            if (httpFileCollection.ContentLength > 0)
            {
                string extension = System.IO.Path.GetExtension(httpFileCollection.FileName).ToLower();
                string path = string.Format("{0}/{1}", Server.MapPath("~/ExcelData/Uploads"), Request.Files["file"].FileName);
                if (!Directory.Exists(path)) // if upload folder path does not exist, create one.
                {
                    Directory.CreateDirectory(Server.MapPath("~/ExcelData/Uploads"));
                }

                var exportPath = path + "/Export";
                if (!Directory.Exists(exportPath))
                {
                    Directory.CreateDirectory(Server.MapPath("~/ExcelData/Uploads/Export"));
                }

                string[] validFileTypes = { ".xls", ".xlsx", ".csv" };
                if (validFileTypes.Contains(extension))
                {
                    if (System.IO.File.Exists(path))
                    {
                        System.IO.File.Delete(path); // if file exist previously, delete previous one
                    }
                    httpFileCollection.SaveAs(path);

                    var performance = new System.Diagnostics.PerformanceCounter("Memory", "Available MBytes");
                    Console.WriteLine(performance.NextValue());
                    var a = performance.NextValue();
                    XslLibrary.ReadDataTable(path);
                    Console.WriteLine(performance.NextValue());
                    // var b = performance.NextValue();
                    // XslLibrary.ImportFromExcelNpoi(path);
                    // Console.WriteLine(performance.NextValue());
                    // var c = performance.NextValue();
                    // TempData["Success"] = $"a: {a} B: {b} c: {c}";
                }


            }
           
            return View("Index");
        }
        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}