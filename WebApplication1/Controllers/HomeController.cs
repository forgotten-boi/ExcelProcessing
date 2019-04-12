using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace WebApplication1.Controllers
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
                string path = string.Format("{0}/{1}", Server.MapPath("~/ExcelData/Uploads"), Request.Files["FileUpload1"].FileName);
                if (!Directory.Exists(path)) // if upload folder path does not exist, create one.
                {
                    Directory.CreateDirectory(Server.MapPath("~/ExcelData/Uploads"));
                }

                string[] validFileTypes = { ".xls", ".xlsx", ".csv" };
                if (validFileTypes.Contains(extension))
                {
                    if (System.IO.File.Exists(path))
                    {
                        System.IO.File.Delete(path); // if file exist previously, delete previous one
                    }
                    httpFileCollection.SaveAs(path);
                }


            }
            TempData["Success"] = "Success";
            return new ContentResult();
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