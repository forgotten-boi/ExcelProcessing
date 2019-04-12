using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;

namespace WebApplication1.Models
{
    public class XslLibrary
    {
        internal DataTable ReadDataTable(string path1, string extension)
        {
            var dt = new DataTable();
            string connString = "";
            //add different connection string for different types of excel
            if (extension == ".csv")
            {
                dt = XslUtil.ConvertCSVtoDataTable(path1);

            }
            //Connection String to Excel Workbook  
            else if (extension.Trim() == ".xls")
            {
                connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path1 + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
                dt = XslUtil.ConvertXSLXtoDataTable(path1, connString);

            }
            else if (extension.Trim() == ".xlsx")
            {
                connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path1 + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
                dt = XslUtil.ConvertXSLXtoDataTable(path1, connString);

            }
            return dt;
        }

        internal string ImportFromExcelNpoi(string filePath)
        {
            ISheet sheet; //Create the ISheet object to read the sheet cell values  
            string filename = Path.GetFileName(filePath); //get the uploaded file name  

            var fileExt = Path.GetExtension(filename); //get the extension of uploaded excel file  
            using (StreamReader sr = new StreamReader(filePath))
            {
                if (fileExt == ".xls")
                {

                    HSSFWorkbook hssfwb = new HSSFWorkbook(sr.BaseStream); //HSSWorkBook object will read the Excel 97-2000 formats  
                    sheet = hssfwb.GetSheetAt(0); //get first Excel sheet from workbook  


                }
                else
                {
                    XSSFWorkbook hssfwb = new XSSFWorkbook(sr.BaseStream); //XSSFWorkBook will read 2007 Excel format  
                    sheet = hssfwb.GetSheetAt(0); //get first Excel sheet from workbook   
                }
            }
            string filteredFile = "";

            for (int row = 0; row <= sheet.LastRowNum; row++) //Loop the records upto filled row  
            {
                if (sheet.GetRow(row) != null) //null is when the row only contains empty cells   
                {
                    string value = sheet.GetRow(row).GetCell(0).StringCellValue; //Here for sample , I just save the value in "value" field, Here you can write your custom logics...  
                }
            }

            return filteredFile;
            //return Json(true, JsonRequestBehavior.AllowGet); //return true to display the success message  
        }



        internal void XslReader(string filePath)
        {
            HSSFWorkbook hssfwb;
            using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new HSSFWorkbook(file);
            }

            ISheet sheet = hssfwb.GetSheet("Arkusz1");
            for (int row = 0; row <= sheet.LastRowNum; row++)
            {
                if (sheet.GetRow(row) != null) //null is when the row only contains empty cells 
                {

                    //MessageBox.Show(string.Format("Row {0} = {1}", row, sheet.GetRow(row).GetCell(0).StringCellValue));
                }
            }
        }


    }
}