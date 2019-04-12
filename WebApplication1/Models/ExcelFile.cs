using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace WebApplication1.Models
{
    public class ExcelFile
    {
        string path = "";
        _Application excel = new _Excel.Application();
        Workbook workbook;
        Worksheet workSheet;
        public ExcelFile(string path, int sheet)
        {
            this.path = path;
            workbook = excel.Workbooks.Open(path);
            workSheet = workbook.Worksheets[sheet];
        }

        public string ReadCell(int i, int j)
        {
            i++;
            j++;
            if (workSheet.Cells[i, j].Value2 != null)
            {
                return workSheet.Cells[i, j].Value2;
            }
            else
                return "";
        }
         public string[,] ReadMultiCell(int starti, int endi, int startj, int endj)
        {
            Range range = (Range)workSheet.Range[workSheet.Cells[starti, startj], workSheet.Cells[endi,endj]]
            object[,] holder = range.Value2;
            string[,] returnString = new string[endi - starti, endj = startj];
            for(int x=1; x<=endi-starti; x++)
            {
                for (int y = 1; y <= endj - startj; y++)
                {
                    returnString[x - 1, y - 1] = holder[x, y].ToString();
                }
            }
            return returnString;
        }

        public void WriteToCell(int i, int j, string s)
        {
            i++;
            j++;
            workSheet.Cells[i, j].Value2 = s;
        }

        public void CreateNewFile()
        {
            this.workbook = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            this.workSheet = workbook.Worksheets[1];
        }

        public void CreateNewSheet()
        {
            Worksheet tempSheet = workbook.Worksheets.Add(After: workSheet);
            
        }

        public void SelectWorksheet(int sheetNumber)
        {
            this.workSheet = workbook.Worksheets[sheetNumber];
        }
        public void DeleteWorksheet(int sheetNumber)
        {
            workbook.Worksheets[sheetNumber].Delete();
        }
        public void Save()
        {
            workbook.Save();
        }
        public void SaveAs(string s)
        {
            workbook.SaveAs(s);
        }
        //public string ExportExcel()
        //{
        //    Workbook wb = new Workbook();
        //    Worksheet ws = new Worksheet();
        //    ws = wb.Worksheets[1];

        //    for (int i=1; i< workSheet.Rows.Count; i++)
        //        for(int j=1; j<workSheet.Columns.Count; j++)
        //        {
        //            if (workSheet.Cells[i,j] !=null)
        //               ws = workSheet.w


        //        }

        //}
     }
}