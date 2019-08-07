using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelProcessor.ExportToExcel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;

using Excel = Microsoft.Office.Interop.Excel;


namespace ExcelProcessor.Models
{
    public static class XslLibrary
    {
        public static DataTable ReadDataTable(string path1)
        {
            var extension = Path.GetExtension(path1);
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

        public static void WriteDatatableToExcel(DataTable table)
        {
            DataSet ds = new DataSet();
            ds.Tables.Add(table);
            
            CreateExcelFile.CreateExcelDocument(ds, "Sheet20.xlsx");
        }

        public static void  WriteDatatableToDataset(DataTable dt) 
        {
            //DataSet ds = new DataSet();
            //ds.Tables.Add(dt);
            string connString = @"data source=JV-PITAMBAR\SQLEXPRESS03;initial catalog=ExcelProcessor;persist security info=True; Integrated Security=SSPI;Connect Timeout=120";
           // string connString = "Data Source=.;uid=nish;pwd=nish;database=ExcelProcessor;Connect Timeout=120";
            var tableName = "tblExcelProcData";
            using (SqlConnection con = new SqlConnection(connString))
            {
                //int count = GetTableStructure(con, TableName).Count;
                var tablestructure = GetTableStructure(con, tableName);
                var count = tablestructure.Count;

                string[] columnNames = dt.Columns.Cast<DataColumn>()
                           .Select(x => x.ColumnName.Trim())
                           .ToArray();
                using (SqlTransaction transaction = con.BeginTransaction())
                {

                    SqlBulkCopy sbc = new SqlBulkCopy(con, SqlBulkCopyOptions.KeepIdentity, transaction);
                    sbc.DestinationTableName = tableName;

                   
             
                    foreach (DataColumn col in dt.Columns)
                    {
                        sbc.ColumnMappings.Add(col.ColumnName.ToString(), col.ColumnName.ToString());
                        //count++;
                    }

                    //sbc.BatchSize = 1000;
                    //sbc.NotifyAfter = 1000;
                    
                    sbc.WriteToServerAsync(dt);
                    transaction.Commit();

                }

                var datatable = new DataTable();
                var query = "select * from [tblExcelProcData]";
                SqlCommand cmd = new SqlCommand(query, con);
                using (var reader = cmd.ExecuteReader(CommandBehavior.SequentialAccess))
                {
                    datatable.Load(reader);
                };


            }

        }

      
        public static Dictionary<string, string> GetTableStructure(SqlConnection con, string tableName)
        {
            Dictionary<string, string> Param = new Dictionary<string, string>();

            string sqlCheckTable = "SELECT c.name as 'ColumnName', CONCAT(t.Name,'(',c.max_length,')') as 'DataType' FROM sys.columns c INNER JOIN sys.types t ON c.user_type_id = t.user_type_id WHERE c.object_id = OBJECT_ID('" + tableName + "')";

            try
            {

                if (con.State != ConnectionState.Open)
                con.Open();

                using (SqlCommand command = con.CreateCommand())
                {
                    command.CommandText = sqlCheckTable;

                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            Param.Add(reader["ColumnName"].ToString(), reader["DataType"].ToString());
                           
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                
                //log.Error(ex.Message);
            }
            return Param;
        }

        public static string ImportFromExcelNpoi(string filePath)
        {
            var wb = WorkbookFactory.Create(filePath);
            var wb2 = new XSSFWorkbook();
            var sheet2 = wb2.CreateSheet("Sheet1");

            string filteredFile = "";
            var sheet = wb.GetSheetAt(0);
            for (int row = 2; row <= sheet.LastRowNum; row++) //Loop the records upto filled row  
            {
                var dr = sheet.GetRow(row);
                var value = dr.GetCell(3)?.StringCellValue;
                if (value == "EA") //null is when the row only contains empty cells   
                {
                    var dr2 = sheet2.CreateRow(row - 2);
                    int count = 0;
                    foreach (var item in dr.Cells)
                    {
                        if(item.CellType != NPOI.SS.UserModel.CellType.Numeric)
                            dr2.CreateCell(count).SetCellValue(item.StringCellValue);
                        else
                            dr2.CreateCell(count).SetCellValue(item.NumericCellValue);
                        count++;
                    }
                    //dr2 = dr;
                }
            }
            //wb2.Write(File.OpenWrite("Sheet.xlsx"));

            FileStream xfile = new FileStream("Sheet10.xlsx", FileMode.Create, System.IO.FileAccess.Write);
            wb2.Write(xfile);

            xfile.Dispose();
            //filteredSheet.Workbook.Close();
            //var exportPath = filePath + "/Export";
            //if (!Directory.Exists(exportPath))
            //{
            //    Directory.CreateDirectory(exportPath);
            //}
            //FileStream xfile = new FileStream(Path.Combine(exportPath, filename + "_Exported"), FileMode.Create, System.IO.FileAccess.Write);
            //hSSFWorkbook.Write(xfile);
            //xfile.Close();



            return filteredFile;
            //return Json(true, JsonRequestBehavior.AllowGet); //return true to display the success message  
        }

        public static void ReadOpenExcel(string fileName)
        {
            using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fs, false))
                {
                    WorkbookPart workbookPart = doc.WorkbookPart;
                    SharedStringTablePart sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                    SharedStringTable sst = sstpart.SharedStringTable;

                    WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                    Worksheet sheet = worksheetPart.Worksheet;

                    var cells = sheet.Descendants<Cell>();
                    var rows = sheet.Descendants<Row>();

                    //Console.WriteLine("Row count = {0}", rows.LongCount());
                    //Console.WriteLine("Cell count = {0}", cells.LongCount());

                    //// One way: go through each cell in the sheet
                    foreach (Cell cell in cells)
                    {
                        if ((cell.DataType != null) && (cell.DataType == CellValues.SharedString))
                        {
                            int ssid = int.Parse(cell.CellValue.Text);
                            string str = sst.ChildElements[ssid].InnerText;
                            //Console.WriteLine("Shared string {0}: {1}", ssid, str);
                        }
                        else if (cell.CellValue != null)
                        {
                            //Console.WriteLine("Cell contents: {0}", cell.CellValue.Text);
                        }
                    }

                    // Or... via each row
                    //foreach (Row row in rows)
                    //{
                    //    foreach (Cell c in row.Elements<Cell>())
                    //    {
                    //        if ((c.DataType != null) && (c.DataType == CellValues.SharedString))
                    //        {
                    //            int ssid = int.Parse(c.CellValue.Text);
                    //            string str = sst.ChildElements[ssid].InnerText;
                    //            //Console.WriteLine("Shared string {0}: {1}", ssid, str);
                    //        }
                    //        else if (c.CellValue != null)
                    //        {
                    //            //Console.WriteLine("Cell contents: {0}", c.CellValue.Text);
                    //        }
                    //    }
                    //}
                }
            }
        }

        public static void XslReader(string filePath)
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
                    Console.WriteLine(sheet.GetRow(row).GetCell(0));
                    //MessageBox.Show(string.Format("Row {0} = {1}", row, sheet.GetRow(row).GetCell(0).StringCellValue));
                }
            }
        }


        public static void ExcelInterloop(string path)
        {
            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!


            

            for (int i = 1; i <= rowCount; i++)
            {
                var checkValue = xlRange.Cells[i, 3]?.Value2?.ToString();
                if (checkValue == "EA")
                {

                }
                //for (int j = 1; j <= colCount; j++)
                //{
                //    //new line
                //    if (j == 1)
                //        //Console.Write("\r\n");

                //        //write the value to the console
                //        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                //        {
                //            string CellValue = xlRange.Cells[i, j].Value2.ToString() + "\t";
                //            //Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");
                //        }
                //}
            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

        }




    }
}