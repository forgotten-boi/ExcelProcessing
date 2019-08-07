using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;


namespace ExcelProcessor.Models
{
    public class XslUtil
    {
        public static DataTable ConvertCSVtoDataTable(string strFilePath)
        {
            DataTable dt = new DataTable();
            using (StreamReader sr = new StreamReader(strFilePath))
            {
                string[] headers = sr.ReadLine().Split(',');
                foreach (string header in headers)
                {
                    dt.Columns.Add(header);
                }

                while (!sr.EndOfStream)
                {
                    string[] rows = sr.ReadLine().Split(',');
                    if (rows.Length > 1)
                    {
                        DataRow dr = dt.NewRow();
                        for (int i = 0; i < headers.Length; i++)
                        {
                            dr[i] = rows[i].Trim();
                        }
                        dt.Rows.Add(dr);
                    }
                }

            }


            return dt;
        }

        public static DataTable ConvertXSLXtoDataTable(string strFilePath, string connString)
        {
            OleDbConnection oledbConn = new OleDbConnection(connString);
            DataTable dt = new DataTable();
            DataSet ds = new DataSet()
            {
                EnforceConstraints = false
            };
            try
            {

                oledbConn.Open();
                int countData = 0;
                using (DataTable Sheets = oledbConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null))
                {

                    for (int i = 0; i < Sheets.Rows.Count; i++)
                    {
                        string worksheets = Sheets.Rows[i]["TABLE_NAME"].ToString();
                        OleDbCommand cmd = new OleDbCommand(String.Format("SELECT * FROM [{0}]", worksheets), oledbConn);
                        //OleDbCommand cmd = new OleDbCommand(String.Format("SELECT Count(*) FROM [{0}]", worksheets), oledbConn);
                        //OleDbCommand cmd = new OleDbCommand(String.Format("SELECT [Item N°], [P/N], [Designation], [Unit of purchase], [Qty purchase 1 year from  01/03/18 until 28/02/19] FROM [{0}] Where [Unit of purchase] = 'EA'", worksheets), oledbConn);
                        List<ExcelData> dataObject = new List<ExcelData>();

                        #region using reader
                        //using (var reader = cmd.ExecuteReader( CommandBehavior.SequentialAccess))
                        //{

                        //    //while (reader.Read())
                        //    //{
                        //    //    dataObject.Add(new ExcelData
                        //    //    {
                        //    //        ID = reader["ID"]?.ToString(),
                        //    //        Designation = reader["Designation"]?.ToString(),
                        //    //        Qty_of_purchase = reader.GetValue(4)?.ToString()//["Qty purchase 1 year from  01/03/18 until 28/02/19"].Cast<string>()

                        //    //    });
                        //    //}

                        //    dt.Load(reader);


                        //}

                        //var count = dataObject.Count;


                        #endregion

                        #region using adapter
                        OleDbDataAdapter oleda = new OleDbDataAdapter();
                        oleda.SelectCommand = cmd;
                        ds.EnforceConstraints = false;

                        oleda.Fill(ds);

                        countData += ds.Tables[0].Rows.Count;

                        #endregion
                        //DataSet dataSet = new DataSet();
                        //oleda.Fill(dataSet);

                        //}

                        //dt = ds.Tables[0];
                    }

                }
                }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {

                oledbConn.Close();
            }
         
            return dt;

        }
      
        public class ExcelData
        {
            public string Designation { get; set; }
            public string Qty_of_purchase { get; set; }
            public string ID { get; internal set; }
        }
    }
}