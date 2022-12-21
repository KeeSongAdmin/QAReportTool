using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace QAReportTool
{
    public class ReadExcel
    {
        public System.Data.DataTable ReadExcelFile(String path, String SheetName,int? SheetPosition, String query)
        {
            System.Data.DataTable dt_result = new System.Data.DataTable();
            Worksheet worksheet;
            Workbook theWorkbook;
            object misValue = System.Reflection.Missing.Value;
            String ConnectionString;
            OleDbDataAdapter objAdapter;
            //var workbooks = ExcelObj.Workbooks;
            //var workbook = workbooks.Open(filename);
            //worksheet = workbook.Worksheets[SheetName];
            //Marshal.ReleaseComObject(workbook);
            //Marshal.ReleaseComObject(workbooks);

            Microsoft.Office.Interop.Excel.Application ExcelObj = new Application();


            string filename = path;

            theWorkbook = ExcelObj.Workbooks.Open(filename, null, true);
            Microsoft.Office.Interop.Excel.Sheets sheets;
            sheets = theWorkbook.Worksheets;

            if (SheetName.Length == 0)
            {
                worksheet = sheets.get_Item(SheetPosition);
                SheetName = worksheet.Name;
                query=Regex.Replace(query,"@SheetName", SheetName);
            }



            try
            {
                ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filename + ";Extended Properties='Excel 12.0 Xml;HDR=No;IMEX=1;'";

                using (objAdapter = new OleDbDataAdapter(query, ConnectionString))
                {
                    objAdapter.Fill(dt_result);


                    objAdapter.Dispose();

                    if (theWorkbook != null)
                    {
                        theWorkbook.Close(false, misValue, misValue);
                        theWorkbook = null;
                        ExcelObj.Quit();
                        ExcelObj = null;
                    }

                }


            }
            catch (Exception ex)
            {
                Console.WriteLine("Excel Connection Fail");
            }
            finally
            {
                if (theWorkbook != null)
                {
                    theWorkbook.Close(false, misValue, misValue);
                    theWorkbook = null;
                    ExcelObj.Quit();
                    ExcelObj = null;
                }

            }


            return dt_result;
        }


        private static string GetTableName(string connectionString, int row = 0)
        {
            OleDbConnection conn = new OleDbConnection(connectionString);
            try
            {
                conn.Open();
                return conn.GetSchema("Tables").Rows[row]["TABLE_NAME"] + "";
            }
            catch { }
            finally { conn.Close(); }
            return "sheet1";
        }

    }
}
