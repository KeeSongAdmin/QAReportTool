using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace QAReportTool
{

    public class WriteExcel
    {
        Microsoft.Office.Interop.Excel.Application oXL;
        Microsoft.Office.Interop.Excel._Workbook oWB;
        Microsoft.Office.Interop.Excel._Worksheet oSheet;
        Microsoft.Office.Interop.Excel.Range oRng;
        object misvalue = System.Reflection.Missing.Value;

        public void writeExcel()
        {
            try
            {
                //Start Excel and get Application object.
                oXL = new Microsoft.Office.Interop.Excel.Application();
                oXL.Visible = true;

                //Get a new workbook.
                oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
                oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

                //Add table headers going cell by cell.
                oSheet.Cells[1, 1] = "First Name";
                oSheet.Cells[1, 2] = "Last Name";
                oSheet.Cells[1, 3] = "Full Name";
                oSheet.Cells[1, 4] = "Salary";

                //Format A1:D1 as bold, vertical alignment = center.
                oSheet.get_Range("A1", "Z1").Font.Bold = true;
                oSheet.get_Range("A1", "Z1").VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

                // Create an array to multiple values at once.
                string[,] saNames = new string[5, 2];

                saNames[0, 0] = "John";
                saNames[0, 1] = "Smith";
                saNames[1, 0] = "Tom";

                saNames[4, 1] = "Johnson";

                //Fill A2:B6 with an array of values (First and Last Names).
                oSheet.get_Range("A2", "B6").Value2 = saNames;

                //Fill C2:C6 with a relative formula (=A2 & " " & B2).
                oRng = oSheet.get_Range("C2", "C6");
                oRng.Formula = "=A2 & \" \" & B2";

                //Fill D2:D6 with a formula(=RAND()*100000) and apply format.
                oRng = oSheet.get_Range("D2", "D6");
                oRng.Formula = "=RAND()*100000";
                oRng.NumberFormat = "$0.00";

                //AutoFit columns A:D.
                oRng = oSheet.get_Range("A1", "D1");
                oRng.EntireColumn.AutoFit();

                oXL.Visible = false;
                oXL.UserControl = false;
                oXL.DisplayAlerts = false;
                oSheet.Select(Type.Missing);
                oWB.SaveAs("c:\\Test\\thuhuzz.xls");




                oWB.Close(0);
                oXL.Quit();

            }
            catch (Exception ex)
            {

            }
            finally
            {

            }




        }

        #region backup
        /*
        public void writeExcel()
        {
            try
            {
                //Start Excel and get Application object.
                oXL = new Microsoft.Office.Interop.Excel.Application();
                oXL.Visible = true;

                //Get a new workbook.
                oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
                oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

                //Add table headers going cell by cell.
                oSheet.Cells[1, 1] = "First Name";
                oSheet.Cells[1, 2] = "Last Name";
                oSheet.Cells[1, 3] = "Full Name";
                oSheet.Cells[1, 4] = "Salary";

                //Format A1:D1 as bold, vertical alignment = center.
                oSheet.get_Range("A1", "D1").Font.Bold = true;
                oSheet.get_Range("A1", "D1").VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

                // Create an array to multiple values at once.
                string[,] saNames = new string[5, 2];

                saNames[0, 0] = "John";
                saNames[0, 1] = "Smith";
                saNames[1, 0] = "Tom";

                saNames[4, 1] = "Johnson";

                //Fill A2:B6 with an array of values (First and Last Names).
                oSheet.get_Range("A2", "B6").Value2 = saNames;

                //Fill C2:C6 with a relative formula (=A2 & " " & B2).
                oRng = oSheet.get_Range("C2", "C6");
                oRng.Formula = "=A2 & \" \" & B2";

                //Fill D2:D6 with a formula(=RAND()*100000) and apply format.
                oRng = oSheet.get_Range("D2", "D6");
                oRng.Formula = "=RAND()*100000";
                oRng.NumberFormat = "$0.00";

                //AutoFit columns A:D.
                oRng = oSheet.get_Range("A1", "D1");
                oRng.EntireColumn.AutoFit();

                oXL.Visible = false;
                oXL.UserControl = false;
                oXL.DisplayAlerts = false;
                oSheet.Select(Type.Missing);
                oWB.SaveAs("c:\\Test\\thuhuzz.xls");




                oWB.Close(0);
                oXL.Quit();

            }
            catch (Exception ex)
            {

            }
            finally
            {

            }




        }
        */
        #endregion

        public void writeExcel(DataTable dt_source)
        {
            try
            {
                Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

                if (xlApp == null)
                {
                    Console.WriteLine("Excel is not properly installed!!");
                    return;
                }

                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);


                #region header row

                for (int j = 0; j < dt_source.Columns.Count; j++)
                {
                    xlWorkSheet.Cells[1, j + 1] = dt_source.Columns[j].ColumnName.ToString();
                }

                //Format A1:D1 as bold, vertical alignment = center.
                xlWorkSheet.get_Range("A1", "Z1").Font.Bold = true;
                xlWorkSheet.get_Range("A1", "Z1").VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;


                #endregion



                for (int i = 0; i < dt_source.Rows.Count; i++)
                {
                    for (int j = 0; j < dt_source.Columns.Count; j++)
                    {
                        xlWorkSheet.Cells[i + 2, j + 1] = dt_source.Rows[i][j].ToString();// == "" ? "0" : dt_source.Rows[i][j].ToString();
                    }
                }


                xlWorkSheet.get_Range("A1", "ZZ2000").EntireColumn.AutoFit();

                string savepath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                xlWorkBook.SaveAs(savepath + @"\ECom Picking Report Category - " + DateTime.Now.ToString("yyyyMMMddHHmmss") + ".xlsx");
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);

                //Console.WriteLine("Excel file created , you can find the file in " + ConfigurationManager.AppSettings["OutputPath"].ToString() + @"PickingList-" + DateTime.Now.ToString("yyyyMMMM") + ".xls");

            }
            catch (Exception ex)
            {

            }
            finally
            {

            }




        }

        public void writeExcelMultiSheet(DataTable dt_source1, DataTable dt_source2, DataTable dt_source3, DataTable dt_source4, DataTable dt_source5)
        {
            try
            {
                Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

                if (xlApp == null)
                {
                    Console.WriteLine("Excel is not properly installed!!");
                    return;
                }

                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                Excel.Sheets worksheets = xlWorkBook.Worksheets;

                #region First Sheet
                DataTable dt_temp = dt_source1;
                #region header row
                var xlNewSheet = (Excel.Worksheet)worksheets.Add(Type.Missing, worksheets[1], Type.Missing, Type.Missing);
                xlNewSheet.Name = "Summary";

                xlNewSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);

                for (int j = 0; j < dt_temp.Columns.Count; j++)
                {
                    xlNewSheet.Cells[1, j + 1] = dt_temp.Columns[j].ColumnName.ToString();
                }

                //Format A1:D1 as bold, vertical alignment = center.
                xlNewSheet.get_Range("A1", "Z1").Font.Bold = true;
                xlNewSheet.get_Range("A1", "Z1").VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;


                #endregion
                #region detail
                for (int i = 0; i < dt_temp.Rows.Count; i++)
                {
                    for (int j = 0; j < dt_temp.Columns.Count; j++)
                    {
                        xlNewSheet.Cells[i + 2, j + 1] = dt_temp.Rows[i][j].ToString();// == "" ? "0" : dt_source.Rows[i][j].ToString();
                    }
                }
                #endregion
                #endregion

                #region Second Sheet
                dt_temp = dt_source2;
                #region header row
                xlNewSheet = (Excel.Worksheet)worksheets.Add(Type.Missing, worksheets[2], Type.Missing, Type.Missing);
                xlNewSheet.Name = "FRESH";

                xlNewSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(3);

                for (int j = 0; j < dt_temp.Columns.Count; j++)
                {
                    xlNewSheet.Cells[1, j + 1] = dt_temp.Columns[j].ColumnName.ToString();
                }

                //Format A1:D1 as bold, vertical alignment = center.
                xlNewSheet.get_Range("A1", "Z1").Font.Bold = true;
                xlNewSheet.get_Range("A1", "Z1").VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;


                #endregion
                #region detail
                for (int i = 0; i < dt_temp.Rows.Count; i++)
                {
                    for (int j = 0; j < dt_temp.Columns.Count; j++)
                    {
                        xlNewSheet.Cells[i + 2, j + 1] = dt_temp.Rows[i][j].ToString();// == "" ? "0" : dt_source.Rows[i][j].ToString();
                    }
                }
                #endregion
                #endregion

                #region Third Sheet
                dt_temp = dt_source3;
                #region header row
                xlNewSheet = (Excel.Worksheet)worksheets.Add(Type.Missing, worksheets[3], Type.Missing, Type.Missing);
                xlNewSheet.Name = "Frozen";

                xlNewSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(4);

                for (int j = 0; j < dt_temp.Columns.Count; j++)
                {
                    xlNewSheet.Cells[1, j + 1] = dt_temp.Columns[j].ColumnName.ToString();
                }

                //Format A1:D1 as bold, vertical alignment = center.
                xlNewSheet.get_Range("A1", "Z1").Font.Bold = true;
                xlNewSheet.get_Range("A1", "Z1").VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;


                #endregion
                #region detail
                for (int i = 0; i < dt_temp.Rows.Count; i++)
                {
                    for (int j = 0; j < dt_temp.Columns.Count; j++)
                    {
                        xlNewSheet.Cells[i + 2, j + 1] = dt_temp.Rows[i][j].ToString();// == "" ? "0" : dt_source.Rows[i][j].ToString();
                    }
                }
                #endregion
                #endregion

                #region Fourth Sheet
                dt_temp = dt_source4;
                #region header row
                xlNewSheet = (Excel.Worksheet)worksheets.Add(Type.Missing, worksheets[4], Type.Missing, Type.Missing);
                xlNewSheet.Name = "FROZEN THAWED";

                xlNewSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(5);

                for (int j = 0; j < dt_temp.Columns.Count; j++)
                {
                    xlNewSheet.Cells[1, j + 1] = dt_temp.Columns[j].ColumnName.ToString();
                }

                //Format A1:D1 as bold, vertical alignment = center.
                xlNewSheet.get_Range("A1", "Z1").Font.Bold = true;
                xlNewSheet.get_Range("A1", "Z1").VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;


                #endregion
                #region detail
                for (int i = 0; i < dt_temp.Rows.Count; i++)
                {
                    for (int j = 0; j < dt_temp.Columns.Count; j++)
                    {
                        xlNewSheet.Cells[i + 2, j + 1] = dt_temp.Rows[i][j].ToString();// == "" ? "0" : dt_source.Rows[i][j].ToString();
                    }
                }
                #endregion
                #endregion

                #region Fifth Sheet
                dt_temp = dt_source5;
                #region header row
                xlNewSheet = (Excel.Worksheet)worksheets.Add(Type.Missing, worksheets[5], Type.Missing, Type.Missing);
                xlNewSheet.Name = "DRY";

                xlNewSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(6);

                for (int j = 0; j < dt_temp.Columns.Count; j++)
                {
                    xlNewSheet.Cells[1, j + 1] = dt_temp.Columns[j].ColumnName.ToString();
                }

                //Format A1:D1 as bold, vertical alignment = center.
                xlNewSheet.get_Range("A1", "Z1").Font.Bold = true;
                xlNewSheet.get_Range("A1", "Z1").VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;


                #endregion
                #region detail
                for (int i = 0; i < dt_temp.Rows.Count; i++)
                {
                    for (int j = 0; j < dt_temp.Columns.Count; j++)
                    {
                        xlNewSheet.Cells[i + 2, j + 1] = dt_temp.Rows[i][j].ToString();// == "" ? "0" : dt_source.Rows[i][j].ToString();
                    }
                }
                #endregion
                #endregion 

                foreach (Excel.Worksheet wrkst in xlWorkBook.Worksheets)
                {
                    Excel.Range usedrange = wrkst.UsedRange;
                    usedrange.Columns.AutoFit();
                }

                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                xlApp.DisplayAlerts = false;
                xlWorkSheet.Delete();
                xlApp.DisplayAlerts = true;

                string savepath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                xlWorkBook.SaveAs(savepath + @"\Product Summary Report - " + DateTime.Now.ToString("yyyyMMdd HHtt") + ".xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);

                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);

                MessageBox.Show("Excel file created , you can find the file in your desktop Product Summary Report - " + DateTime.Now.ToString("yyyyMMdd HHtt") + ".xls");

            }
            catch (Exception ex)
            {

            }
            finally
            {

            }




        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }

        public void AppendExcelSheet(String path, DataTable dt_input)
        {
            try
            {
                Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

                if (xlApp == null)
                {
                    MessageBox.Show("Excel is not properly installed!!");
                    return;
                }

                xlApp.DisplayAlerts = false;
                string filePath = path;
                Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(filePath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                Excel.Sheets worksheets = xlWorkBook.Worksheets;

                //MessageBox.Show("0");
                #region details
                var xlNewSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                for (int i = 0; i < dt_input.Rows.Count; i++)
                {
                    xlNewSheet.Cells[i + 16, 2] = dt_input.Rows[i][8].ToString();
                    xlNewSheet.Cells[i + 16, 15] = dt_input.Rows[i][3].ToString();
                }
                #endregion

                String oripath = path;
                //string[] oripaths = oripath.Split('.');
                //oripath = oripath.Replace(oripaths[oripaths.Length - 1], "- RESULT " + DateTime.Now.ToString("yyyyMMMddHHmmss") +"."+ oripaths[oripaths.Length - 1]);
                //xlWorkBook.SaveAs(oripath); 
                string savepath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                xlWorkBook.SaveAs(savepath + @"\Driver Report - " + DateTime.Now.ToString("yyyyMMMddHHmmss") + ".xlsx");

                //xlWorkBook.SaveAs(@"C:\Test\sresult12345677777.xls");
                xlWorkBook.Close();

                releaseObject(xlNewSheet);
                releaseObject(worksheets);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);

                MessageBox.Show("New Worksheet Created!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("EXCEL CLASS ERROR");
            }

        }



        public void AppendExcelMultipleSheet(String path, String outFullFilePath, List<DataTable> dt_inputList)
        {
            try
            {
                Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

                if (xlApp == null)
                {
                    MessageBox.Show("Excel is not properly installed!!");
                    return;
                }

                xlApp.DisplayAlerts = false;
                string filePath = path;
                Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(filePath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                Excel.Sheets worksheets = xlWorkBook.Worksheets;


                for (int p = 0; p < dt_inputList.Count; p++)
                {
                    DataTable dt_input = dt_inputList[p];
                    #region details
                    var xlNewSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(p + 1);

                    List<string> itemQADescList = dt_input.AsEnumerable().Select(x => string.Format("{0}", x[4])).ToList();
                    List<string> itemQACodeList = dt_input.AsEnumerable().Select(x => string.Format("{0}", x[5])).ToList();

                    String productName = String.Join(Environment.NewLine, itemQADescList);
                    String productCode = String.Join(Environment.NewLine, itemQACodeList);

                    xlNewSheet.Cells[2, 2] = productName;
                    xlNewSheet.Cells[3, 2] = productCode;
                    for (int i = 0; i < dt_input.Rows.Count; i++)
                    {
                        xlNewSheet.Cells[i + 4, 2] = dt_input.Rows[i][6].ToString();
                        xlNewSheet.Cells[i + 4, 3] = dt_input.Rows[i][7].ToString();
                    }
                    #endregion
                    releaseObject(xlNewSheet);
                }

                xlWorkBook.SaveAs(outFullFilePath);
                xlWorkBook.Close();

                //releaseObject(xlNewSheet);
                releaseObject(worksheets);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);

                MessageBox.Show("New Worksheet Created!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("EXCEL CLASS ERROR");
            }

        }

        public void AppendExcelMultipleSheetBySheetIndex(String path, String outFullFilePath, DataTable dt_input, int rSheetIndex)
        {
            try
            {
                Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

                if (xlApp == null)
                {
                    MessageBox.Show("Excel is not properly installed!!");
                    return;
                }

                xlApp.DisplayAlerts = false;
                string filePath = path;
                Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(filePath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                Excel.Sheets worksheets = xlWorkBook.Worksheets;

                #region details
                var xlNewSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(rSheetIndex);

                List<string> itemQADescList = dt_input.AsEnumerable().Select(x => string.Format("{0}", x[4])).ToList();
                List<string> itemQACodeList = dt_input.AsEnumerable().Select(x => string.Format("{0}", x[5])).ToList();

                String productName = String.Join(Environment.NewLine, itemQADescList);
                String productCode = String.Join(Environment.NewLine, itemQACodeList);

                xlNewSheet.Cells[2, 2] = productName;
                xlNewSheet.Cells[3, 2] = productCode;
                for (int i = 0; i < dt_input.Rows.Count; i++)
                {
                    xlNewSheet.Cells[i + 4, 2] = dt_input.Rows[i][6].ToString();
                    xlNewSheet.Cells[i + 4, 3] = dt_input.Rows[i][7].ToString();
                }
                #endregion


                xlWorkBook.SaveAs(outFullFilePath);
                xlWorkBook.Close();

                releaseObject(xlNewSheet);
                releaseObject(worksheets);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);

                //MessageBox.Show("New Worksheet Created!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("EXCEL CLASS ERROR");
            }

        }

    }
}
