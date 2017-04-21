using CustomExtensions;
using System;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace CS_OneHarborApp
{
    public class ExcelUtility
    {
        public static DataTable FillDataTable(string type, string file, string tableName, string query = null, string headerRow = "No", string sortColumn = "")        
        {
            OleDbConnection odcCN = null;
            OleDbDataAdapter odcDA = null;
            DataTable datatable = null;
            try
            {
                // Use the correct connection string based on the type, whether its .xlsx (Excel 2010+), .csv, etc.  Go to http://www.connectionstrings.com/ for a full list.
                switch (type)
                {
                    case ".xlsx":
                        odcCN = new OleDbConnection("provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + file + ";Extended Properties='Excel 12.0;HDR=" + headerRow + ";IMEX=1;'");
                        break;
                    case ".xls":
                        odcCN = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + file + ";Extended Properties='Excel 8.0;HDR=" + headerRow + ";IMEX=1;'");
                        break;
                    case ".csv":
                        odcCN = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + file + ";Extended Properties=text;HDR=" + headerRow + ";FMT=Delimited;");
                        break;                    
                }

                odcCN.Open();

                // This is used to get the sheet name if you don't know it.
                System.Data.DataTable dt = odcCN.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                string sheetName = dt.Rows[0].ItemArray[2].ToString();
                query = "SELECT * FROM [" + sheetName + "]";

                odcDA = new OleDbDataAdapter(query, odcCN); // Use the dataAdapter to get the data from the selected Connection
                datatable = new DataTable();                // Create a dataTable to put all the data from the dataAdapter into
                odcDA.Fill(datatable);                      // Fill the dataTable with the pulled info from the dataAdapter
                odcCN.Close();                              // Close the connection once the data is imported

                // Get the correct column names from the excel sheets

                for (int titleIndex = 0; titleIndex < datatable.Rows.Count; titleIndex++)
                {
                    if (datatable.Rows[titleIndex].ItemArray[0].ToString() == "Cracks" || (datatable.Rows[titleIndex].ItemArray[0].ToString().Trim()) == "" || datatable.Rows[titleIndex].ItemArray[0].ToString() == "Department Summary")
                    {
                        datatable.Rows.RemoveAt(titleIndex);
                        titleIndex--;
                    }
                    else if (datatable.Rows[titleIndex].ItemArray[0].ToString() == "Last Name")
                    {
                        for (int column = 0; column < datatable.Columns.Count; column++)
                        {
                            if (datatable.Rows[titleIndex].ItemArray[column].ToString() == "")
                            {
                                datatable.Columns.RemoveAt(column);
                                column--;
                            }
                            else
                            {
                                datatable.Columns[column].ColumnName = datatable.Rows[titleIndex].ItemArray[column].ToString();
                            }                            
                        }
                        datatable.Rows.RemoveAt(titleIndex);
                        break;
                    }
                    else if (datatable.Rows[titleIndex].ItemArray[0].ToString() == "Department")
                    {
                        for (int column = 0; column < datatable.Columns.Count; column++)
                        {
                            if (datatable.Rows[titleIndex].ItemArray[column].ToString() == "")
                            {
                                datatable.Columns.RemoveAt(column);
                                column--;
                            }
                            else
                            {
                                datatable.Columns[column].ColumnName = datatable.Rows[titleIndex].ItemArray[column].ToString();
                            }                            
                        }
                        datatable.Rows.RemoveAt(titleIndex);
                        break;
                    }
                }
                datatable.TableName = tableName;
                return datatable;

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);                // If there's an error, display the error message.  ** POSSIBLE FUTURE UPDATE: Automatically email Jim a copy of the error message? **
                return null;
            }
        }

        public static void CreateExcel(DataSet ds, string tableName, string excelPath, bool headerRow = false)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            try
            {
                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                if (headerRow)
                {
                    for (int h = 0; h < ds.Tables[tableName].Columns.Count; h++)
                    {
                        xlWorkSheet.Cells[1, h + 1] = ds.Tables[tableName].Columns[h].ColumnName;
                    }
                }
                for (int i = 0; i < ds.Tables[tableName].Rows.Count; i++)
                {
                    for (int j = 0; j < ds.Tables[tableName].Columns.Count; j++)
                    {
                        if (headerRow)
                        {
                            xlWorkSheet.Cells[i + 2, j + 1] = ds.Tables[tableName].Rows[i].ItemArray[j].ToString();
                        }
                        else
                        {
                            xlWorkSheet.Cells[i + 1, j + 1] = ds.Tables[tableName].Rows[i].ItemArray[j].ToString();
                        }                        
                    }
                }

                // For .xls, use this one \|/  Otherwise, for .xlsx, use other                
                xlApp.DisplayAlerts = false;
                xlWorkBook.SaveAs(excelPath);
                xlApp.DisplayAlerts = true;
                xlWorkBook.Close();
                xlApp.Quit();

                ReleaseObject(xlApp);
                ReleaseObject(xlWorkBook);
                ReleaseObject(xlWorkSheet);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        public static void Fix255Rule(string fileName)
        {
            try
            {
                var xlApp = new Excel.Application();
                var xlBook = xlApp.Workbooks.Open(fileName);
                var xlSheet = (Excel.Worksheet)xlBook.Worksheets[1];                

                for (int row = 1; row <= 10; row++)
                {                    
                    if (xlSheet.GetCellValue(row, 1) == "Last Name")
                    {
                        for (int column = 1; column <= 20; column++)
                        {                            
                            if (xlSheet.GetCellValue(row, column) == "Group - All Non-administrative")
                            {
                                //((Excel.Range)xlSheet.Cells[row, column]).Value2 = "Group - All Non-administrative                                                                                                                                                                                                                                    '";
                                xlSheet.SetCellValue(row, column, "string", "Group - All Non-administrative                                                                                                                                                                                                                                    '");
                                xlBook.Save();
                                xlBook.Close();
                                xlApp.Quit();

                                ReleaseObject(xlSheet);
                                ReleaseObject(xlBook);
                                ReleaseObject(xlApp);

                                return;
                            }
                            else if (xlSheet.GetCellValue(row, column) == "Group - All Non-administrative                                                                                                                                                                                                                                    '")
                            {
                                xlBook.Close();
                                xlApp.Quit();

                                ReleaseObject(xlSheet);
                                ReleaseObject(xlBook);
                                ReleaseObject(xlApp);

                                return;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error in Fix255Rule: " + ex.Message);
            }
        }

        public static string GetOversight(string communityGroup)
        {
            string oversightPastor = "None Found";

            for (int index = 0; index < Program.DeptSumTable.Rows.Count; index++)
            {
                if (communityGroup.Contains("Community Group Leaders"))
                {
                    oversightPastor = "Scott Beierwaltes";
                }
                else if(communityGroup == Program.DeptSumTable.Rows[index].ItemArray[2].ToString())
                {
                    oversightPastor = Program.DeptSumTable.Rows[index].ItemArray[5].ToString();
                    break;
                }
            }

            return oversightPastor;
        }

        public static void MakeOversightPastorExcelWorkbooks()
        {
            try
            {
                // Get each oversight pastor name
                var oversightPastorName = (from pastor in Program.MemberInfoTable.AsEnumerable()
                                           where pastor.ItemArray[7].ToString() != "" 
                                           select pastor.ItemArray[7].ToString()).ToList().Distinct();

                // Get each person's info who has the current oversightPastorName
                foreach (var pastor in oversightPastorName)
                {
                    var oversightPastor = (from T in Program.MemberInfoTable.AsEnumerable()
                                           where T.ItemArray[7].ToString() == pastor.ToString()
                                           select T).ToList();

                    DataTable table = new DataTable();
                    table = Program.MemberInfoTable.Clone();
                    table.TableName = "Temp";
                    DataSet temp = new DataSet("Temp");
                    temp.Tables.Add(table);

                    if (oversightPastor != null)
                    {
                        foreach (DataRow row in oversightPastor)
                        {
                            table.Rows.Add(row.ItemArray);
                        }
                    }

                    CreateExcel(temp, "Temp", Properties.Settings.Default.FolderLocation + @"\" + pastor, true);
                    
                    table.Dispose();
                    temp.Dispose();
                }
                System.Windows.Forms.MessageBox.Show("Oversight Pastor sheets successfully created in " + Properties.Settings.Default.FolderLocation + ".");
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error during Oversight Pastor Sheet creation." + @"/n/n" + "Error is as follows: " + ex.Message + @"/n/n" + "Please let Jim Lynch know about this error.", "Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }

        }

        private static void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
