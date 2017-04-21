using System;
using System.Data;
using System.Windows.Forms;

namespace CS_OneHarborApp
{
    public class Program
    {
        // public static DataSet OHDataSet = new DataSet("OHFiles");
        public static DataSet OHDataSet { get; set; } = new DataSet("OHFiles");
        public static DataTable CustReptTable { get; set; }
        public static DataTable DeptSumTable { get; set; }
        public static DataTable MemberInfoTable { get; set; }            

        static string PartnerExcelFilePath
        {
            get { return Properties.Settings.Default.PartnerFileName; }
        }

        static string PastorExcelFilePath
        {
            get { return Properties.Settings.Default.PastorFileName; }
        }

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Main());            
        }

        public void ConvertAndCombine()
        {
            string newFileName = string.Empty;
            try
            {
                // Fix the Excel OleDB problem with 255 character limit
                ExcelUtility.Fix255Rule(PartnerExcelFilePath);

                // Fill the tables with the info from the excel files and the dataset with the tables
                PopulateTables();                

                if (MessageBox.Show("Name the new file.", "Name New File", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.Cancel)
                {
                    MessageBox.Show("File creation cancelled.", "Cancelled", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                newFileName = SaveFileAs();

                if (newFileName != "")
                {
                    ExcelUtility.CreateExcel(OHDataSet, "MemberInfo", newFileName, true);
                }
                else
                {
                    MessageBox.Show("File creation cancelled.", "Cancelled", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (Properties.Settings.Default.MakeOversightPastorFiles)
                {
                    ExcelUtility.MakeOversightPastorExcelWorkbooks();
                }

                OHDataSet.Tables.Clear();
                OHDataSet.Dispose();
                MessageBox.Show("All done!", "One Harbor Partner Services", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void PopulateTables()
        {
            try
            {
                CustReptTable = ExcelUtility.FillDataTable(".xlsx", PartnerExcelFilePath, "CustRept");
                DeptSumTable = ExcelUtility.FillDataTable(".xlsx", PastorExcelFilePath, "DeptSum");
                MemberInfoTable = FillNewTable("MemberInfo");
                OHDataSet.Tables.Add(CustReptTable);
                OHDataSet.Tables.Add(DeptSumTable);
                OHDataSet.Tables.Add(MemberInfoTable);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private DataTable FillNewTable(string tableName)
        {            
            try
            {
                DataTable table = new DataTable(tableName);
                table.Columns.Add("FullName");
                table.Columns.Add("CommunityGroup");
                table.Columns.Add("Hospitality");
                table.Columns.Add("KidsMinistry");
                table.Columns.Add("Youth");
                table.Columns.Add("MusicAndMedia");
                table.Columns.Add("Benevolence");
                table.Columns.Add("OversightPastor");
                table.Columns.Add("Community");
                table.Columns.Add("Serving");
                table.Columns.Add("Partner");
                table.Columns.Add("Site");
                table.Columns.Add("FirstStepsClass");
                table.Columns.Add("PartnershipClass");
                                
                string[] _site = null;                
                string[] stringArray = null;
                object[] custReptRowData = null;
                string[] separator = { " - " };

                for (int customReportRow = 0; customReportRow < CustReptTable.Rows.Count; customReportRow++)
                {
                    custReptRowData = CustReptTable.Rows[customReportRow].ItemArray;
                    MemberInfo member = new MemberInfo()
                    {
                        FullName = custReptRowData[1].ToString() + " " + custReptRowData[0].ToString()
                    };
                    stringArray = custReptRowData[3].ToString().Split(',');
                    for (int index = 0; index <= stringArray.Length - 1; index++)
                    {                        
                        if (stringArray[index].Contains("Community")) { member.Community = "Yes"; member.CommunityGroup = stringArray[index].TrimStart(' '); }
                        if (stringArray[index].Contains("Hospitality")) { member.Hospitality = "Yes"; }
                        if (stringArray[index].Contains("KM Volunteer")) { member.KidsMinistry = "Yes"; }
                        if (stringArray[index].Contains("Youth Leaders")) { member.Youth = "Yes"; }
                        if (stringArray[index].Contains("Music")) { member.MusicAndMedia = "Yes"; }
                        if (stringArray[index].Contains("Benevolence")) { member.Benevolence = "Yes"; }
                    }

                    if (member.Hospitality == "Yes" || member.KidsMinistry == "Yes" || member.Youth == "Yes" || member.MusicAndMedia == "Yes" || member.Benevolence == "Yes")
                    {
                        member.Serving = "Yes";
                    }
                    else
                    {
                        member.Serving = "No";
                    }

                    if (member.CommunityGroup != "None") { member.OversightPastor = ExcelUtility.GetOversight(member.CommunityGroup); }
                    
                    member.Partner = custReptRowData[6].ToString();
                    _site = custReptRowData[2].ToString().Split(separator, StringSplitOptions.RemoveEmptyEntries);
                    member.Site = _site[1].Trim();
                    member.FirstStepsDate = custReptRowData[4].ToString();                    
                    if (custReptRowData[5].ToString() == "-")
                    {
                        member.PartnershipClassDate = "-";
                    }
                    else
                    {
                        member.PartnershipClassDate = DateTime.Parse(custReptRowData[5].ToString()).ToString("MM/dd/yyyy");
                    }                    
                    table.Rows.Add(member.FullName, member.CommunityGroup, member.Hospitality, member.KidsMinistry, member.Youth, member.MusicAndMedia, member.Benevolence, member.OversightPastor, member.Community, member.Serving, member.Partner, member.Site, member.FirstStepsDate, member.PartnershipClassDate);
                }
                
                return table;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;                
            }
            
        }

        private string SaveFileAs()
        {
            try
            {
                SaveFileDialog SFDExcel = new SaveFileDialog()
                {
                    Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
                    FilterIndex = 1,
                    FileName = "MemberInfo.xlsx"
                };

                if (SFDExcel.ShowDialog() == DialogResult.Cancel)
                {
                    SFDExcel.FileName = null;
                }

                return SFDExcel.FileName;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;                
            }            
        }

        private string SelectFile()
        {
            OpenFileDialog ofdExcel = new OpenFileDialog()
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
                FilterIndex = 1,
                RestoreDirectory = false,
                Multiselect = false
            };

            if (ofdExcel.ShowDialog() == DialogResult.OK)
            {
                return ofdExcel.FileName;
            }
            else
            {
                return "Cancelled";
            }
        }
    }

    public class MemberInfo
    {
        public string FullName = string.Empty;
        public string CommunityGroup = "None";
        public string Hospitality = "No";
        public string KidsMinistry = "No";
        public string Youth = "No";
        public string MusicAndMedia = "No";
        public string Benevolence = "No";
        public string OversightPastor = "None";
        public string Community = "No";
        public string Serving = "No";
        public string Partner = string.Empty;
        public string Site = string.Empty;
        public string FirstStepsDate = string.Empty;
        public string PartnershipClassDate = string.Empty;
    }

}
