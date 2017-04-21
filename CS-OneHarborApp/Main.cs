using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CS_OneHarborApp
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
            FileLocationTextbox.Text = Properties.Settings.Default.FolderLocation;
            PartnerFilenameTextbox.Text = Properties.Settings.Default.PartnerFileName;
            DepartmentFilenameTextbox.Text = Properties.Settings.Default.PastorFileName;
            CreateOversightSheetsCheckBox.Checked = Properties.Settings.Default.MakeOversightPastorFiles;
        }

        private void StartButton_Click(object sender, EventArgs e)
        {
            var runThis = new Program();
            runThis.ConvertAndCombine();
        }

        private void SaveButton_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.FolderLocation = FileLocationTextbox.Text;
            Properties.Settings.Default.PartnerFileName = PartnerFilenameTextbox.Text;
            Properties.Settings.Default.PastorFileName = DepartmentFilenameTextbox.Text;
            Properties.Settings.Default.MakeOversightPastorFiles = CreateOversightSheetsCheckBox.Checked;
            Properties.Settings.Default.Save();
        }

        private void FileLocationFolderDialogButton_Click(object sender, EventArgs e)
        {
            var openFolderDialog = new FolderBrowserDialog();

            if (openFolderDialog.ShowDialog() == DialogResult.OK)
            {
                FileLocationTextbox.Text = openFolderDialog.SelectedPath;
            }
        }

        private void PartnerFileNameFileDialogButton_Click(object sender, EventArgs e)
        {
            var openFileDialog = new OpenFileDialog() { Multiselect = false };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                PartnerFilenameTextbox.Text = openFileDialog.FileName;
            }
        }

        private void DepartmentFileNameFileDialogButton_Click(object sender, EventArgs e)
        {
            var openFileDialog = new OpenFileDialog() { Multiselect = false };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                DepartmentFilenameTextbox.Text = openFileDialog.FileName;
            }
        }
    }
}
