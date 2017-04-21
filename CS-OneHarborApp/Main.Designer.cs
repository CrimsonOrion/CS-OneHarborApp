namespace CS_OneHarborApp
{
    partial class Main
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Main));
            this.StartButton = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.FileLocationTextbox = new System.Windows.Forms.TextBox();
            this.PartnerFilenameTextbox = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.DepartmentFilenameTextbox = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.CreateOversightSheetsCheckBox = new System.Windows.Forms.CheckBox();
            this.label4 = new System.Windows.Forms.Label();
            this.SaveButton = new System.Windows.Forms.Button();
            this.FileLocationFolderDialogButton = new System.Windows.Forms.Button();
            this.PartnerFileNameFileDialogButton = new System.Windows.Forms.Button();
            this.DepartmentFileNameFileDialogButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // StartButton
            // 
            this.StartButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.StartButton.Location = new System.Drawing.Point(268, 147);
            this.StartButton.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.StartButton.Name = "StartButton";
            this.StartButton.Size = new System.Drawing.Size(150, 50);
            this.StartButton.TabIndex = 0;
            this.StartButton.Text = "Start";
            this.StartButton.UseVisualStyleBackColor = true;
            this.StartButton.Click += new System.EventHandler(this.StartButton_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(103, 20);
            this.label1.TabIndex = 1;
            this.label1.Text = "File Location:";
            // 
            // FileLocationTextbox
            // 
            this.FileLocationTextbox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.FileLocationTextbox.Location = new System.Drawing.Point(122, 10);
            this.FileLocationTextbox.Name = "FileLocationTextbox";
            this.FileLocationTextbox.Size = new System.Drawing.Size(258, 26);
            this.FileLocationTextbox.TabIndex = 2;
            // 
            // PartnerFilenameTextbox
            // 
            this.PartnerFilenameTextbox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.PartnerFilenameTextbox.Location = new System.Drawing.Point(186, 46);
            this.PartnerFilenameTextbox.Name = "PartnerFilenameTextbox";
            this.PartnerFilenameTextbox.Size = new System.Drawing.Size(194, 26);
            this.PartnerFilenameTextbox.TabIndex = 4;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(13, 49);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(134, 20);
            this.label2.TabIndex = 3;
            this.label2.Text = "Partner Filename:";
            // 
            // DepartmentFilenameTextbox
            // 
            this.DepartmentFilenameTextbox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.DepartmentFilenameTextbox.Location = new System.Drawing.Point(186, 82);
            this.DepartmentFilenameTextbox.Name = "DepartmentFilenameTextbox";
            this.DepartmentFilenameTextbox.Size = new System.Drawing.Size(194, 26);
            this.DepartmentFilenameTextbox.TabIndex = 6;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(13, 85);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(167, 20);
            this.label3.TabIndex = 5;
            this.label3.Text = "Department Filename:";
            // 
            // CreateOversightSheetsCheckBox
            // 
            this.CreateOversightSheetsCheckBox.AutoSize = true;
            this.CreateOversightSheetsCheckBox.Location = new System.Drawing.Point(285, 125);
            this.CreateOversightSheetsCheckBox.Name = "CreateOversightSheetsCheckBox";
            this.CreateOversightSheetsCheckBox.Size = new System.Drawing.Size(15, 14);
            this.CreateOversightSheetsCheckBox.TabIndex = 7;
            this.CreateOversightSheetsCheckBox.UseVisualStyleBackColor = true;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(13, 121);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(266, 20);
            this.label4.TabIndex = 8;
            this.label4.Text = "Create Oversight Pastor Workbooks:";
            // 
            // SaveButton
            // 
            this.SaveButton.Location = new System.Drawing.Point(17, 147);
            this.SaveButton.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.SaveButton.Name = "SaveButton";
            this.SaveButton.Size = new System.Drawing.Size(150, 50);
            this.SaveButton.TabIndex = 9;
            this.SaveButton.Text = "Save Settings";
            this.SaveButton.UseVisualStyleBackColor = true;
            this.SaveButton.Click += new System.EventHandler(this.SaveButton_Click);
            // 
            // FileLocationFolderDialogButton
            // 
            this.FileLocationFolderDialogButton.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.FileLocationFolderDialogButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.FileLocationFolderDialogButton.Location = new System.Drawing.Point(386, 10);
            this.FileLocationFolderDialogButton.Name = "FileLocationFolderDialogButton";
            this.FileLocationFolderDialogButton.Size = new System.Drawing.Size(32, 26);
            this.FileLocationFolderDialogButton.TabIndex = 10;
            this.FileLocationFolderDialogButton.Text = "...";
            this.FileLocationFolderDialogButton.UseVisualStyleBackColor = true;
            this.FileLocationFolderDialogButton.Click += new System.EventHandler(this.FileLocationFolderDialogButton_Click);
            // 
            // PartnerFileNameFileDialogButton
            // 
            this.PartnerFileNameFileDialogButton.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.PartnerFileNameFileDialogButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.PartnerFileNameFileDialogButton.Location = new System.Drawing.Point(386, 46);
            this.PartnerFileNameFileDialogButton.Name = "PartnerFileNameFileDialogButton";
            this.PartnerFileNameFileDialogButton.Size = new System.Drawing.Size(32, 26);
            this.PartnerFileNameFileDialogButton.TabIndex = 11;
            this.PartnerFileNameFileDialogButton.Text = "...";
            this.PartnerFileNameFileDialogButton.UseVisualStyleBackColor = true;
            this.PartnerFileNameFileDialogButton.Click += new System.EventHandler(this.PartnerFileNameFileDialogButton_Click);
            // 
            // DepartmentFileNameFileDialogButton
            // 
            this.DepartmentFileNameFileDialogButton.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.DepartmentFileNameFileDialogButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.DepartmentFileNameFileDialogButton.Location = new System.Drawing.Point(386, 82);
            this.DepartmentFileNameFileDialogButton.Name = "DepartmentFileNameFileDialogButton";
            this.DepartmentFileNameFileDialogButton.Size = new System.Drawing.Size(32, 26);
            this.DepartmentFileNameFileDialogButton.TabIndex = 12;
            this.DepartmentFileNameFileDialogButton.Text = "...";
            this.DepartmentFileNameFileDialogButton.UseVisualStyleBackColor = true;
            this.DepartmentFileNameFileDialogButton.Click += new System.EventHandler(this.DepartmentFileNameFileDialogButton_Click);
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(434, 211);
            this.Controls.Add(this.DepartmentFileNameFileDialogButton);
            this.Controls.Add(this.PartnerFileNameFileDialogButton);
            this.Controls.Add(this.FileLocationFolderDialogButton);
            this.Controls.Add(this.SaveButton);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.CreateOversightSheetsCheckBox);
            this.Controls.Add(this.DepartmentFilenameTextbox);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.PartnerFilenameTextbox);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.FileLocationTextbox);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.StartButton);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.MinimumSize = new System.Drawing.Size(450, 250);
            this.Name = "Main";
            this.Text = "One Harbor Partner Services";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button StartButton;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox FileLocationTextbox;
        private System.Windows.Forms.TextBox PartnerFilenameTextbox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox DepartmentFilenameTextbox;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.CheckBox CreateOversightSheetsCheckBox;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button SaveButton;
        private System.Windows.Forms.Button FileLocationFolderDialogButton;
        private System.Windows.Forms.Button PartnerFileNameFileDialogButton;
        private System.Windows.Forms.Button DepartmentFileNameFileDialogButton;
    }
}

