using System.ComponentModel;
using System.Windows.Forms;

namespace VisioPowerTools2010
{
    partial class FormExportSelectionAsFormat : Form
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private IContainer components = null;

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
            this.labelDoc = new System.Windows.Forms.Label();
            this.labelPage = new System.Windows.Forms.Label();
            this.labelDocumentName = new System.Windows.Forms.Label();
            this.labelPageName = new System.Windows.Forms.Label();
            this.labelOutputFile = new System.Windows.Forms.Label();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.buttonOK = new System.Windows.Forms.Button();
            this.filenamePicker1 = new VisioAutomation.UI.FilenamePicker();
            this.labelFormat = new System.Windows.Forms.Label();
            this.labelFormatChoice = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // labelDoc
            // 
            this.labelDoc.AutoSize = true;
            this.labelDoc.Location = new System.Drawing.Point(22, 13);
            this.labelDoc.Name = "labelDoc";
            this.labelDoc.Size = new System.Drawing.Size(56, 13);
            this.labelDoc.TabIndex = 0;
            this.labelDoc.Text = "Document";
            // 
            // labelPage
            // 
            this.labelPage.AutoSize = true;
            this.labelPage.Location = new System.Drawing.Point(22, 30);
            this.labelPage.Name = "labelPage";
            this.labelPage.Size = new System.Drawing.Size(32, 13);
            this.labelPage.TabIndex = 1;
            this.labelPage.Text = "Page";
            // 
            // labelDocumentName
            // 
            this.labelDocumentName.AutoSize = true;
            this.labelDocumentName.Location = new System.Drawing.Point(101, 13);
            this.labelDocumentName.Name = "labelDocumentName";
            this.labelDocumentName.Size = new System.Drawing.Size(95, 13);
            this.labelDocumentName.TabIndex = 2;
            this.labelDocumentName.Text = "<document name>";
            // 
            // labelPageName
            // 
            this.labelPageName.AutoSize = true;
            this.labelPageName.Location = new System.Drawing.Point(101, 30);
            this.labelPageName.Name = "labelPageName";
            this.labelPageName.Size = new System.Drawing.Size(72, 13);
            this.labelPageName.TabIndex = 3;
            this.labelPageName.Text = "<page name>";
            // 
            // labelOutputFile
            // 
            this.labelOutputFile.AutoSize = true;
            this.labelOutputFile.Location = new System.Drawing.Point(19, 88);
            this.labelOutputFile.Name = "labelOutputFile";
            this.labelOutputFile.Size = new System.Drawing.Size(58, 13);
            this.labelOutputFile.TabIndex = 4;
            this.labelOutputFile.Text = "Output File";
            // 
            // buttonCancel
            // 
            this.buttonCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonCancel.Location = new System.Drawing.Point(397, 327);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(75, 23);
            this.buttonCancel.TabIndex = 1;
            this.buttonCancel.Text = "Cancel";
            this.buttonCancel.UseVisualStyleBackColor = true;
            this.buttonCancel.Click += new System.EventHandler(this.buttonCancel_Click);
            // 
            // buttonOK
            // 
            this.buttonOK.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonOK.Location = new System.Drawing.Point(316, 327);
            this.buttonOK.Name = "buttonOK";
            this.buttonOK.Size = new System.Drawing.Size(75, 23);
            this.buttonOK.TabIndex = 0;
            this.buttonOK.Text = "OK";
            this.buttonOK.UseVisualStyleBackColor = true;
            this.buttonOK.Click += new System.EventHandler(this.buttonOK_Click);
            // 
            // filenamePicker1
            // 
            this.filenamePicker1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.filenamePicker1.Filename = "";
            this.filenamePicker1.Location = new System.Drawing.Point(22, 104);
            this.filenamePicker1.Name = "filenamePicker1";
            this.filenamePicker1.ReadOnly = false;
            this.filenamePicker1.Size = new System.Drawing.Size(449, 106);
            this.filenamePicker1.TabIndex = 5;
            // 
            // labelFormat
            // 
            this.labelFormat.AutoSize = true;
            this.labelFormat.Location = new System.Drawing.Point(22, 47);
            this.labelFormat.Name = "labelFormat";
            this.labelFormat.Size = new System.Drawing.Size(72, 13);
            this.labelFormat.TabIndex = 6;
            this.labelFormat.Text = "Export Format";
            // 
            // labelFormatChoice
            // 
            this.labelFormatChoice.AutoSize = true;
            this.labelFormatChoice.Location = new System.Drawing.Point(101, 47);
            this.labelFormatChoice.Name = "labelFormatChoice";
            this.labelFormatChoice.Size = new System.Drawing.Size(80, 13);
            this.labelFormatChoice.TabIndex = 7;
            this.labelFormatChoice.Text = "<formatchocie>";
            // 
            // FormExportSelectionAsFormat
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(484, 362);
            this.Controls.Add(this.labelFormatChoice);
            this.Controls.Add(this.labelFormat);
            this.Controls.Add(this.filenamePicker1);
            this.Controls.Add(this.buttonOK);
            this.Controls.Add(this.buttonCancel);
            this.Controls.Add(this.labelOutputFile);
            this.Controls.Add(this.labelPageName);
            this.Controls.Add(this.labelDocumentName);
            this.Controls.Add(this.labelPage);
            this.Controls.Add(this.labelDoc);
            this.MinimumSize = new System.Drawing.Size(500, 400);
            this.Name = "FormExportSelectionAsFormat";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Text = "Export Selection";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Label labelDoc;
        private Label labelPage;
        private Label labelDocumentName;
        private Label labelPageName;
        private Label labelOutputFile;
        private Button buttonCancel;
        private Button buttonOK;
        private VisioAutomation.UI.FilenamePicker filenamePicker1;
        private Label labelFormat;
        private Label labelFormatChoice;
    }
}