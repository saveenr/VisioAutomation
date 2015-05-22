namespace VisioPowerTools
{
    partial class FormImportFlowChartXML : System.Windows.Forms.Form
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
            this.buttonCancel = new System.Windows.Forms.Button();
            this.buttonImport = new System.Windows.Forms.Button();
            this.textBoxOutput = new System.Windows.Forms.TextBox();
            this.labelInputFilename = new System.Windows.Forms.Label();
            this.labelMessageLog = new System.Windows.Forms.Label();
            this.filenamePicker1 = new VisioAutomation.UI.CommonControls.FilenamePicker();
            this.SuspendLayout();
            // 
            // buttonCancel
            // 
            this.buttonCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonCancel.Location = new System.Drawing.Point(396, 327);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(75, 23);
            this.buttonCancel.TabIndex = 1;
            this.buttonCancel.Text = "Cancel";
            this.buttonCancel.UseVisualStyleBackColor = true;
            this.buttonCancel.Click += new System.EventHandler(this.buttonCancel_Click);
            // 
            // buttonImport
            // 
            this.buttonImport.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonImport.Location = new System.Drawing.Point(315, 327);
            this.buttonImport.Name = "buttonImport";
            this.buttonImport.Size = new System.Drawing.Size(75, 23);
            this.buttonImport.TabIndex = 2;
            this.buttonImport.Text = "Import";
            this.buttonImport.UseVisualStyleBackColor = true;
            this.buttonImport.Click += new System.EventHandler(this.buttonImport_Click);
            // 
            // textBoxOutput
            // 
            this.textBoxOutput.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxOutput.Location = new System.Drawing.Point(13, 134);
            this.textBoxOutput.Multiline = true;
            this.textBoxOutput.Name = "textBoxOutput";
            this.textBoxOutput.Size = new System.Drawing.Size(458, 187);
            this.textBoxOutput.TabIndex = 3;
            // 
            // labelInputFilename
            // 
            this.labelInputFilename.AutoSize = true;
            this.labelInputFilename.Location = new System.Drawing.Point(13, 25);
            this.labelInputFilename.Name = "labelInputFilename";
            this.labelInputFilename.Size = new System.Drawing.Size(160, 13);
            this.labelInputFilename.TabIndex = 4;
            this.labelInputFilename.Text = "Input Flowchart XML filename";
            // 
            // labelMessageLog
            // 
            this.labelMessageLog.AutoSize = true;
            this.labelMessageLog.Location = new System.Drawing.Point(16, 113);
            this.labelMessageLog.Name = "labelMessageLog";
            this.labelMessageLog.Size = new System.Drawing.Size(74, 13);
            this.labelMessageLog.TabIndex = 5;
            this.labelMessageLog.Text = "Message Log";
            // 
            // filenamePicker1
            // 
            this.filenamePicker1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.filenamePicker1.Filename = "";
            this.filenamePicker1.Location = new System.Drawing.Point(13, 42);
            this.filenamePicker1.Name = "filenamePicker1";
            this.filenamePicker1.ReadOnly = false;
            this.filenamePicker1.Size = new System.Drawing.Size(458, 47);
            this.filenamePicker1.TabIndex = 6;
            // 
            // FormImportFlowChartXML
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(484, 362);
            this.Controls.Add(this.filenamePicker1);
            this.Controls.Add(this.labelMessageLog);
            this.Controls.Add(this.labelInputFilename);
            this.Controls.Add(this.textBoxOutput);
            this.Controls.Add(this.buttonImport);
            this.Controls.Add(this.buttonCancel);
            this.MinimumSize = new System.Drawing.Size(500, 400);
            this.Name = "FormImportFlowChartXML";
            this.Text = "Import FlowChart XML";
            this.Load += new System.EventHandler(this.FormImportFlowChartXML_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonCancel;
        private System.Windows.Forms.Button buttonImport;
        private System.Windows.Forms.TextBox textBoxOutput;
        private System.Windows.Forms.Label labelInputFilename;
        private System.Windows.Forms.Label labelMessageLog;
        private VisioAutomation.UI.CommonControls.FilenamePicker filenamePicker1;
    }
}