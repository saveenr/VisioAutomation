namespace VisioAutomationDevTool
{
    partial class FormVADevTool
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
            this.buttonLaunchVisio = new System.Windows.Forms.Button();
            this.buttonGetUnitTestCode = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // buttonLaunchVisio
            // 
            this.buttonLaunchVisio.Location = new System.Drawing.Point(13, 13);
            this.buttonLaunchVisio.Name = "buttonLaunchVisio";
            this.buttonLaunchVisio.Size = new System.Drawing.Size(139, 23);
            this.buttonLaunchVisio.TabIndex = 0;
            this.buttonLaunchVisio.Text = "Launch Visio";
            this.buttonLaunchVisio.UseVisualStyleBackColor = true;
            this.buttonLaunchVisio.Click += new System.EventHandler(this.buttonLaunchVisio_Click);
            // 
            // buttonGetUnitTestCode
            // 
            this.buttonGetUnitTestCode.Location = new System.Drawing.Point(13, 63);
            this.buttonGetUnitTestCode.Name = "buttonGetUnitTestCode";
            this.buttonGetUnitTestCode.Size = new System.Drawing.Size(139, 23);
            this.buttonGetUnitTestCode.TabIndex = 1;
            this.buttonGetUnitTestCode.Text = "Get Unit Test Code";
            this.buttonGetUnitTestCode.UseVisualStyleBackColor = true;
            this.buttonGetUnitTestCode.Click += new System.EventHandler(this.buttonGetUnitTestCode_Click);
            // 
            // FormVADevTool
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(526, 278);
            this.Controls.Add(this.buttonGetUnitTestCode);
            this.Controls.Add(this.buttonLaunchVisio);
            this.Name = "FormVADevTool";
            this.Text = "VisioAutomation Dev Tool";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button buttonLaunchVisio;
        private System.Windows.Forms.Button buttonGetUnitTestCode;
    }
}

