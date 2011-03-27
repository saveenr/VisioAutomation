namespace VisioPowerTools
{
    partial class FormLayoutTool : System.Windows.Forms.Form
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
            this.buttonResizePageToFitContents = new System.Windows.Forms.Button();
            this.labelPage = new System.Windows.Forms.Label();
            this.buttonDuplicatePage = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // buttonResizePageToFitContents
            // 
            this.buttonResizePageToFitContents.FlatAppearance.BorderColor = System.Drawing.SystemColors.ActiveBorder;
            this.buttonResizePageToFitContents.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonResizePageToFitContents.Location = new System.Drawing.Point(12, 28);
            this.buttonResizePageToFitContents.Name = "buttonResizePageToFitContents";
            this.buttonResizePageToFitContents.Size = new System.Drawing.Size(122, 23);
            this.buttonResizePageToFitContents.TabIndex = 0;
            this.buttonResizePageToFitContents.Text = "Fit to Contents";
            this.buttonResizePageToFitContents.UseVisualStyleBackColor = true;
            this.buttonResizePageToFitContents.Click += new System.EventHandler(this.buttonResizePageToFitContents_Click);
            // 
            // labelPage
            // 
            this.labelPage.AutoSize = true;
            this.labelPage.Location = new System.Drawing.Point(9, 10);
            this.labelPage.Name = "labelPage";
            this.labelPage.Size = new System.Drawing.Size(32, 13);
            this.labelPage.TabIndex = 1;
            this.labelPage.Text = "Page";
            // 
            // buttonDuplicatePage
            // 
            this.buttonDuplicatePage.FlatAppearance.BorderColor = System.Drawing.SystemColors.ActiveBorder;
            this.buttonDuplicatePage.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonDuplicatePage.Location = new System.Drawing.Point(12, 57);
            this.buttonDuplicatePage.Name = "buttonDuplicatePage";
            this.buttonDuplicatePage.Size = new System.Drawing.Size(122, 23);
            this.buttonDuplicatePage.TabIndex = 2;
            this.buttonDuplicatePage.Text = "Duplicate";
            this.buttonDuplicatePage.UseVisualStyleBackColor = true;
            this.buttonDuplicatePage.Click += new System.EventHandler(this.buttonDuplicatePage_Click);
            // 
            // FormLayoutTool
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(149, 95);
            this.Controls.Add(this.buttonDuplicatePage);
            this.Controls.Add(this.labelPage);
            this.Controls.Add(this.buttonResizePageToFitContents);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "FormLayoutTool";
            this.Text = "Layout";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonResizePageToFitContents;
        private System.Windows.Forms.Label labelPage;
        private System.Windows.Forms.Button buttonDuplicatePage;
    }
}