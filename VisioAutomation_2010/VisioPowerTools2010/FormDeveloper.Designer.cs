namespace VisioPowerTools2010
{
    partial class FormDeveloper
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
            this.buttonHierarchy = new System.Windows.Forms.Button();
            this.buttonDiagramWithClasses = new System.Windows.Forms.Button();
            this.labelDrawDiagrams = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // buttonHierarchy
            // 
            this.buttonHierarchy.Location = new System.Drawing.Point(12, 30);
            this.buttonHierarchy.Name = "buttonHierarchy";
            this.buttonHierarchy.Size = new System.Drawing.Size(155, 23);
            this.buttonHierarchy.TabIndex = 0;
            this.buttonHierarchy.Text = "Namespaces";
            this.buttonHierarchy.UseVisualStyleBackColor = true;
            this.buttonHierarchy.Click += new System.EventHandler(this.buttonHierarchy_Click);
            // 
            // buttonDiagramWithClasses
            // 
            this.buttonDiagramWithClasses.Location = new System.Drawing.Point(12, 59);
            this.buttonDiagramWithClasses.Name = "buttonDiagramWithClasses";
            this.buttonDiagramWithClasses.Size = new System.Drawing.Size(155, 23);
            this.buttonDiagramWithClasses.TabIndex = 1;
            this.buttonDiagramWithClasses.Text = "Namespaces and Types";
            this.buttonDiagramWithClasses.UseVisualStyleBackColor = true;
            this.buttonDiagramWithClasses.Click += new System.EventHandler(this.buttonDiagramWithClasses_Click);
            // 
            // labelDrawDiagrams
            // 
            this.labelDrawDiagrams.AutoSize = true;
            this.labelDrawDiagrams.Location = new System.Drawing.Point(12, 11);
            this.labelDrawDiagrams.Name = "labelDrawDiagrams";
            this.labelDrawDiagrams.Size = new System.Drawing.Size(96, 13);
            this.labelDrawDiagrams.TabIndex = 2;
            this.labelDrawDiagrams.Text = "Generate diagrams";
            // 
            // FormDeveloper
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(235, 144);
            this.Controls.Add(this.labelDrawDiagrams);
            this.Controls.Add(this.buttonDiagramWithClasses);
            this.Controls.Add(this.buttonHierarchy);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "FormDeveloper";
            this.Text = "Developer Tools";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonHierarchy;
        private System.Windows.Forms.Button buttonDiagramWithClasses;
        private System.Windows.Forms.Label labelDrawDiagrams;
    }
}