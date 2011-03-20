namespace VisioPowerTools
{
    partial class FormViewTool : System.Windows.Forms.Form
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
            this.buttonPreviousPage = new System.Windows.Forms.Button();
            this.buttonNextPage = new System.Windows.Forms.Button();
            this.buttonZoomToSelection = new System.Windows.Forms.Button();
            this.buttonZoomToPage = new System.Windows.Forms.Button();
            this.buttonZoomOut = new System.Windows.Forms.Button();
            this.buttonZoomIn = new System.Windows.Forms.Button();
            this.labelZoomTo = new System.Windows.Forms.Label();
            this.labelPage = new System.Windows.Forms.Label();
            this.buttonPageLast = new System.Windows.Forms.Button();
            this.buttonFirstPage = new System.Windows.Forms.Button();
            this.labelZoomLevel = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // buttonPreviousPage
            // 
            this.buttonPreviousPage.FlatAppearance.BorderColor = System.Drawing.SystemColors.ActiveBorder;
            this.buttonPreviousPage.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonPreviousPage.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonPreviousPage.Location = new System.Drawing.Point(45, 126);
            this.buttonPreviousPage.Name = "buttonPreviousPage";
            this.buttonPreviousPage.Size = new System.Drawing.Size(30, 23);
            this.buttonPreviousPage.TabIndex = 0;
            this.buttonPreviousPage.Text = "<";
            this.buttonPreviousPage.UseVisualStyleBackColor = true;
            this.buttonPreviousPage.Click += new System.EventHandler(this.buttonPreviousPage_Click);
            // 
            // buttonNextPage
            // 
            this.buttonNextPage.FlatAppearance.BorderColor = System.Drawing.SystemColors.ActiveBorder;
            this.buttonNextPage.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonNextPage.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonNextPage.Location = new System.Drawing.Point(81, 126);
            this.buttonNextPage.Name = "buttonNextPage";
            this.buttonNextPage.Size = new System.Drawing.Size(30, 23);
            this.buttonNextPage.TabIndex = 1;
            this.buttonNextPage.Text = ">";
            this.buttonNextPage.UseVisualStyleBackColor = true;
            this.buttonNextPage.Click += new System.EventHandler(this.buttonNextPage_Click);
            // 
            // buttonZoomToSelection
            // 
            this.buttonZoomToSelection.FlatAppearance.BorderColor = System.Drawing.SystemColors.ActiveBorder;
            this.buttonZoomToSelection.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonZoomToSelection.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonZoomToSelection.Location = new System.Drawing.Point(64, 22);
            this.buttonZoomToSelection.Name = "buttonZoomToSelection";
            this.buttonZoomToSelection.Size = new System.Drawing.Size(70, 23);
            this.buttonZoomToSelection.TabIndex = 3;
            this.buttonZoomToSelection.Text = "Selection";
            this.buttonZoomToSelection.UseVisualStyleBackColor = true;
            this.buttonZoomToSelection.Click += new System.EventHandler(this.buttonZoomToSelection_Click);
            // 
            // buttonZoomToPage
            // 
            this.buttonZoomToPage.FlatAppearance.BorderColor = System.Drawing.SystemColors.ActiveBorder;
            this.buttonZoomToPage.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonZoomToPage.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonZoomToPage.Location = new System.Drawing.Point(8, 22);
            this.buttonZoomToPage.Name = "buttonZoomToPage";
            this.buttonZoomToPage.Size = new System.Drawing.Size(50, 23);
            this.buttonZoomToPage.TabIndex = 2;
            this.buttonZoomToPage.Text = "Page";
            this.buttonZoomToPage.UseVisualStyleBackColor = true;
            this.buttonZoomToPage.Click += new System.EventHandler(this.buttonZoomToPage_Click);
            // 
            // buttonZoomOut
            // 
            this.buttonZoomOut.FlatAppearance.BorderColor = System.Drawing.SystemColors.ActiveBorder;
            this.buttonZoomOut.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonZoomOut.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonZoomOut.Location = new System.Drawing.Point(44, 74);
            this.buttonZoomOut.Name = "buttonZoomOut";
            this.buttonZoomOut.Size = new System.Drawing.Size(30, 23);
            this.buttonZoomOut.TabIndex = 5;
            this.buttonZoomOut.Text = "-";
            this.buttonZoomOut.UseVisualStyleBackColor = true;
            this.buttonZoomOut.Click += new System.EventHandler(this.buttonZoomOut_Click);
            // 
            // buttonZoomIn
            // 
            this.buttonZoomIn.FlatAppearance.BorderColor = System.Drawing.SystemColors.ActiveBorder;
            this.buttonZoomIn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonZoomIn.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonZoomIn.Location = new System.Drawing.Point(8, 74);
            this.buttonZoomIn.Name = "buttonZoomIn";
            this.buttonZoomIn.Size = new System.Drawing.Size(30, 23);
            this.buttonZoomIn.TabIndex = 4;
            this.buttonZoomIn.Text = "+";
            this.buttonZoomIn.UseVisualStyleBackColor = true;
            this.buttonZoomIn.Click += new System.EventHandler(this.buttonZoomIn_Click);
            // 
            // labelZoomTo
            // 
            this.labelZoomTo.AutoSize = true;
            this.labelZoomTo.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelZoomTo.Location = new System.Drawing.Point(6, 4);
            this.labelZoomTo.Name = "labelZoomTo";
            this.labelZoomTo.Size = new System.Drawing.Size(58, 13);
            this.labelZoomTo.TabIndex = 6;
            this.labelZoomTo.Text = "ZOOM TO";
            // 
            // labelPage
            // 
            this.labelPage.AutoSize = true;
            this.labelPage.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelPage.Location = new System.Drawing.Point(7, 107);
            this.labelPage.Name = "labelPage";
            this.labelPage.Size = new System.Drawing.Size(34, 13);
            this.labelPage.TabIndex = 7;
            this.labelPage.Text = "PAGE";
            // 
            // buttonPageLast
            // 
            this.buttonPageLast.FlatAppearance.BorderColor = System.Drawing.SystemColors.ActiveBorder;
            this.buttonPageLast.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonPageLast.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonPageLast.Location = new System.Drawing.Point(117, 126);
            this.buttonPageLast.Name = "buttonPageLast";
            this.buttonPageLast.Size = new System.Drawing.Size(30, 23);
            this.buttonPageLast.TabIndex = 9;
            this.buttonPageLast.Text = ">|";
            this.buttonPageLast.UseVisualStyleBackColor = true;
            this.buttonPageLast.Click += new System.EventHandler(this.buttonPageLast_Click);
            // 
            // buttonFirstPage
            // 
            this.buttonFirstPage.FlatAppearance.BorderColor = System.Drawing.SystemColors.ActiveBorder;
            this.buttonFirstPage.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonFirstPage.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonFirstPage.Location = new System.Drawing.Point(9, 126);
            this.buttonFirstPage.Name = "buttonFirstPage";
            this.buttonFirstPage.Size = new System.Drawing.Size(30, 23);
            this.buttonFirstPage.TabIndex = 8;
            this.buttonFirstPage.Text = "|<";
            this.buttonFirstPage.UseVisualStyleBackColor = true;
            this.buttonFirstPage.Click += new System.EventHandler(this.buttonFirstPage_Click);
            // 
            // labelZoomLevel
            // 
            this.labelZoomLevel.AutoSize = true;
            this.labelZoomLevel.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelZoomLevel.Location = new System.Drawing.Point(6, 56);
            this.labelZoomLevel.Name = "labelZoomLevel";
            this.labelZoomLevel.Size = new System.Drawing.Size(73, 13);
            this.labelZoomLevel.TabIndex = 10;
            this.labelZoomLevel.Text = "ZOOM LEVEL";
            // 
            // FormViewTool
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(155, 161);
            this.Controls.Add(this.labelZoomLevel);
            this.Controls.Add(this.buttonPageLast);
            this.Controls.Add(this.buttonFirstPage);
            this.Controls.Add(this.labelPage);
            this.Controls.Add(this.labelZoomTo);
            this.Controls.Add(this.buttonZoomOut);
            this.Controls.Add(this.buttonZoomIn);
            this.Controls.Add(this.buttonZoomToSelection);
            this.Controls.Add(this.buttonZoomToPage);
            this.Controls.Add(this.buttonNextPage);
            this.Controls.Add(this.buttonPreviousPage);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "FormViewTool";
            this.Text = "View";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonPreviousPage;
        private System.Windows.Forms.Button buttonNextPage;
        private System.Windows.Forms.Button buttonZoomToSelection;
        private System.Windows.Forms.Button buttonZoomToPage;
        private System.Windows.Forms.Button buttonZoomOut;
        private System.Windows.Forms.Button buttonZoomIn;
        private System.Windows.Forms.Label labelZoomTo;
        private System.Windows.Forms.Label labelPage;
        private System.Windows.Forms.Button buttonPageLast;
        private System.Windows.Forms.Button buttonFirstPage;
        private System.Windows.Forms.Label labelZoomLevel;
    }
}