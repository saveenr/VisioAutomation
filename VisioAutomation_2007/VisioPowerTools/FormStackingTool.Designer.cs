namespace VisioPowerTools
{
    partial class FormStackingTool
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
            this.label2 = new System.Windows.Forms.Label();
            this.comboBoxSnapDelta = new System.Windows.Forms.ComboBox();
            this.buttonLayoutE2EH = new System.Windows.Forms.Button();
            this.buttonLayoutE2eV = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(7, 63);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(130, 13);
            this.label2.TabIndex = 16;
            this.label2.Text = "Distance between shapes";
            // 
            // comboBoxSnapDelta
            // 
            this.comboBoxSnapDelta.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.comboBoxSnapDelta.FormattingEnabled = true;
            this.comboBoxSnapDelta.Items.AddRange(new object[] {
            "0.25",
            "0.5",
            "1.0",
            "2.0",
            "3.0",
            "4.0",
            ""});
            this.comboBoxSnapDelta.Location = new System.Drawing.Point(12, 79);
            this.comboBoxSnapDelta.Name = "comboBoxSnapDelta";
            this.comboBoxSnapDelta.Size = new System.Drawing.Size(110, 21);
            this.comboBoxSnapDelta.TabIndex = 17;
            this.comboBoxSnapDelta.Text = "0.25";
            this.comboBoxSnapDelta.ValueMemberChanged += new System.EventHandler(this.comboBoxSnapDelta_ValueMemberChanged);
            this.comboBoxSnapDelta.TextChanged += new System.EventHandler(this.comboBoxSnapDelta_TextChanged);
            // 
            // buttonLayoutE2EH
            // 
            this.buttonLayoutE2EH.FlatAppearance.BorderColor = System.Drawing.SystemColors.ButtonShadow;
            this.buttonLayoutE2EH.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonLayoutE2EH.Location = new System.Drawing.Point(12, 28);
            this.buttonLayoutE2EH.Name = "buttonLayoutE2EH";
            this.buttonLayoutE2EH.Size = new System.Drawing.Size(60, 23);
            this.buttonLayoutE2EH.TabIndex = 27;
            this.buttonLayoutE2EH.Text = "Row";
            this.buttonLayoutE2EH.UseVisualStyleBackColor = true;
            this.buttonLayoutE2EH.Click += new System.EventHandler(this.buttonLayoutE2EH_Click);
            // 
            // buttonLayoutE2eV
            // 
            this.buttonLayoutE2eV.FlatAppearance.BorderColor = System.Drawing.SystemColors.ButtonShadow;
            this.buttonLayoutE2eV.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonLayoutE2eV.Location = new System.Drawing.Point(78, 28);
            this.buttonLayoutE2eV.Name = "buttonLayoutE2eV";
            this.buttonLayoutE2eV.Size = new System.Drawing.Size(60, 23);
            this.buttonLayoutE2eV.TabIndex = 28;
            this.buttonLayoutE2eV.Text = "Column";
            this.buttonLayoutE2eV.UseVisualStyleBackColor = true;
            this.buttonLayoutE2eV.Click += new System.EventHandler(this.buttonLayoutE2eV_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(7, 9);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(86, 13);
            this.label4.TabIndex = 29;
            this.label4.Text = "Stack shapes as";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(128, 82);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(38, 13);
            this.label1.TabIndex = 31;
            this.label1.Text = "inches";
            // 
            // FormStackingTool
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(210, 116);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.buttonLayoutE2eV);
            this.Controls.Add(this.buttonLayoutE2EH);
            this.Controls.Add(this.comboBoxSnapDelta);
            this.Controls.Add(this.label2);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "FormStackingTool";
            this.Text = "Arrange";
            this.Load += new System.EventHandler(this.FormArrangeTool_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox comboBoxSnapDelta;
        private System.Windows.Forms.Button buttonLayoutE2EH;
        private System.Windows.Forms.Button buttonLayoutE2eV;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label1;
    }
}