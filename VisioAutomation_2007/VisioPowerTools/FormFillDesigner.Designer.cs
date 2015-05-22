using VA = VisioAutomation;

namespace VisioPowerTools
{
    partial class FormFillDesigner : System.Windows.Forms.Form
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
            this.buttonSet2ColorGlow = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPageGradient = new System.Windows.Forms.TabPage();
            this.buttonUpdateFill = new System.Windows.Forms.Button();
            this.buttonSetFillGradient = new System.Windows.Forms.Button();
            this.fillGradient1 = new VA.UI.FillControl();
            this.tabControl1.SuspendLayout();
            this.tabPageGradient.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPageGradient);
            this.tabControl1.Location = new System.Drawing.Point(4, 2);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(365, 315);
            this.tabControl1.TabIndex = 8;
            // 
            // tabPageGradient
            // 
            this.tabPageGradient.Controls.Add(this.buttonUpdateFill);
            this.tabPageGradient.Controls.Add(this.buttonSetFillGradient);
            this.tabPageGradient.Controls.Add(this.fillGradient1);
            this.tabPageGradient.Location = new System.Drawing.Point(4, 22);
            this.tabPageGradient.Name = "tabPageGradient";
            this.tabPageGradient.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageGradient.Size = new System.Drawing.Size(357, 289);
            this.tabPageGradient.TabIndex = 2;
            this.tabPageGradient.Text = "Fill";
            this.tabPageGradient.UseVisualStyleBackColor = true;
            // 
            // buttonUpdateFill
            // 
            this.buttonUpdateFill.Location = new System.Drawing.Point(205, 259);
            this.buttonUpdateFill.Name = "buttonUpdateFill";
            this.buttonUpdateFill.Size = new System.Drawing.Size(70, 23);
            this.buttonUpdateFill.TabIndex = 2;
            this.buttonUpdateFill.Text = "Read";
            this.buttonUpdateFill.UseVisualStyleBackColor = true;
            this.buttonUpdateFill.Click += new System.EventHandler(this.buttonUpdateFill_Click);
            // 
            // buttonSetFillGradient
            // 
            this.buttonSetFillGradient.Location = new System.Drawing.Point(281, 259);
            this.buttonSetFillGradient.Name = "buttonSetFillGradient";
            this.buttonSetFillGradient.Size = new System.Drawing.Size(70, 23);
            this.buttonSetFillGradient.TabIndex = 1;
            this.buttonSetFillGradient.Text = "Apply";
            this.buttonSetFillGradient.UseVisualStyleBackColor = true;
            this.buttonSetFillGradient.Click += new System.EventHandler(this.buttonSetFillGradient_Click);

            // 
            // fillGradient1
            // 
            this.fillGradient1.Location = new System.Drawing.Point(7, 7);
            this.fillGradient1.Name = "fillGradient1";
            this.fillGradient1.Size = new System.Drawing.Size(344, 275);
            this.fillGradient1.TabIndex = 0;
            // 
            // FormFillDesigner
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(373, 320);
            this.Controls.Add(this.tabControl1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormFillDesigner";
            this.Text = "Fill";
            this.tabControl1.ResumeLayout(false);
            this.tabPageGradient.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button buttonSet2ColorGlow;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPageGradient;
        private System.Windows.Forms.Button buttonSetFillGradient;
        private VA.UI.FillControl fillGradient1;
        private System.Windows.Forms.Button buttonUpdateFill;
    }
}

