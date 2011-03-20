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
            this.tabPage3PointGradient = new System.Windows.Forms.TabPage();
            this.buttonSet3PointFill = new System.Windows.Forms.Button();
            this.tabPage2ColorGlow = new System.Windows.Forms.TabPage();
            this.fillGradient1 = new VA.UI.FillControl();
            this.uC3PointFill1 = new VA.UI.ThreePointFillControl();
            this.uC2ColorGlow1 = new VA.UI.TwoColorGlowControl();
            this.tabControl1.SuspendLayout();
            this.tabPageGradient.SuspendLayout();
            this.tabPage3PointGradient.SuspendLayout();
            this.tabPage2ColorGlow.SuspendLayout();
            this.SuspendLayout();
            // 
            // buttonSet2ColorGlow
            // 
            this.buttonSet2ColorGlow.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonSet2ColorGlow.Location = new System.Drawing.Point(281, 260);
            this.buttonSet2ColorGlow.Name = "buttonSet2ColorGlow";
            this.buttonSet2ColorGlow.Size = new System.Drawing.Size(70, 23);
            this.buttonSet2ColorGlow.TabIndex = 4;
            this.buttonSet2ColorGlow.Text = "Apply";
            this.buttonSet2ColorGlow.UseVisualStyleBackColor = true;
            this.buttonSet2ColorGlow.Click += new System.EventHandler(this.buttonSet2ColorGlow_Click);
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPageGradient);
            this.tabControl1.Controls.Add(this.tabPage3PointGradient);
            this.tabControl1.Controls.Add(this.tabPage2ColorGlow);
            this.tabControl1.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
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
            this.tabPageGradient.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
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
            this.buttonUpdateFill.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
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
            this.buttonSetFillGradient.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonSetFillGradient.Location = new System.Drawing.Point(281, 259);
            this.buttonSetFillGradient.Name = "buttonSetFillGradient";
            this.buttonSetFillGradient.Size = new System.Drawing.Size(70, 23);
            this.buttonSetFillGradient.TabIndex = 1;
            this.buttonSetFillGradient.Text = "Apply";
            this.buttonSetFillGradient.UseVisualStyleBackColor = true;
            this.buttonSetFillGradient.Click += new System.EventHandler(this.buttonSetFillGradient_Click);
            // 
            // tabPage3PointGradient
            // 
            this.tabPage3PointGradient.BackColor = System.Drawing.Color.Transparent;
            this.tabPage3PointGradient.Controls.Add(this.buttonSet3PointFill);
            this.tabPage3PointGradient.Controls.Add(this.uC3PointFill1);
            this.tabPage3PointGradient.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabPage3PointGradient.Location = new System.Drawing.Point(4, 22);
            this.tabPage3PointGradient.Name = "tabPage3PointGradient";
            this.tabPage3PointGradient.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage3PointGradient.Size = new System.Drawing.Size(357, 289);
            this.tabPage3PointGradient.TabIndex = 1;
            this.tabPage3PointGradient.Text = "3 Color Fill";
            this.tabPage3PointGradient.UseVisualStyleBackColor = true;
            // 
            // buttonSet3PointFill
            // 
            this.buttonSet3PointFill.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonSet3PointFill.Location = new System.Drawing.Point(281, 260);
            this.buttonSet3PointFill.Name = "buttonSet3PointFill";
            this.buttonSet3PointFill.Size = new System.Drawing.Size(70, 23);
            this.buttonSet3PointFill.TabIndex = 1;
            this.buttonSet3PointFill.Text = "Apply";
            this.buttonSet3PointFill.UseVisualStyleBackColor = true;
            this.buttonSet3PointFill.Click += new System.EventHandler(this.buttonSet3PointFill_Click);
            // 
            // tabPage2ColorGlow
            // 
            this.tabPage2ColorGlow.BackColor = System.Drawing.Color.Transparent;
            this.tabPage2ColorGlow.Controls.Add(this.uC2ColorGlow1);
            this.tabPage2ColorGlow.Controls.Add(this.buttonSet2ColorGlow);
            this.tabPage2ColorGlow.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabPage2ColorGlow.Location = new System.Drawing.Point(4, 22);
            this.tabPage2ColorGlow.Name = "tabPage2ColorGlow";
            this.tabPage2ColorGlow.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2ColorGlow.Size = new System.Drawing.Size(357, 289);
            this.tabPage2ColorGlow.TabIndex = 0;
            this.tabPage2ColorGlow.Text = "2 Color Glow";
            this.tabPage2ColorGlow.UseVisualStyleBackColor = true;
            // 
            // fillGradient1
            // 
            this.fillGradient1.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.fillGradient1.Location = new System.Drawing.Point(7, 7);
            this.fillGradient1.Name = "fillGradient1";
            this.fillGradient1.Size = new System.Drawing.Size(344, 275);
            this.fillGradient1.TabIndex = 0;
            // 
            // uC3PointFill1
            // 
            this.uC3PointFill1.Corner1Color = System.Drawing.Color.Red;
            this.uC3PointFill1.Corner2Color = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.uC3PointFill1.Direction = VA.Drawing.DirectionRelative.Up;
            this.uC3PointFill1.EdgeColor = System.Drawing.SystemColors.ActiveCaption;
            this.uC3PointFill1.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.uC3PointFill1.Location = new System.Drawing.Point(3, 6);
            this.uC3PointFill1.Name = "uC3PointFill1";
            this.uC3PointFill1.Size = new System.Drawing.Size(281, 248);
            this.uC3PointFill1.TabIndex = 0;
            // 
            // uC2ColorGlow1
            // 
            this.uC2ColorGlow1.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.uC2ColorGlow1.GlowSize = 170;
            this.uC2ColorGlow1.Location = new System.Drawing.Point(6, 6);
            this.uC2ColorGlow1.LowerColor = System.Drawing.Color.DeepPink;
            this.uC2ColorGlow1.LowerTransparency = 0;
            this.uC2ColorGlow1.Name = "uC2ColorGlow1";
            this.uC2ColorGlow1.Size = new System.Drawing.Size(288, 168);
            this.uC2ColorGlow1.TabIndex = 5;
            this.uC2ColorGlow1.UpperColor = System.Drawing.Color.Red;
            this.uC2ColorGlow1.UpperTransparency = 0;
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
            this.tabPage3PointGradient.ResumeLayout(false);
            this.tabPage2ColorGlow.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button buttonSet2ColorGlow;
        private VA.UI.TwoColorGlowControl uC2ColorGlow1;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage2ColorGlow;
        private System.Windows.Forms.TabPage tabPage3PointGradient;
        private System.Windows.Forms.Button buttonSet3PointFill;
        private VA.UI.ThreePointFillControl uC3PointFill1;
        private System.Windows.Forms.TabPage tabPageGradient;
        private System.Windows.Forms.Button buttonSetFillGradient;
        private VA.UI.FillControl fillGradient1;
        private System.Windows.Forms.Button buttonUpdateFill;
    }
}

