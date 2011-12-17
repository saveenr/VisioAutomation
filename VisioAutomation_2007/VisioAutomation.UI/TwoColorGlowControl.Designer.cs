namespace VisioAutomation.UI
{
    partial class TwoColorGlowControl
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.labelUpperGlow = new System.Windows.Forms.Label();
            this.labelLowerGlow = new System.Windows.Forms.Label();
            this.labelLowerGlowSize = new System.Windows.Forms.Label();
            this.colorPickerLowerGlow = new VisioAutomation.UI.CommonControls.ColorSelectorSmall();
            this.colorPickerUpperGlow = new VisioAutomation.UI.CommonControls.ColorSelectorSmall();
            this.glowSize1 = new VisioAutomation.UI.GlowSizeControl();
            this.transparency2 = new VisioAutomation.UI.TransparencyControl();
            this.transparency1 = new VisioAutomation.UI.TransparencyControl();
            this.SuspendLayout();
            // 
            // labelUpperGlow
            // 
            this.labelUpperGlow.AutoSize = true;
            this.labelUpperGlow.Location = new System.Drawing.Point(10, 12);
            this.labelUpperGlow.Name = "labelUpperGlow";
            this.labelUpperGlow.Size = new System.Drawing.Size(63, 13);
            this.labelUpperGlow.TabIndex = 16;
            this.labelUpperGlow.Text = "Upper Color";
            // 
            // labelLowerGlow
            // 
            this.labelLowerGlow.AutoSize = true;
            this.labelLowerGlow.Location = new System.Drawing.Point(8, 63);
            this.labelLowerGlow.Name = "labelLowerGlow";
            this.labelLowerGlow.Size = new System.Drawing.Size(63, 13);
            this.labelLowerGlow.TabIndex = 15;
            this.labelLowerGlow.Text = "Lower Color";
            // 
            // labelLowerGlowSize
            // 
            this.labelLowerGlowSize.AutoSize = true;
            this.labelLowerGlowSize.Location = new System.Drawing.Point(10, 114);
            this.labelLowerGlowSize.Name = "labelLowerGlowSize";
            this.labelLowerGlowSize.Size = new System.Drawing.Size(103, 13);
            this.labelLowerGlowSize.TabIndex = 18;
            this.labelLowerGlowSize.Text = "Lower Glow Size (%)";
            // 
            // colorPickerLowerGlow
            // 
            this.colorPickerLowerGlow.Color = System.Drawing.Color.Fuchsia;
            this.colorPickerLowerGlow.Location = new System.Drawing.Point(10, 84);
            this.colorPickerLowerGlow.Name = "colorPickerLowerGlow";
            this.colorPickerLowerGlow.Size = new System.Drawing.Size(42, 22);
            this.colorPickerLowerGlow.TabIndex = 21;
            // 
            // colorPickerUpperGlow
            // 
            this.colorPickerUpperGlow.Color = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.colorPickerUpperGlow.Location = new System.Drawing.Point(10, 33);
            this.colorPickerUpperGlow.Name = "colorPickerUpperGlow";
            this.colorPickerUpperGlow.Size = new System.Drawing.Size(42, 22);
            this.colorPickerUpperGlow.TabIndex = 20;
            // 
            // glowSize1
            // 
            this.glowSize1.GlowSize = 150;
            this.glowSize1.Location = new System.Drawing.Point(10, 135);
            this.glowSize1.Name = "glowSize1";
            this.glowSize1.Size = new System.Drawing.Size(204, 22);
            this.glowSize1.TabIndex = 19;
            // 
            // transparency2
            // 
            this.transparency2.BackColor = System.Drawing.Color.Transparent;
            this.transparency2.Location = new System.Drawing.Point(58, 84);
            this.transparency2.Name = "transparency2";
            this.transparency2.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.transparency2.Size = new System.Drawing.Size(213, 22);
            this.transparency2.TabIndex = 14;
            this.transparency2.TransparencyPercent = 0;
            // 
            // transparency1
            // 
            this.transparency1.BackColor = System.Drawing.Color.Transparent;
            this.transparency1.Location = new System.Drawing.Point(58, 33);
            this.transparency1.Name = "transparency1";
            this.transparency1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.transparency1.Size = new System.Drawing.Size(213, 22);
            this.transparency1.TabIndex = 13;
            this.transparency1.TransparencyPercent = 0;
            // 
            // TwoColorGlowControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.colorPickerLowerGlow);
            this.Controls.Add(this.colorPickerUpperGlow);
            this.Controls.Add(this.glowSize1);
            this.Controls.Add(this.labelLowerGlowSize);
            this.Controls.Add(this.labelUpperGlow);
            this.Controls.Add(this.labelLowerGlow);
            this.Controls.Add(this.transparency2);
            this.Controls.Add(this.transparency1);
            this.Name = "TwoColorGlowControl";
            this.Size = new System.Drawing.Size(261, 172);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label labelUpperGlow;
        private System.Windows.Forms.Label labelLowerGlow;
        private TransparencyControl transparency2;
        private TransparencyControl transparency1;
        private System.Windows.Forms.Label labelLowerGlowSize;
        private GlowSizeControl glowSize1;
        private VisioAutomation.UI.CommonControls.ColorSelectorSmall colorPickerUpperGlow;
        private VisioAutomation.UI.CommonControls.ColorSelectorSmall colorPickerLowerGlow;
    }
}
