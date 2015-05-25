namespace VisioAutomation.UI
{
    partial class ColorSelectorMedium
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
            if (disposing && (this.components != null))
            {
                this.components.Dispose();
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
            this.smallColorPicker1 = new ColorSelectorSmall();
            this.colorSelectorSmall1 = new ColorSelectorSmall();
            this.SuspendLayout();
            // 
            // smallColorPicker1
            // 
            this.smallColorPicker1.Color = System.Drawing.SystemColors.Control;
            this.smallColorPicker1.Location = new System.Drawing.Point(0, 0);
            this.smallColorPicker1.Name = "smallColorPicker1";
            this.smallColorPicker1.Size = new System.Drawing.Size(48, 22);
            this.smallColorPicker1.TabIndex = 3;
            // 
            // colorSelectorSmall1
            // 
            this.colorSelectorSmall1.Color = System.Drawing.SystemColors.Control;
            this.colorSelectorSmall1.Location = new System.Drawing.Point(0, 0);
            this.colorSelectorSmall1.Name = "colorSelectorSmall1";
            this.colorSelectorSmall1.Size = new System.Drawing.Size(42, 22);
            this.colorSelectorSmall1.TabIndex = 4;
            // 
            // ColorSelectorMedium
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.colorSelectorSmall1);
            this.Controls.Add(this.smallColorPicker1);
            this.Name = "ColorSelectorMedium";
            this.Size = new System.Drawing.Size(169, 22);
            this.ResumeLayout(false);

        }

        #endregion

        private ColorSelectorSmall smallColorPicker1;
        private ColorSelectorSmall colorSelectorSmall1;
    }
}
