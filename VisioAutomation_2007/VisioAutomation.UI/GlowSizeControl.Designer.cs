namespace VisioAutomation.UI
{
    partial class GlowSizeControl
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
            this.numericUpDown1 = new System.Windows.Forms.NumericUpDown();
            this.sliderGlowSize = new VisioAutomation.UI.CommonControls.Slider();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown1)).BeginInit();
            this.SuspendLayout();
            // 
            // numericUpDown1
            // 
            this.numericUpDown1.Location = new System.Drawing.Point(159, 1);
            this.numericUpDown1.Maximum = new decimal(new int[] {
            200,
            0,
            0,
            0});
            this.numericUpDown1.Name = "numericUpDown1";
            this.numericUpDown1.Size = new System.Drawing.Size(41, 20);
            this.numericUpDown1.TabIndex = 3;
            this.numericUpDown1.Value = new decimal(new int[] {
            150,
            0,
            0,
            0});
            this.numericUpDown1.ValueChanged += new System.EventHandler(this.numericUpDown1_ValueChanged);
            // 
            // sliderGlowSize
            // 
            this.sliderGlowSize.Location = new System.Drawing.Point(1, 1);
            this.sliderGlowSize.Max = 200F;
            this.sliderGlowSize.Min = 0F;
            this.sliderGlowSize.Name = "sliderGlowSize";
            this.sliderGlowSize.Size = new System.Drawing.Size(150, 22);
            this.sliderGlowSize.TabIndex = 4;
            this.sliderGlowSize.Value = 0F;
            this.sliderGlowSize.ValueChanged += new VisioAutomation.UI.CommonControls.Slider.ValueChangedEventHandler(this.ucSliderGlowSize_ValueChanged);
            // 
            // GlowSizeControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.sliderGlowSize);
            this.Controls.Add(this.numericUpDown1);
            this.Name = "GlowSizeControl";
            this.Size = new System.Drawing.Size(200, 22);
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.NumericUpDown numericUpDown1;
        private VisioAutomation.UI.CommonControls.Slider sliderGlowSize;

    }
}
