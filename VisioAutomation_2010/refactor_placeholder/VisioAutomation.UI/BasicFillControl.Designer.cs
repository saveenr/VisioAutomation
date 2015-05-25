namespace VisioAutomation.UI
{
    partial class BasicFillControl
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
            this.comboBoxPattern = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.colorPickerBackground = new ColorSelectorSmall();
            this.colorPickerForeground = new ColorSelectorSmall();
            this.linkLabelTools = new System.Windows.Forms.LinkLabel();
            this.ucTransparency2 = new VisioAutomation.UI.TransparencyControl();
            this.ucTransparency1 = new VisioAutomation.UI.TransparencyControl();
            this.SuspendLayout();
            // 
            // comboBoxPattern
            // 
            this.comboBoxPattern.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxPattern.FormattingEnabled = true;
            this.comboBoxPattern.Location = new System.Drawing.Point(75, 59);
            this.comboBoxPattern.Name = "comboBoxPattern";
            this.comboBoxPattern.Size = new System.Drawing.Size(143, 21);
            this.comboBoxPattern.TabIndex = 4;
            this.comboBoxPattern.SelectedIndexChanged += new System.EventHandler(this.comboBoxGradient_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(8, 5);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(61, 13);
            this.label1.TabIndex = 5;
            this.label1.Text = "Foreground";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(4, 36);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(65, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "Background";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(28, 63);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(41, 13);
            this.label3.TabIndex = 7;
            this.label3.Text = "Pattern";
            // 
            // colorPickerBackground
            // 
            this.colorPickerBackground.Color = System.Drawing.Color.Black;
            this.colorPickerBackground.Location = new System.Drawing.Point(75, 31);
            this.colorPickerBackground.Name = "colorPickerBackground";
            this.colorPickerBackground.Size = new System.Drawing.Size(47, 22);
            this.colorPickerBackground.TabIndex = 2;
            // 
            // colorPickerForeground
            // 
            this.colorPickerForeground.Color = System.Drawing.Color.Red;
            this.colorPickerForeground.Location = new System.Drawing.Point(75, 0);
            this.colorPickerForeground.Name = "colorPickerForeground";
            this.colorPickerForeground.Size = new System.Drawing.Size(47, 22);
            this.colorPickerForeground.TabIndex = 0;
            // 
            // linkLabelTools
            // 
            this.linkLabelTools.AutoSize = true;
            this.linkLabelTools.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.linkLabelTools.Location = new System.Drawing.Point(289, 61);
            this.linkLabelTools.Name = "linkLabelTools";
            this.linkLabelTools.Size = new System.Drawing.Size(27, 12);
            this.linkLabelTools.TabIndex = 8;
            this.linkLabelTools.TabStop = true;
            this.linkLabelTools.Text = "Tools";
            this.linkLabelTools.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelTools_LinkClicked);
            // 
            // ucTransparency2
            // 
            this.ucTransparency2.BackColor = System.Drawing.Color.Transparent;
            this.ucTransparency2.Location = new System.Drawing.Point(128, 31);
            this.ucTransparency2.Name = "ucTransparency2";
            this.ucTransparency2.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.ucTransparency2.Size = new System.Drawing.Size(194, 22);
            this.ucTransparency2.TabIndex = 3;
            this.ucTransparency2.TransparencyPercent = 0;
            // 
            // ucTransparency1
            // 
            this.ucTransparency1.BackColor = System.Drawing.Color.Transparent;
            this.ucTransparency1.Location = new System.Drawing.Point(128, 0);
            this.ucTransparency1.Name = "ucTransparency1";
            this.ucTransparency1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.ucTransparency1.Size = new System.Drawing.Size(194, 22);
            this.ucTransparency1.TabIndex = 1;
            this.ucTransparency1.TransparencyPercent = 0;
            // 
            // BasicFillControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.linkLabelTools);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.comboBoxPattern);
            this.Controls.Add(this.ucTransparency2);
            this.Controls.Add(this.colorPickerBackground);
            this.Controls.Add(this.ucTransparency1);
            this.Controls.Add(this.colorPickerForeground);
            this.Name = "BasicFillControl";
            this.Size = new System.Drawing.Size(322, 80);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ColorSelectorSmall colorPickerForeground;
        private TransparencyControl ucTransparency1;
        private ColorSelectorSmall colorPickerBackground;
        private TransparencyControl ucTransparency2;
        private System.Windows.Forms.ComboBox comboBoxPattern;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.LinkLabel linkLabelTools;
    }
}
