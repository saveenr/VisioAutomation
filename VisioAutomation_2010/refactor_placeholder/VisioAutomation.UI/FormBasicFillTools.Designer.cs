namespace VisioAutomation.UI
{
    partial class FormBasicFillTools
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
            this.buttonCancel = new System.Windows.Forms.Button();
            this.buttonOK = new System.Windows.Forms.Button();
            this.buttonSwapColors = new System.Windows.Forms.Button();
            this.colorSelectorSmallForeground = new ColorSelectorSmall();
            this.colorSelectorSmallBackground = new ColorSelectorSmall();
            this.buttonCopyFgtoBg = new System.Windows.Forms.Button();
            this.buttonCopyBgToFg = new System.Windows.Forms.Button();
            this.labelfg = new System.Windows.Forms.Label();
            this.labelbg = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // buttonCancel
            // 
            this.buttonCancel.FlatAppearance.BorderColor = System.Drawing.SystemColors.ActiveBorder;
            this.buttonCancel.Location = new System.Drawing.Point(225, 120);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(75, 23);
            this.buttonCancel.TabIndex = 0;
            this.buttonCancel.Text = "Cancel";
            this.buttonCancel.UseVisualStyleBackColor = true;
            this.buttonCancel.Click += new System.EventHandler(this.buttonCancel_Click);
            // 
            // buttonOK
            // 
            this.buttonOK.FlatAppearance.BorderColor = System.Drawing.SystemColors.ActiveBorder;
            this.buttonOK.Location = new System.Drawing.Point(144, 120);
            this.buttonOK.Name = "buttonOK";
            this.buttonOK.Size = new System.Drawing.Size(75, 23);
            this.buttonOK.TabIndex = 1;
            this.buttonOK.Text = "OK";
            this.buttonOK.UseVisualStyleBackColor = true;
            this.buttonOK.Click += new System.EventHandler(this.buttonOK_Click);
            // 
            // buttonSwapColors
            // 
            this.buttonSwapColors.FlatAppearance.BorderColor = System.Drawing.SystemColors.ActiveBorder;
            this.buttonSwapColors.Location = new System.Drawing.Point(225, 27);
            this.buttonSwapColors.Name = "buttonSwapColors";
            this.buttonSwapColors.Size = new System.Drawing.Size(75, 23);
            this.buttonSwapColors.TabIndex = 2;
            this.buttonSwapColors.Text = "Swap";
            this.buttonSwapColors.UseVisualStyleBackColor = true;
            this.buttonSwapColors.Click += new System.EventHandler(this.buttonSwapColors_Click);
            // 
            // colorSelectorSmallForeground
            // 
            this.colorSelectorSmallForeground.Color = System.Drawing.SystemColors.Control;
            this.colorSelectorSmallForeground.Location = new System.Drawing.Point(19, 28);
            this.colorSelectorSmallForeground.Name = "colorSelectorSmallForeground";
            this.colorSelectorSmallForeground.Size = new System.Drawing.Size(42, 22);
            this.colorSelectorSmallForeground.TabIndex = 3;
            // 
            // colorSelectorSmallBackground
            // 
            this.colorSelectorSmallBackground.Color = System.Drawing.SystemColors.Control;
            this.colorSelectorSmallBackground.Location = new System.Drawing.Point(19, 75);
            this.colorSelectorSmallBackground.Name = "colorSelectorSmallBackground";
            this.colorSelectorSmallBackground.Size = new System.Drawing.Size(42, 22);
            this.colorSelectorSmallBackground.TabIndex = 4;
            // 
            // buttonCopyFgtoBg
            // 
            this.buttonCopyFgtoBg.FlatAppearance.BorderColor = System.Drawing.SystemColors.ActiveBorder;
            this.buttonCopyFgtoBg.Location = new System.Drawing.Point(78, 75);
            this.buttonCopyFgtoBg.Name = "buttonCopyFgtoBg";
            this.buttonCopyFgtoBg.Size = new System.Drawing.Size(118, 23);
            this.buttonCopyFgtoBg.TabIndex = 5;
            this.buttonCopyFgtoBg.Text = "Set foreground";
            this.buttonCopyFgtoBg.UseVisualStyleBackColor = true;
            this.buttonCopyFgtoBg.Click += new System.EventHandler(this.buttonCopyFgtoBg_Click);
            // 
            // buttonCopyBgToFg
            // 
            this.buttonCopyBgToFg.FlatAppearance.BorderColor = System.Drawing.SystemColors.ActiveBorder;
            this.buttonCopyBgToFg.Location = new System.Drawing.Point(78, 28);
            this.buttonCopyBgToFg.Name = "buttonCopyBgToFg";
            this.buttonCopyBgToFg.Size = new System.Drawing.Size(118, 23);
            this.buttonCopyBgToFg.TabIndex = 6;
            this.buttonCopyBgToFg.Text = "Set background";
            this.buttonCopyBgToFg.UseVisualStyleBackColor = true;
            this.buttonCopyBgToFg.Click += new System.EventHandler(this.buttonCopyBgToFg_Click);
            // 
            // labelfg
            // 
            this.labelfg.AutoSize = true;
            this.labelfg.Location = new System.Drawing.Point(16, 12);
            this.labelfg.Name = "labelfg";
            this.labelfg.Size = new System.Drawing.Size(69, 13);
            this.labelfg.TabIndex = 7;
            this.labelfg.Text = "Foreground";
            // 
            // labelbg
            // 
            this.labelbg.AutoSize = true;
            this.labelbg.Location = new System.Drawing.Point(16, 59);
            this.labelbg.Name = "labelbg";
            this.labelbg.Size = new System.Drawing.Size(70, 13);
            this.labelbg.TabIndex = 8;
            this.labelbg.Text = "Background";
            // 
            // FormBasicFillTools
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(313, 150);
            this.Controls.Add(this.labelbg);
            this.Controls.Add(this.labelfg);
            this.Controls.Add(this.buttonCopyBgToFg);
            this.Controls.Add(this.buttonCopyFgtoBg);
            this.Controls.Add(this.colorSelectorSmallBackground);
            this.Controls.Add(this.colorSelectorSmallForeground);
            this.Controls.Add(this.buttonSwapColors);
            this.Controls.Add(this.buttonOK);
            this.Controls.Add(this.buttonCancel);
            this.Name = "FormBasicFillTools";
            this.Text = "Fill options";
            this.Load += new System.EventHandler(this.FormBasicFillTools_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonCancel;
        private System.Windows.Forms.Button buttonOK;
        private System.Windows.Forms.Button buttonSwapColors;
        private ColorSelectorSmall colorSelectorSmallForeground;
        private ColorSelectorSmall colorSelectorSmallBackground;
        private System.Windows.Forms.Button buttonCopyFgtoBg;
        private System.Windows.Forms.Button buttonCopyBgToFg;
        private System.Windows.Forms.Label labelfg;
        private System.Windows.Forms.Label labelbg;
    }
}