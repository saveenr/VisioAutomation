namespace VisioPowerTools
{
    partial class FormTextTool : System.Windows.Forms.Form
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
            this.buttonSwitchTextCase = new System.Windows.Forms.Button();
            this.buttonTextToBottom = new System.Windows.Forms.Button();
            this.buttonResizeToFitText = new System.Windows.Forms.Button();
            this.buttonEnableTextWrapping = new System.Windows.Forms.Button();
            this.buttonDisableTextWrapping = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.labelText = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // buttonSwitchTextCase
            // 
            this.buttonSwitchTextCase.FlatAppearance.BorderColor = System.Drawing.SystemColors.ActiveBorder;
            this.buttonSwitchTextCase.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonSwitchTextCase.Location = new System.Drawing.Point(10, 28);
            this.buttonSwitchTextCase.Name = "buttonSwitchTextCase";
            this.buttonSwitchTextCase.Size = new System.Drawing.Size(80, 23);
            this.buttonSwitchTextCase.TabIndex = 0;
            this.buttonSwitchTextCase.Text = "Toggle Case";
            this.buttonSwitchTextCase.UseVisualStyleBackColor = true;
            this.buttonSwitchTextCase.Click += new System.EventHandler(this.buttonSwitchTextCase_Click);
            // 
            // buttonTextToBottom
            // 
            this.buttonTextToBottom.FlatAppearance.BorderColor = System.Drawing.SystemColors.ActiveBorder;
            this.buttonTextToBottom.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonTextToBottom.Location = new System.Drawing.Point(10, 90);
            this.buttonTextToBottom.Name = "buttonTextToBottom";
            this.buttonTextToBottom.Size = new System.Drawing.Size(117, 23);
            this.buttonTextToBottom.TabIndex = 2;
            this.buttonTextToBottom.Text = "Move below shape";
            this.buttonTextToBottom.UseVisualStyleBackColor = true;
            this.buttonTextToBottom.Click += new System.EventHandler(this.buttonTextToBottom_Click);
            // 
            // buttonResizeToFitText
            // 
            this.buttonResizeToFitText.FlatAppearance.BorderColor = System.Drawing.SystemColors.ActiveBorder;
            this.buttonResizeToFitText.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonResizeToFitText.Location = new System.Drawing.Point(10, 141);
            this.buttonResizeToFitText.Name = "buttonResizeToFitText";
            this.buttonResizeToFitText.Size = new System.Drawing.Size(117, 23);
            this.buttonResizeToFitText.TabIndex = 3;
            this.buttonResizeToFitText.Text = "Resize to Fit Text";
            this.buttonResizeToFitText.UseVisualStyleBackColor = true;
            this.buttonResizeToFitText.Click += new System.EventHandler(this.buttonResizeToFitText_Click);
            // 
            // buttonEnableTextWrapping
            // 
            this.buttonEnableTextWrapping.FlatAppearance.BorderColor = System.Drawing.SystemColors.ActiveBorder;
            this.buttonEnableTextWrapping.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonEnableTextWrapping.Location = new System.Drawing.Point(10, 197);
            this.buttonEnableTextWrapping.Name = "buttonEnableTextWrapping";
            this.buttonEnableTextWrapping.Size = new System.Drawing.Size(60, 23);
            this.buttonEnableTextWrapping.TabIndex = 4;
            this.buttonEnableTextWrapping.Text = "Enable";
            this.buttonEnableTextWrapping.UseVisualStyleBackColor = true;
            this.buttonEnableTextWrapping.Click += new System.EventHandler(this.buttonEnableTextWrapping_Click);
            // 
            // buttonDisableTextWrapping
            // 
            this.buttonDisableTextWrapping.FlatAppearance.BorderColor = System.Drawing.SystemColors.ActiveBorder;
            this.buttonDisableTextWrapping.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonDisableTextWrapping.Location = new System.Drawing.Point(76, 197);
            this.buttonDisableTextWrapping.Name = "buttonDisableTextWrapping";
            this.buttonDisableTextWrapping.Size = new System.Drawing.Size(60, 23);
            this.buttonDisableTextWrapping.TabIndex = 5;
            this.buttonDisableTextWrapping.Text = "Disable";
            this.buttonDisableTextWrapping.UseVisualStyleBackColor = true;
            this.buttonDisableTextWrapping.Click += new System.EventHandler(this.buttonDisableTextWrapping_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(10, 180);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(77, 13);
            this.label2.TabIndex = 7;
            this.label2.Text = "Text Wrapping";
            // 
            // labelText
            // 
            this.labelText.AutoSize = true;
            this.labelText.Location = new System.Drawing.Point(10, 11);
            this.labelText.Name = "labelText";
            this.labelText.Size = new System.Drawing.Size(67, 13);
            this.labelText.TabIndex = 8;
            this.labelText.Text = "Text content";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(10, 73);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(48, 13);
            this.label3.TabIndex = 9;
            this.label3.Text = "Text box";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(10, 124);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(59, 13);
            this.label4.TabIndex = 10;
            this.label4.Text = "Shape size";
            // 
            // FormTextTool
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(151, 233);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.labelText);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.buttonDisableTextWrapping);
            this.Controls.Add(this.buttonEnableTextWrapping);
            this.Controls.Add(this.buttonResizeToFitText);
            this.Controls.Add(this.buttonTextToBottom);
            this.Controls.Add(this.buttonSwitchTextCase);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "FormTextTool";
            this.Text = "Text";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonSwitchTextCase;
        private System.Windows.Forms.Button buttonTextToBottom;
        private System.Windows.Forms.Button buttonResizeToFitText;
        private System.Windows.Forms.Button buttonEnableTextWrapping;
        private System.Windows.Forms.Button buttonDisableTextWrapping;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label labelText;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
    }
}