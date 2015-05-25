namespace VisioPowerTools2010
{
    partial class FormCreateStyle
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
            this.labelName = new System.Windows.Forms.Label();
            this.textName = new System.Windows.Forms.TextBox();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.buttonOK = new System.Windows.Forms.Button();
            this.checkBoxIncludesText = new System.Windows.Forms.CheckBox();
            this.checkBoxIncludesLIne = new System.Windows.Forms.CheckBox();
            this.checkBoxIncludesFill = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // labelName
            // 
            this.labelName.AutoSize = true;
            this.labelName.Location = new System.Drawing.Point(13, 13);
            this.labelName.Name = "labelName";
            this.labelName.Size = new System.Drawing.Size(35, 13);
            this.labelName.TabIndex = 0;
            this.labelName.Text = "Name";
            // 
            // textName
            // 
            this.textName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textName.Location = new System.Drawing.Point(16, 30);
            this.textName.Name = "textName";
            this.textName.Size = new System.Drawing.Size(443, 20);
            this.textName.TabIndex = 1;
            // 
            // buttonCancel
            // 
            this.buttonCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.buttonCancel.Location = new System.Drawing.Point(384, 113);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(75, 23);
            this.buttonCancel.TabIndex = 2;
            this.buttonCancel.Text = "Cancel";
            this.buttonCancel.UseVisualStyleBackColor = true;
            // 
            // buttonOK
            // 
            this.buttonOK.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.buttonOK.Location = new System.Drawing.Point(303, 113);
            this.buttonOK.Name = "buttonOK";
            this.buttonOK.Size = new System.Drawing.Size(75, 23);
            this.buttonOK.TabIndex = 3;
            this.buttonOK.Text = "OK";
            this.buttonOK.UseVisualStyleBackColor = true;
            // 
            // checkBoxIncludesText
            // 
            this.checkBoxIncludesText.AutoSize = true;
            this.checkBoxIncludesText.Checked = true;
            this.checkBoxIncludesText.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxIncludesText.Location = new System.Drawing.Point(16, 65);
            this.checkBoxIncludesText.Name = "checkBoxIncludesText";
            this.checkBoxIncludesText.Size = new System.Drawing.Size(86, 17);
            this.checkBoxIncludesText.TabIndex = 4;
            this.checkBoxIncludesText.Text = "Includes text";
            this.checkBoxIncludesText.UseVisualStyleBackColor = true;
            // 
            // checkBoxIncludesLIne
            // 
            this.checkBoxIncludesLIne.AutoSize = true;
            this.checkBoxIncludesLIne.Checked = true;
            this.checkBoxIncludesLIne.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxIncludesLIne.Location = new System.Drawing.Point(16, 88);
            this.checkBoxIncludesLIne.Name = "checkBoxIncludesLIne";
            this.checkBoxIncludesLIne.Size = new System.Drawing.Size(85, 17);
            this.checkBoxIncludesLIne.TabIndex = 5;
            this.checkBoxIncludesLIne.Text = "Includes line";
            this.checkBoxIncludesLIne.UseVisualStyleBackColor = true;
            // 
            // checkBoxIncludesFill
            // 
            this.checkBoxIncludesFill.AutoSize = true;
            this.checkBoxIncludesFill.Checked = true;
            this.checkBoxIncludesFill.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxIncludesFill.Location = new System.Drawing.Point(16, 111);
            this.checkBoxIncludesFill.Name = "checkBoxIncludesFill";
            this.checkBoxIncludesFill.Size = new System.Drawing.Size(78, 17);
            this.checkBoxIncludesFill.TabIndex = 6;
            this.checkBoxIncludesFill.Text = "Includes fill";
            this.checkBoxIncludesFill.UseVisualStyleBackColor = true;
            // 
            // FormCreateStyle
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(471, 148);
            this.Controls.Add(this.checkBoxIncludesFill);
            this.Controls.Add(this.checkBoxIncludesLIne);
            this.Controls.Add(this.checkBoxIncludesText);
            this.Controls.Add(this.buttonOK);
            this.Controls.Add(this.buttonCancel);
            this.Controls.Add(this.textName);
            this.Controls.Add(this.labelName);
            this.Name = "FormCreateStyle";
            this.Text = "Create Style";
            this.Load += new System.EventHandler(this.FormCreateStyle_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label labelName;
        private System.Windows.Forms.TextBox textName;
        private System.Windows.Forms.Button buttonCancel;
        private System.Windows.Forms.Button buttonOK;
        private System.Windows.Forms.CheckBox checkBoxIncludesText;
        private System.Windows.Forms.CheckBox checkBoxIncludesLIne;
        private System.Windows.Forms.CheckBox checkBoxIncludesFill;
    }
}