namespace VisioAutomation.UI
{
    partial class ColorSelectorLarge
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;



        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.panelColor = new System.Windows.Forms.Panel();
            this.buttonOK = new System.Windows.Forms.Button();
            this.pictureBoxHue = new System.Windows.Forms.PictureBox();
            this.buttonClose = new System.Windows.Forms.Button();
            this.bindingSource1 = new System.Windows.Forms.BindingSource(this.components);
            this.pictureBoxGradient = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxHue)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxGradient)).BeginInit();
            this.SuspendLayout();
            // 
            // panelColor
            // 
            this.panelColor.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panelColor.Location = new System.Drawing.Point(5, 5);
            this.panelColor.Name = "panelColor";
            this.panelColor.Size = new System.Drawing.Size(70, 70);
            this.panelColor.TabIndex = 2;
            // 
            // buttonOK
            // 
            this.buttonOK.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonOK.Location = new System.Drawing.Point(178, 232);
            this.buttonOK.Name = "buttonOK";
            this.buttonOK.Size = new System.Drawing.Size(42, 23);
            this.buttonOK.TabIndex = 4;
            this.buttonOK.Text = "OK";
            this.buttonOK.UseVisualStyleBackColor = true;
            this.buttonOK.Click += new System.EventHandler(this.buttonOK_Click);
            // 
            // pictureBoxHue
            // 
            this.pictureBoxHue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pictureBoxHue.Location = new System.Drawing.Point(77, 209);
            this.pictureBoxHue.Name = "pictureBoxHue";
            this.pictureBoxHue.Size = new System.Drawing.Size(202, 20);
            this.pictureBoxHue.TabIndex = 5;
            this.pictureBoxHue.TabStop = false;
            this.pictureBoxHue.Paint += new System.Windows.Forms.PaintEventHandler(this.pictureBoxHue_Paint);
            this.pictureBoxHue.MouseDown += new System.Windows.Forms.MouseEventHandler(this.pictureBoxHue_MouseDown);
            this.pictureBoxHue.MouseMove += new System.Windows.Forms.MouseEventHandler(this.pictureBoxHue_MouseMove);
            // 
            // buttonClose
            // 
            this.buttonClose.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonClose.Location = new System.Drawing.Point(226, 232);
            this.buttonClose.Name = "buttonClose";
            this.buttonClose.Size = new System.Drawing.Size(52, 23);
            this.buttonClose.TabIndex = 6;
            this.buttonClose.Text = "Close";
            this.buttonClose.UseVisualStyleBackColor = true;
            this.buttonClose.Click += new System.EventHandler(this.buttonClose_Click);
            // 
            // pictureBoxGradient
            // 
            this.pictureBoxGradient.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pictureBoxGradient.Location = new System.Drawing.Point(77, 5);
            this.pictureBoxGradient.Name = "pictureBoxGradient";
            this.pictureBoxGradient.Size = new System.Drawing.Size(202, 202);
            this.pictureBoxGradient.TabIndex = 10;
            this.pictureBoxGradient.TabStop = false;
            this.pictureBoxGradient.Paint += new System.Windows.Forms.PaintEventHandler(this.pictureBoxGradient_Paint);
            this.pictureBoxGradient.MouseDown += new System.Windows.Forms.MouseEventHandler(this.pictureBoxGradient_MouseDown);
            this.pictureBoxGradient.MouseMove += new System.Windows.Forms.MouseEventHandler(this.pictureBoxGradient_MouseMove);
            this.pictureBoxGradient.MouseUp += new System.Windows.Forms.MouseEventHandler(this.pictureBoxGradient_MouseUp);
            // 
            // ColorSelectorLarge
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.Controls.Add(this.pictureBoxGradient);
            this.Controls.Add(this.buttonClose);
            this.Controls.Add(this.pictureBoxHue);
            this.Controls.Add(this.buttonOK);
            this.Controls.Add(this.panelColor);
            this.Name = "ColorSelectorLarge";
            this.Size = new System.Drawing.Size(282, 259);
            this.Paint += new System.Windows.Forms.PaintEventHandler(this.ColorSelectorLarge_Paint);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxHue)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxGradient)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.BindingSource bindingSource1;
        private System.Windows.Forms.Panel panelColor;
        private System.Windows.Forms.Button buttonOK;
        private System.Windows.Forms.PictureBox pictureBoxHue;
        private System.Windows.Forms.Button buttonClose;
        private System.Windows.Forms.PictureBox pictureBoxGradient;
    }
}
