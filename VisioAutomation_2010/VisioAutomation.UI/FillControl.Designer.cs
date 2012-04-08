using VA=VisioAutomation;

namespace VisioAutomation.UI
{
    partial class FillControl
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
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.basicFillControlFill = new BasicFillControl();
            this.basicFillControlShadow = new BasicFillControl();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(75, 11);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(19, 13);
            this.label1.TabIndex = 5;
            this.label1.Text = "Fill";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(75, 129);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(46, 13);
            this.label2.TabIndex = 11;
            this.label2.Text = "Shadow";
            // 
            // basicFillControlFill
            // 
            this.basicFillControlFill.BackgroundColor = System.Drawing.Color.Black;
            this.basicFillControlFill.BackgroundTransparency = 0;
            this.basicFillControlFill.FillPattern = VA.UI.FillPattern.None;
            this.basicFillControlFill.ForegroundColor = System.Drawing.Color.Red;
            this.basicFillControlFill.ForegroundTransparency = 0;
            this.basicFillControlFill.Location = new System.Drawing.Point(3, 27);
            this.basicFillControlFill.Name = "basicFillControlFill";
            this.basicFillControlFill.Size = new System.Drawing.Size(330, 87);
            this.basicFillControlFill.TabIndex = 10;
            // 
            // basicFillControlShadow
            // 
            this.basicFillControlShadow.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.basicFillControlShadow.BackgroundTransparency = 0;
            this.basicFillControlShadow.FillPattern = VA.UI.FillPattern.None;
            this.basicFillControlShadow.ForegroundColor = System.Drawing.Color.Blue;
            this.basicFillControlShadow.ForegroundTransparency = 0;
            this.basicFillControlShadow.Location = new System.Drawing.Point(3, 145);
            this.basicFillControlShadow.Name = "basicFillControlShadow";
            this.basicFillControlShadow.Size = new System.Drawing.Size(330, 88);
            this.basicFillControlShadow.TabIndex = 9;
            // 
            // FillControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.label2);
            this.Controls.Add(this.basicFillControlFill);
            this.Controls.Add(this.basicFillControlShadow);
            this.Controls.Add(this.label1);
            this.Name = "FillControl";
            this.Size = new System.Drawing.Size(336, 240);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private BasicFillControl basicFillControlShadow;
        private BasicFillControl basicFillControlFill;
        private System.Windows.Forms.Label label2;
    }
}
