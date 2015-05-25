using System.ComponentModel;
using System.Windows.Forms;

namespace VisioPowerTools2010
{
    partial class FormImportColors
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private IContainer components = null;

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
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.buttonOK = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPageFromText = new System.Windows.Forms.TabPage();
            this.tabPageFromOnline = new System.Windows.Forms.TabPage();
            this.textBoxURL = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.labelURL = new System.Windows.Forms.Label();
            this.labelExamples = new System.Windows.Forms.Label();
            this.tabControl1.SuspendLayout();
            this.tabPageFromText.SuspendLayout();
            this.tabPageFromOnline.SuspendLayout();
            this.SuspendLayout();
            // 
            // textBox1
            // 
            this.textBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox1.Location = new System.Drawing.Point(6, 6);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(343, 195);
            this.textBox1.TabIndex = 0;
            this.textBox1.Text = "// basic rgb\r\n239, 62, 54\r\n13,117,144\r\n// basic argb\r\n128, 13,117,144\r\n//webcolor" +
    "\r\n#ff0000\r\n//webcolor with alpha\r\n#80ff0000\r\n\r\n";
            // 
            // buttonCancel
            // 
            this.buttonCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonCancel.Location = new System.Drawing.Point(299, 249);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(75, 23);
            this.buttonCancel.TabIndex = 1;
            this.buttonCancel.Text = "Cancel";
            this.buttonCancel.UseVisualStyleBackColor = true;
            this.buttonCancel.Click += new System.EventHandler(this.buttonCancel_Click);
            // 
            // buttonOK
            // 
            this.buttonOK.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonOK.Location = new System.Drawing.Point(218, 249);
            this.buttonOK.Name = "buttonOK";
            this.buttonOK.Size = new System.Drawing.Size(75, 23);
            this.buttonOK.TabIndex = 2;
            this.buttonOK.Text = "OK";
            this.buttonOK.UseVisualStyleBackColor = true;
            this.buttonOK.Click += new System.EventHandler(this.buttonOK_Click);
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Controls.Add(this.tabPageFromText);
            this.tabControl1.Controls.Add(this.tabPageFromOnline);
            this.tabControl1.Location = new System.Drawing.Point(12, 13);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(363, 233);
            this.tabControl1.TabIndex = 3;
            // 
            // tabPageFromText
            // 
            this.tabPageFromText.Controls.Add(this.textBox1);
            this.tabPageFromText.Location = new System.Drawing.Point(4, 22);
            this.tabPageFromText.Name = "tabPageFromText";
            this.tabPageFromText.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageFromText.Size = new System.Drawing.Size(355, 207);
            this.tabPageFromText.TabIndex = 0;
            this.tabPageFromText.Text = "From Text";
            this.tabPageFromText.UseVisualStyleBackColor = true;
            // 
            // tabPageFromOnline
            // 
            this.tabPageFromOnline.Controls.Add(this.labelExamples);
            this.tabPageFromOnline.Controls.Add(this.labelURL);
            this.tabPageFromOnline.Controls.Add(this.textBox2);
            this.tabPageFromOnline.Controls.Add(this.textBoxURL);
            this.tabPageFromOnline.Location = new System.Drawing.Point(4, 22);
            this.tabPageFromOnline.Name = "tabPageFromOnline";
            this.tabPageFromOnline.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageFromOnline.Size = new System.Drawing.Size(355, 207);
            this.tabPageFromOnline.TabIndex = 1;
            this.tabPageFromOnline.Text = "From Online";
            this.tabPageFromOnline.UseVisualStyleBackColor = true;
            // 
            // textBoxURL
            // 
            this.textBoxURL.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxURL.Location = new System.Drawing.Point(7, 34);
            this.textBoxURL.Name = "textBoxURL";
            this.textBoxURL.Size = new System.Drawing.Size(342, 20);
            this.textBoxURL.TabIndex = 0;
            this.textBoxURL.Text = "http://kuler.adobe.com/#themeID/1785951";
            // 
            // textBox2
            // 
            this.textBox2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox2.Location = new System.Drawing.Point(7, 86);
            this.textBox2.Multiline = true;
            this.textBox2.Name = "textBox2";
            this.textBox2.ReadOnly = true;
            this.textBox2.Size = new System.Drawing.Size(342, 115);
            this.textBox2.TabIndex = 1;
            this.textBox2.Text = "// ColourLovers example URL\r\nhttp://www.colourlovers.com/palette/2074058\r\n\r\n// Ad" +
    "obe Kuler example URL\r\nhttp://kuler.adobe.com/#themeID/1785951\r\n";
            // 
            // labelURL
            // 
            this.labelURL.AutoSize = true;
            this.labelURL.Location = new System.Drawing.Point(7, 15);
            this.labelURL.Name = "labelURL";
            this.labelURL.Size = new System.Drawing.Size(176, 13);
            this.labelURL.TabIndex = 2;
            this.labelURL.Text = "URL (ColourLovers or Adober Kuler)";
            // 
            // labelExamples
            // 
            this.labelExamples.AutoSize = true;
            this.labelExamples.Location = new System.Drawing.Point(7, 67);
            this.labelExamples.Name = "labelExamples";
            this.labelExamples.Size = new System.Drawing.Size(77, 13);
            this.labelExamples.TabIndex = 3;
            this.labelExamples.Text = "Example URLs";
            // 
            // FormImportColors
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(387, 278);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.buttonOK);
            this.Controls.Add(this.buttonCancel);
            this.Name = "FormImportColors";
            this.Text = "Import Colors";
            this.tabControl1.ResumeLayout(false);
            this.tabPageFromText.ResumeLayout(false);
            this.tabPageFromText.PerformLayout();
            this.tabPageFromOnline.ResumeLayout(false);
            this.tabPageFromOnline.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private TextBox textBox1;
        private Button buttonCancel;
        private Button buttonOK;
        private TabControl tabControl1;
        private TabPage tabPageFromText;
        private TabPage tabPageFromOnline;
        private TextBox textBoxURL;
        private TextBox textBox2;
        private Label labelExamples;
        private Label labelURL;
    }
}