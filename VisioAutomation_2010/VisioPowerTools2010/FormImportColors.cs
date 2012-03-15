using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using VisioAutomation.Drawing;
using VA=VisioAutomation;

namespace VisioPowerTools2010
{
    public partial class FormImportColors : Form
    {
        public List<VA.Drawing.ColorRGB> Colors; 
        public FormImportColors()
        {
            InitializeComponent();
            this.Colors = new List<ColorRGB>();
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            var seps = new[] {','};
            foreach (string line in this.textBox1.Lines)
            {
                var tline = line.Trim();
                if (tline.Length > 0)
                {
                    var tokens = tline.Split(seps);
                    if (tokens.Length >= 3)
                    {
                        var r = getcomp(tokens[0]);
                        var g = getcomp(tokens[1]);
                        var b = getcomp(tokens[2]);

                        var rgb = new VA.Drawing.ColorRGB(r, g, b);

                        this.Colors.Add(rgb);

                    }
                }
            }

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();

        }

        byte getcomp(string s)
        {
            var rs = s.Trim();

            int ri = 0;
            int.TryParse(rs, out ri);
            ri = System.Math.Max(0, System.Math.Min(255, ri));
            return (byte) ri;
        }
    }
}
