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
        public List<System.Drawing.Color> Colors; 
        public FormImportColors()
        {
            InitializeComponent();
            this.Colors = new List<System.Drawing.Color>();
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            var sso = System.StringSplitOptions.RemoveEmptyEntries;

            var seps = new[] {',',' '};
            int linenum = 1;
            foreach (string line in this.textBox1.Lines)
            {
                var tline = line.Trim();
                if (tline.Length > 0)
                {
                    if (tline.StartsWith("//"))
                    {
                        // skip comments
                        continue;
                    }
                    else if (tline.StartsWith("#") && tline.Length==1)
                    {
                        // skip comments
                        continue;
                    }
                    else if (tline.StartsWith("#"))
                    {
                        var rgb = System.Drawing.ColorTranslator.FromHtml(tline);
                        this.Colors.Add(rgb);
                    }
                    else
                    {
                        var tokens = tline.Split(seps,sso);
                        if (tokens.Length >= 3)
                        {

                            var components = tokens.Select(v => getcomp(v)).ToArray();

                            bool has_alpha = components.Length > 3;
                            int i = has_alpha ? 1 : 0;
                            var a = components[0];
                            var r = components[i + 0];
                            var g = components[i + 1];
                            var b = components[i + 2];

                            if (has_alpha)
                            {
                                var rgb = System.Drawing.Color.FromArgb(a, r, g, b);
                                this.Colors.Add(rgb);
                            }
                            else
                            {
                                var rgb = System.Drawing.Color.FromArgb(r, g, b);
                                this.Colors.Add(rgb);
                            }
                        }
                    }
                }

                linenum++;
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
