using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows.Forms;
using VA = VisioAutomation;

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

        private void from_text()
        {
            var sso = System.StringSplitOptions.RemoveEmptyEntries;

            var seps = new[] {',', ' '};
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
                    else if (tline.StartsWith("#") && tline.Length == 1)
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
                        var tokens = tline.Split(seps, sso);
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
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            if (this.tabControl1.SelectedTab == this.tabPageFromText)
            {
                this.from_text();
            }
            else if (this.tabControl1.SelectedTab == this.tabPageFromOnline)
            {
                this.from_online();
            }

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private string GetURL()
        {
            return this.textBoxURL.Text.Trim();
        }

        private void from_online()
        {
            var url = new System.Uri(this.GetURL());

            string authority = url.Authority.ToLower();
            if (authority == "colourlovers.com" || authority == "www.colourlovers.com")
            {
                var tokens = url.AbsolutePath.ToLower().Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
                if (tokens.Length != 2 && tokens.Length != 3)
                {
                    return;
                }

                if (tokens[0] != "palette")
                {
                    return;
                }

                int palnum = -1;
                if (!int.TryParse(tokens[1], out palnum))
                {
                    return;
                }


                var new_url = "http://www.colourlovers.com/api/palette/" + palnum.ToString();

                var wc = new System.Net.WebClient();
                var data = wc.DownloadString(new_url);

                var xdoc = System.Xml.Linq.XDocument.Parse(data);
                var root = xdoc.Root;

                var palette_el = root.Element("palette");
                var colors_el = palette_el.Element("colors");
                var hex_els = colors_el.Elements("hex").ToList();

                foreach (var hex_el in hex_els)
                {
                    var color_str = "#" + hex_el.Value;
                    var c = System.Drawing.ColorTranslator.FromHtml(color_str);
                    this.Colors.Add(c);
                }
            }
            else if (authority == "kuler.adobe.com")
            {
                var tokens = url.Fragment.ToLower().Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
                if (tokens.Length != 2 && tokens.Length != 3)
                {
                    return;
                }

                if (tokens[0] != "#themeid")
                {
                    return;
                }

                int palnum = -1;
                if (!int.TryParse(tokens[1], out palnum))
                {
                    return;
                }

 
                var new_url = "http://kuler.adobe.com/kuler/API/rss/search.cfm?searchQuery=themeid:" + palnum.ToString() + "&key=85387E02911F6599FB93A8EFF7173821";

                var wc = new System.Net.WebClient();
                var data = wc.DownloadString(new_url);

                var xdoc = System.Xml.Linq.XDocument.Parse(data);
                var root = xdoc.Root;
                
                foreach (var e in root.DescendantsAndSelf())
                {
                    if (e.Name.Namespace != System.Xml.Linq.XNamespace.None)
                    {
                        e.Name = System.Xml.Linq.XNamespace.None.GetName(e.Name.LocalName);
                    }
                    if (e.Attributes().Where(a => a.IsNamespaceDeclaration || a.Name.Namespace != System.Xml.Linq.XNamespace.None).Any())
                    {
                        e.ReplaceAttributes(e.Attributes().Select(a => a.IsNamespaceDeclaration ? null : a.Name.Namespace != System.Xml.Linq.XNamespace.None ? new System.Xml.Linq.XAttribute(System.Xml.Linq.XNamespace.None.GetName(a.Name.LocalName), a.Value) : a));
                    }
                }

                var t = root.ToString();

                var palette_el = root.Element("channel").Element("item").Element("themeItem");
                var colors_el = palette_el.Element("themeSwatches");
                var hex_els = colors_el.Elements("swatch").ToList();

                foreach (var hex_el in hex_els)
                {
                    var shex = hex_el.Element("swatchHexColor");
                    var color_str = "#" + shex.Value;
                    var c = System.Drawing.ColorTranslator.FromHtml(color_str);
                    this.Colors.Add(c);
                }
                
            }
            else
            {
                
            }

        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private byte getcomp(string s)
        {
            var rs = s.Trim();

            int ri = 0;
            int.TryParse(rs, out ri);
            ri = System.Math.Max(0, System.Math.Min(255, ri));
            return (byte) ri;
        }
    }
}