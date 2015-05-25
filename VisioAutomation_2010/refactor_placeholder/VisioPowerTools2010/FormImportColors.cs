using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using SXL=System.Xml.Linq;
using VA = VisioAutomation;

namespace VisioPowerTools2010
{
    public partial class FormImportColors : Form
    {
        public List<System.Drawing.Color> Colors;

        public FormImportColors()
        {
            this.InitializeComponent();
            this.Colors = new List<System.Drawing.Color>();
        }

        private void from_text()
        {
            this.Colors.Clear();

            var seps = new[] {',', ' '};

            foreach (string rawline in this.textBox1.Lines)
            {
                var line = rawline.Trim();
                if (line.Length > 0)
                {
                    if (line.StartsWith("//"))
                    {
                        // skip comments
                        continue;
                    }
                    else if (line.StartsWith("#"))
                    {
                        var rgb = System.Drawing.ColorTranslator.FromHtml(line);
                        this.Colors.Add(rgb);
                    }
                    else
                    {
                        var tokens = line.Split(seps, StringSplitOptions.RemoveEmptyEntries);
                        if (tokens.Length >= 3)
                        {
                            var color_components = tokens.Select(this.getcomp).ToArray();

                            bool has_alpha = color_components.Length > 3;
                            int i = has_alpha ? 1 : 0;
                            var a = color_components[0];
                            var r = color_components[i + 0];
                            var g = color_components[i + 1];
                            var b = color_components[i + 2];

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
            else
            {
                string msg = "Unhandeled case";
                MessageBox.Show(msg);
            }

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private string GetURL()
        {
            return this.textBoxURL.Text.Trim();
        }

        public SXL.XDocument download_xml(string url)
        {
            var wc = new System.Net.WebClient();
            string data;
            try
            {
                data = wc.DownloadString(url);
            }
            catch (Exception)
            {
                string msg = "Failed to download data";
                MessageBox.Show(msg);
                return null;
            }

            SXL.XDocument xdoc;

            try
            {
                xdoc = SXL.XDocument.Parse(data);
            }
            catch (Exception)
            {
                string msg = "Failed to parse XML";
                MessageBox.Show(msg);
                return null;
            }
            return xdoc;
        }
        private void from_online()
        {
            this.Colors.Clear();

            var url = new Uri(this.GetURL());

            string authority = url.Authority.ToLower();
            if (authority == "colourlovers.com" || authority == "www.colourlovers.com")
            {
                var tokens = url.AbsolutePath.ToLower().Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);

                bool incorrect_format = (tokens.Length < 2) || (tokens[0] != "palette");

                if (incorrect_format)
                {
                    string msg = "Incorrect format for ColourLovers URL";
                    MessageBox.Show(msg);
                    return;                    
                }

                string palette_id = tokens[1];

                var palette_api_url = "http://www.colourlovers.com/api/palette/" + palette_id;
                var xdoc = this.download_xml(palette_api_url);

                if (xdoc == null)
                {
                    return;
                }

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

                bool incorrect_format = tokens.Length < 2 || tokens[0] != "#themeid";

                if (incorrect_format)
                {
                    string msg = "Incorrect format for kuler URL";
                    MessageBox.Show(msg);
                    return;
                }

                string theme_id = tokens[1];
 
                var theme_api_url = "http://kuler.adobe.com/kuler/API/rss/search.cfm?searchQuery=themeid:" + theme_id + "&key=85387E02911F6599FB93A8EFF7173821";
                var xdoc = this.download_xml(theme_api_url);

                if (xdoc == null)
                {
                    return;
                }

                var root = xdoc.Root;
                
                FormImportColors.strip_namespaces(root);
                
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
                string msg = "Unknown URL format";
                MessageBox.Show(msg);
            }
        }

        private static void strip_namespaces(SXL.XElement root)
        {
            foreach (var e in root.DescendantsAndSelf())
            {
                if (e.Name.Namespace != SXL.XNamespace.None)
                {
                    e.Name = SXL.XNamespace.None.GetName(e.Name.LocalName);
                }
                if (
                    e.Attributes().Any(a => a.IsNamespaceDeclaration || a.Name.Namespace != SXL.XNamespace.None))
                {
                    e.ReplaceAttributes(
                        e.Attributes().Select(
                            a =>
                            a.IsNamespaceDeclaration
                                ? null
                                : a.Name.Namespace != SXL.XNamespace.None
                                      ? new SXL.XAttribute(
                                            SXL.XNamespace.None.GetName(a.Name.LocalName), a.Value)
                                      : a));
                }
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

            int ri;
            int.TryParse(rs, out ri);
            ri = Math.Max(0, Math.Min(255, ri));
            return (byte) ri;
        }
    }
}