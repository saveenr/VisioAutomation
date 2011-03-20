using System.Collections.Generic;
using VisioAutomation.VDX.Internal;

namespace VisioAutomation.VDX.Elements
{
    public class Text
    {
        private List<TextRun> m_runs;

        public Text()
        {
            this.m_runs = new List<TextRun>(0);
        }

        private List<TextRun> Runs
        {
            get { return this.m_runs; }
        }

        public void Add(TextRun run)
        {
            this.m_runs.Add(run);
        }

        public void Add(string text)
        {
            var run = new TextRun(text);
            this.m_runs.Add(run);
        }

        public void Add(string text, int? cp, int? pp, int? tp)
        {
            var tr = new TextRun(text, cp, pp, tp);
            this.Add(tr);
        }

        public void AddToElement(System.Xml.Linq.XElement parent)
        {
            if (this.Runs != null)
            {
                var text_el = XMLUtil.CreateVisioSchema2003Element("Text");
                foreach (var ft in this.m_runs)
                {
                    if (ft.CharacterFormatIndex.HasValue)
                    {
                        var xcp = XMLUtil.CreateVisioSchema2003Element("cp");
                        xcp.SetAttributeValue("IX", ft.CharacterFormatIndex.Value);
                        text_el.Add(xcp);
                    }
                    if (ft.ParagraphFormatIndex.HasValue)
                    {
                        var xpp = XMLUtil.CreateVisioSchema2003Element("pp");
                        xpp.SetAttributeValue("IX", ft.ParagraphFormatIndex.Value);
                        text_el.Add(xpp);
                    }
                    if (ft.TabsFormatIndex.HasValue)
                    {
                        var xtp = XMLUtil.CreateVisioSchema2003Element("tp");
                        xtp.SetAttributeValue("IX", ft.TabsFormatIndex.Value);
                        text_el.Add(xtp);
                    }
                    text_el.Add(ft.Text);
                }
                parent.Add(text_el);
            }
        }
    }
}