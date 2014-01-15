using VisioAutomation.VDX.Internal.Extensions;
using System.Collections.Generic;
using VisioAutomation.VDX.Internal;
using VisioAutomation.VDX.ShapeSheet;
using SXL = System.Xml.Linq;

namespace VisioAutomation.VDX.Sections
{
    public class Geom
    {
        public BoolCell NoFill = new BoolCell();
        public BoolCell NoLine = new BoolCell();
        public BoolCell NoShow = new BoolCell();
        public BoolCell NoSnap = new BoolCell();

        public readonly List<GeomRow> Rows;

        public Geom()
        {
            this.Rows = new List<GeomRow>();
        }

        public void AddToElement(SXL.XElement parent, int index)
        {
            var el = XMLUtil.CreateVisioSchema2003Element("Geom");
            el.SetAttributeValueInt("IX", index);
            el.Add(this.NoFill.ToXml("NoFill"));
            el.Add(this.NoLine.ToXml("NoLine"));
            el.Add(this.NoShow.ToXml("NoShow"));
            el.Add(this.NoSnap.ToXml("NoSnap"));

            int ix = 0;
            foreach (var geomrow in this.Rows)
            {
                geomrow.AddToElement(el, ix + 1);
                ix++;
            }

            parent.Add(el);
        }
    }
}