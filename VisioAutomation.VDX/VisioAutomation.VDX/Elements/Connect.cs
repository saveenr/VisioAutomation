using VisioAutomation.VDX.Internal;
using SXL = System.Xml.Linq;

namespace VisioAutomation.VDX.Elements
{
    public class Connect
    {
        public Connect(Shape from_shape, string from_cell, Shape to_shape, string to_cell)
        {
            if (string.IsNullOrEmpty(to_cell))
            {
                throw new System.ArgumentException("to_cell cannot be null or empty", "to_cell");
            }

            if (string.IsNullOrEmpty(from_cell))
            {
                throw new System.ArgumentException("from_cell cannot be null or empty", "from_cell");
            }

            this.FromSheet = from_shape.ID;
            this.FromCell = from_cell;
            this.ToSheet = to_shape.ID;
            this.ToCell = to_cell;
        }

        public int FromSheet { get; set; }
        public int ToSheet { get; set; }
        public string FromCell { get; set; }
        public string ToCell { get; set; }
        public int? FromPart { get; set; }
        public int? ToPart { get; set; }

        public void AddToElement(SXL.XElement parent)
        {
            var connect_el = XMLUtil.CreateVisioSchema2003Element("Connect");
            connect_el.SetAttributeValue("FromSheet", this.FromSheet);

            if (this.FromCell != null)
            {
                connect_el.SetAttributeValue("FromCell", this.FromCell);
            }

            if (this.FromPart.HasValue)
            {
                connect_el.SetAttributeValue("FromPart", this.FromPart.Value);
            }

            connect_el.SetAttributeValue("ToSheet", this.ToSheet);

            if (this.ToCell != null)
            {
                connect_el.SetAttributeValue("ToCell", this.ToCell);
            }

            if (this.ToPart.HasValue)
            {
                connect_el.SetAttributeValue("ToPart", this.ToPart.Value);
            }

            parent.Add(connect_el);
        }
    }
}