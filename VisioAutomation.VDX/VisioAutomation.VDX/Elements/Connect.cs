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

            FromSheet = from_shape.ID;
            FromCell = from_cell;
            ToSheet = to_shape.ID;
            ToCell = to_cell;
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
            connect_el.SetAttributeValue("FromSheet", FromSheet);

            if (this.FromCell != null)
            {
                connect_el.SetAttributeValue("FromCell", FromCell);
            }

            if (this.FromPart.HasValue)
            {
                connect_el.SetAttributeValue("FromPart", FromPart.Value);
            }

            connect_el.SetAttributeValue("ToSheet", ToSheet);

            if (this.ToCell != null)
            {
                connect_el.SetAttributeValue("ToCell", ToCell);
            }

            if (this.ToPart.HasValue)
            {
                connect_el.SetAttributeValue("ToPart", ToPart.Value);
            }

            parent.Add(connect_el);
        }
    }
}