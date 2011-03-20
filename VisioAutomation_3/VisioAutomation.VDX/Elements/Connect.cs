using VisioAutomation.VDX.Internal;

namespace VisioAutomation.VDX.Elements
{
    public class Connect
    {
        public Connect(Shape from_shape, string from_cell, Shape to_shape, string to_cell)
        {
            if (to_shape == null)
            {
                throw new System.ArgumentNullException("shape1");
            }

            if (from_shape == null)
            {
                throw new System.ArgumentNullException("connectorshape");
            }

            if (string.IsNullOrEmpty(to_cell))
            {
                throw new System.ArgumentException("cell1");
            }

            if (string.IsNullOrEmpty(from_cell))
            {
                throw new System.ArgumentException("cell2");
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

        public void AddToElement(System.Xml.Linq.XElement parent)
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