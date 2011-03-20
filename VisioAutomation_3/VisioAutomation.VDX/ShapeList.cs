namespace VisioAutomation.VDX
{
    public class ShapeList : NamedNodeList<Elements.Shape>
    {
        private Elements.Page page_el;

        public ShapeList(Elements.Page page_el) :
            base(shape => shape.Name)
        {
            this.page_el = page_el;
        }

        public override void Add(Elements.Shape shape)
        {
            if (this.page_el.Drawing == null)
            {
                throw new System.ArgumentException(
                    "page must to added to a drawing before shapes can be added to the page");
            }

            var master_md = this.page_el.Drawing.GetMasterMetData(shape.Master);


            shape.Page = this.page_el;
            shape._id = this.page_el.Drawing.GetNextShapeID();
            var culture = System.Globalization.CultureInfo.InvariantCulture;
            shape.Name = string.Format(culture, "Shape.{0}", shape._id);

            base.Add(shape);

            this.page_el.Drawing.AccountForMasteSubshapes(master_md.SubShapeCount);
        }
    }
}