using System.Collections.Generic;
using VA=VisioAutomation;
using SXL = System.Xml.Linq;

namespace VisioAutomation.VDX.Elements
{
    public class Page : Node
    {
        public readonly ShapeList Shapes;
        public Sections.PageProperties PageProperties;
        public Sections.PageLayout PageLayout;
        public Sections.PrintProperties PrintProperties;

        private readonly int _id;
        public string Name;
        public readonly List<Connect> Connects;
        public readonly LayerList Layers;
        public Drawing Drawing;

        private static readonly Internal.IDGenerator idgen = new Internal.IDGenerator(0);

        public Page(double width, double height)
        {
            if (width < 0)
            {
                throw new System.ArgumentOutOfRangeException("width");
            }

            if (height < 0)
            {
                throw new System.ArgumentOutOfRangeException("height");
            }

            this.Shapes = new ShapeList(this);
            this.Connects = new List<Connect>();
            this.PageProperties = new Sections.PageProperties();
            this.PageProperties.PageWidth.Result = width;
            this.PageProperties.PageHeight.Result = height;
            this.PrintProperties = new Sections.PrintProperties();
            this.PageLayout = new Sections.PageLayout();
            this._id = Page.idgen.GetNextID();
            var culture = System.Globalization.CultureInfo.InvariantCulture;
            this.Name = string.Format(culture, "Page-{0}", this._id + 1);
            this.Layers = new LayerList();
        }

        public int ID
        {
            get { return this._id; }
        }

        internal void AddToElement(SXL.XElement parent)
        {
            var page_el = Internal.XMLUtil.CreateVisioSchema2003Element("Page");
            var invariant_culture = System.Globalization.CultureInfo.InvariantCulture;
            page_el.SetAttributeValue("ID", this._id.ToString(invariant_culture));
            page_el.SetAttributeValue("NameU", this.Name);

            var pagesheet_el = Internal.XMLUtil.CreateVisioSchema2003Element("PageSheet");
            page_el.Add(pagesheet_el);

            foreach (var layer in this.Layers.Items)
            {
                layer.AddToElement(pagesheet_el);
            }

            this.PageProperties.AddToElement(pagesheet_el);
            this.PageLayout.AddToElement(pagesheet_el);
            this.PrintProperties.AddToElement(pagesheet_el);
            var shapes_el = Internal.XMLUtil.CreateVisioSchema2003Element("Shapes");
            page_el.Add(shapes_el);

            foreach (var vshape in this.Shapes.Items)
            {
                vshape.AddToElement(shapes_el);
            }

            if (this.Connects.Count > 0)
            {
                var xconnects = Internal.XMLUtil.CreateVisioSchema2003Element("Connects");
                foreach (var connect in this.Connects)
                {
                    connect.AddToElement(xconnects);
                }

                page_el.Add(xconnects);
            }
            parent.Add(page_el);
        }

        public void ConnectShapesViaConnector(Shape connectorshape, Shape shape1, Shape shape2)
        {
            if (shape1 == null)
            {
                throw new System.ArgumentNullException("shape1");
            }
            if (shape2 == null)
            {
                throw new System.ArgumentNullException("shape2");
            }

            if (shape1 == shape2)
            {
                throw new System.ArgumentException("cannot connect shape to itself");
            }

            var connect1 = new Connect(connectorshape, "BeginX", shape1, "PinX");
            var connect2 = new Connect(connectorshape, "EndX", shape2, "PinX");

            this.Connects.Add(connect1);
            this.Connects.Add(connect2);
        }

        public Layer AddLayer(string name, int index)
        {
            var layer1 = new Layer(name,index);
            this.Layers.Add(layer1);
            return layer1;
        }
    }
}