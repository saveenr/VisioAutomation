using VisioAutomation.VDX.Internal;
using VisioAutomation.VDX.ShapeSheet;
using SXL = System.Xml.Linq;

namespace VisioAutomation.VDX.Sections
{
    public class PrintProperties
    {
        public DistanceCell PageLeftMargin = new DistanceCell();
        public DistanceCell PageRightMargin = new DistanceCell();
        public DistanceCell PageTopMargin = new DistanceCell();
        public DistanceCell PageBottomMargin = new DistanceCell();

        public DoubleCell ScaleX = new DoubleCell();
        public DoubleCell ScaleY = new DoubleCell();

        public IntCell PagesX = new IntCell();
        public IntCell PagesY = new IntCell();

        public DistanceCell CenterX = new DistanceCell();
        public DistanceCell CenterY = new DistanceCell();

        public BoolCell OnPage = new BoolCell();
        public BoolCell PrintGrid = new BoolCell();
        public BoolCell PrintPageOrientation = new BoolCell();
        public IntCell PaperKind = new IntCell();
        public IntCell PaperSource = new IntCell();

        public IntCell ShdwObliqueAngle = new IntCell();
        public IntCell ShdwScaleFactor = new IntCell();

        public void AddToElement(SXL.XElement parent)
        {
            var el = XMLUtil.CreateVisioSchema2003Element("PrintProps");
            el.Add(this.PageLeftMargin.ToXml("PageLeftMargin"));
            el.Add(this.PageRightMargin.ToXml("PageRightMargin"));
            el.Add(this.PageTopMargin.ToXml("PageTopMargin"));
            el.Add(this.PageBottomMargin.ToXml("PageBottomMargin"));

            el.Add(this.ScaleX.ToXml("ScaleX"));
            el.Add(this.ScaleY.ToXml("ScaleY"));

            el.Add(this.PagesX.ToXml("PagesX"));
            el.Add(this.PagesY.ToXml("PagesY"));

            el.Add(this.CenterX.ToXml("CenterX"));
            el.Add(this.CenterY.ToXml("CenterY"));

            el.Add(this.OnPage.ToXml("OnPage"));
            el.Add(this.PrintGrid.ToXml("PrintGrid"));

            el.Add(this.PrintPageOrientation.ToXml("PrintPageOrientation"));
            el.Add(this.PaperKind.ToXml("PaperKind"));

            el.Add(this.PaperSource.ToXml("PaperSource"));
            el.Add(this.ShdwObliqueAngle.ToXml("ShdwObliqueAngle"));

            el.Add(this.ShdwScaleFactor.ToXml("ShdwScaleFactor"));

            parent.Add(el);
        }
    }
}