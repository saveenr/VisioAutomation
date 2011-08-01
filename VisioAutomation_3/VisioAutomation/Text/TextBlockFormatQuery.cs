using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.Text
{
    class TextBlockFormatQuery : VA.ShapeSheet.Query.CellQuery
    {
        public VA.ShapeSheet.Query.CellQueryColumn BottomMargin { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn LeftMargin { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn RightMargin { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn TopMargin { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn DefaultTabStop { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn TextBkgnd { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn TextBkgndTrans { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn TextDirection { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn VerticalAlign { get; set; }

        public TextBlockFormatQuery() :
            base()
        {
            BottomMargin = this.AddColumn(VA.ShapeSheet.SRCConstants.BottomMargin, "BottomMargin");
            LeftMargin = this.AddColumn(VA.ShapeSheet.SRCConstants.LeftMargin, "LeftMargin");
            RightMargin = this.AddColumn(VA.ShapeSheet.SRCConstants.RightMargin, "RightMargin");
            TopMargin = this.AddColumn(VA.ShapeSheet.SRCConstants.TopMargin, "TopMargin");


            DefaultTabStop = this.AddColumn(VA.ShapeSheet.SRCConstants.DefaultTabStop, "DefaultTabStop");
            TextBkgnd = this.AddColumn(VA.ShapeSheet.SRCConstants.TextBkgnd, "TextBkgnd");
            TextBkgndTrans = this.AddColumn(VA.ShapeSheet.SRCConstants.TextBkgndTrans, "TextBkgndTrans");
            TextDirection = this.AddColumn(VA.ShapeSheet.SRCConstants.TextDirection, "TextDirection");
            VerticalAlign = this.AddColumn(VA.ShapeSheet.SRCConstants.VerticalAlign, "VerticalAlign");
        }
    }
}