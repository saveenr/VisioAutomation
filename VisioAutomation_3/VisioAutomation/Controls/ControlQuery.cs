using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.Controls
{
    class ControlQuery : VA.ShapeSheet.Query.SectionQuery
    {
        public VA.ShapeSheet.Query.SectionQueryColumn Glue { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn Tip { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn X { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn Y { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn YCon { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn XCon { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn XDyn { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn YDyn { get; set; }

        public ControlQuery() :
            base(IVisio.VisSectionIndices.visSectionControls)
        {
            Glue = AddColumn(IVisio.VisCellIndices.visCtlGlue, "Glue");
            Tip = AddColumn(IVisio.VisCellIndices.visCtlTip, "Tip");
            X = AddColumn(IVisio.VisCellIndices.visCtlX, "X");
            Y = AddColumn(IVisio.VisCellIndices.visCtlY, "Y");
            XDyn = AddColumn(IVisio.VisCellIndices.visCtlXDyn, "XDyn");
            YDyn = AddColumn(IVisio.VisCellIndices.visCtlYDyn, "YDyn");
            XCon = AddColumn(IVisio.VisCellIndices.visCtlXCon, "XCon");
            YCon = AddColumn(IVisio.VisCellIndices.visCtlYCon, "YCon");
        }
    }
}