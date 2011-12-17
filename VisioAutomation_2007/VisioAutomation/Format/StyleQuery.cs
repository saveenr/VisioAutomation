using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Format
{
    internal class StyleQuery : VA.ShapeSheet.Query.CellQuery
    {
       public VA.ShapeSheet.Query.CellQueryColumn EnableFillProps  {get; set;}
       public VA.ShapeSheet.Query.CellQueryColumn EnableLineProps  {get; set;}
       public VA.ShapeSheet.Query.CellQueryColumn EnableTextProps  {get; set;}
       public VA.ShapeSheet.Query.CellQueryColumn HideForApply  {get; set;}

        public  StyleQuery () :
                base(IVisio.VisSectionIndices.  )
        {
            this.EnableFillProps = this.AddColumn(VA.ShapeSheet.SRCConstants. EnableFillProps , "EnableFillProps");
            this.EnableLineProps = this.AddColumn(VA.ShapeSheet.SRCConstants. EnableLineProps , "EnableLineProps");
            this.EnableTextProps = this.AddColumn(VA.ShapeSheet.SRCConstants. EnableTextProps , "EnableTextProps");
            this.HideForApply = this.AddColumn(VA.ShapeSheet.SRCConstants. HideForApply , "HideForApply");
        }

    }
}