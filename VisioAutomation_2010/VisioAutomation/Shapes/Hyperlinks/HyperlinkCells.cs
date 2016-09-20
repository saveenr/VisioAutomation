using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.CellGroups.Queries;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes.Hyperlinks
{
    public class HyperlinkCells : ShapeSheet.CellGroups.CellGroupMultiRow
    {
        public ShapeSheet.CellData Address { get; set; }
        public ShapeSheet.CellData Description { get; set; }
        public ShapeSheet.CellData ExtraInfo { get; set; }
        public ShapeSheet.CellData Frame { get; set; }
        public ShapeSheet.CellData SortKey { get; set; }
        public ShapeSheet.CellData SubAddress { get; set; }

        public ShapeSheet.CellData NewWindow { get; set; }
        public ShapeSheet.CellData Default { get; set; }
        public ShapeSheet.CellData Invisible { get; set; }

        public override IEnumerable<SRCFormulaPair> Pairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SRCConstants.Hyperlink_Address, this.Address.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Hyperlink_Description, this.Description.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Hyperlink_ExtraInfo, this.ExtraInfo.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Hyperlink_Frame, this.Frame.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Hyperlink_SortKey, this.SortKey.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Hyperlink_SubAddress, this.SubAddress.Formula);


                yield return this.newpair(ShapeSheet.SRCConstants.Hyperlink_NewWindow, this.NewWindow.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Hyperlink_Default, this.Default.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Hyperlink_Invisible, this.Invisible.Formula);

            }
        }

        public static IList<List<HyperlinkCells>> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = HyperlinkCells.lazy_query.Value;
            return query.GetCellGroups(page, shapeids);
        }

        public static IList<HyperlinkCells> GetCells(IVisio.Shape shape)
        {
            var query = HyperlinkCells.lazy_query.Value;
            return query.GetCellGroups(shape);
        }

        private static System.Lazy<HyperlinkCellsQuery> lazy_query = new System.Lazy<HyperlinkCellsQuery>();
    }
}