using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes.Hyperlinks
{
    public class HyperlinkCells : ShapeSheetQuery.CellGroups.CellGroupMultiRow
    {
        public ShapeSheet.CellData<string> Address { get; set; }
        public ShapeSheet.CellData<string> Description { get; set; }
        public ShapeSheet.CellData<string> ExtraInfo { get; set; }
        public ShapeSheet.CellData<string> Frame { get; set; }
        public ShapeSheet.CellData<string> SortKey { get; set; }
        public ShapeSheet.CellData<string> SubAddress { get; set; }

        public ShapeSheet.CellData<bool> NewWindow { get; set; }
        public ShapeSheet.CellData<bool> Default { get; set; }
        public ShapeSheet.CellData<bool> Invisible { get; set; }

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
            return ShapeSheetQuery.CellGroups.CellGroupMultiRow._GetCells<HyperlinkCells, double>(page, shapeids, query, query.GetCells);
        }

        public static IList<HyperlinkCells> GetCells(IVisio.Shape shape)
        {
            var query = HyperlinkCells.lazy_query.Value;
            return ShapeSheetQuery.CellGroups.CellGroupMultiRow._GetCells<HyperlinkCells, double>(shape, query, query.GetCells);
        }

        private static System.Lazy<ShapeSheetQuery.Common.HyperlinkCellsQuery> lazy_query = new System.Lazy<ShapeSheetQuery.Common.HyperlinkCellsQuery>();
    }
}