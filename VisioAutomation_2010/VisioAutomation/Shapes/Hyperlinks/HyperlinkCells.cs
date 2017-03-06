using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
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

        public override IEnumerable<SrcFormulaPair> SrcFormulaPairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SrcConstants.HyperlinkAddress, this.Address.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.HyperlinkDescription, this.Description.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.HyperlinkExtraInfo, this.ExtraInfo.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.HyperlinkFrame, this.Frame.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.HyperlinkSortKey, this.SortKey.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.HyperlinkSubAddress, this.SubAddress.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.HyperlinkNewWindow, this.NewWindow.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.HyperlinkDefault, this.Default.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.HyperlinkInvisible, this.Invisible.Formula);
            }
        }

        public static List<List<HyperlinkCells>> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = HyperlinkCells.lazy_query.Value;
            return query.GetCellGroups(page, shapeids);
        }

        public static List<HyperlinkCells> GetCells(IVisio.Shape shape)
        {
            var query = HyperlinkCells.lazy_query.Value;
            return query.GetCellGroups(shape);
        }

        private static readonly System.Lazy<HyperlinkCellsReader> lazy_query = new System.Lazy<HyperlinkCellsReader>();
    }
}