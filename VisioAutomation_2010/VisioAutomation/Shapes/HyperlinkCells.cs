using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public class HyperlinkCells : ShapeSheet.CellGroups.CellGroupMultiRow
    {
        public VisioAutomation.ShapeSheet.CellValueLiteral Address { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Description { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral ExtraInfo { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Frame { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral SortKey { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral SubAddress { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral NewWindow { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Default { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Invisible { get; set; }

        public override IEnumerable<SrcFormulaPair> SrcFormulaPairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SrcConstants.HyperlinkAddress, this.Address.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.HyperlinkDescription, this.Description.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.HyperlinkExtraInfo, this.ExtraInfo.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.HyperlinkFrame, this.Frame.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.HyperlinkSortKey, this.SortKey.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.HyperlinkSubAddress, this.SubAddress.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.HyperlinkNewWindow, this.NewWindow.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.HyperlinkDefault, this.Default.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.HyperlinkInvisible, this.Invisible.Value);
            }
        }

        public static List<List<HyperlinkCells>> GetCells(IVisio.Page page, IList<int> shapeids, VisioAutomation.ShapeSheet.CellValueType cvt)
        {
            var query = HyperlinkCells.lazy_query.Value;
            return query.GetCellGroups(page, shapeids, cvt);
        }

        public static List<HyperlinkCells> GetCells(IVisio.Shape shape, VisioAutomation.ShapeSheet.CellValueType cvt)
        {
            var query = HyperlinkCells.lazy_query.Value;
            return query.GetCellGroups(shape, cvt);
        }

        private static readonly System.Lazy<HyperlinkCellsReader> lazy_query = new System.Lazy<HyperlinkCellsReader>();
    }
}