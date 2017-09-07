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
                yield return SrcFormulaPair.Create(ShapeSheet.SrcConstants.HyperlinkAddress, this.Address.Value);
                yield return SrcFormulaPair.Create(ShapeSheet.SrcConstants.HyperlinkDescription, this.Description.Value);
                yield return SrcFormulaPair.Create(ShapeSheet.SrcConstants.HyperlinkExtraInfo, this.ExtraInfo.Value);
                yield return SrcFormulaPair.Create(ShapeSheet.SrcConstants.HyperlinkFrame, this.Frame.Value);
                yield return SrcFormulaPair.Create(ShapeSheet.SrcConstants.HyperlinkSortKey, this.SortKey.Value);
                yield return SrcFormulaPair.Create(ShapeSheet.SrcConstants.HyperlinkSubAddress, this.SubAddress.Value);
                yield return SrcFormulaPair.Create(ShapeSheet.SrcConstants.HyperlinkNewWindow, this.NewWindow.Value);
                yield return SrcFormulaPair.Create(ShapeSheet.SrcConstants.HyperlinkDefault, this.Default.Value);
                yield return SrcFormulaPair.Create(ShapeSheet.SrcConstants.HyperlinkInvisible, this.Invisible.Value);
            }
        }

        public static List<List<HyperlinkCells>> GetFormulas(IVisio.Page page, IList<int> shapeids)
        {
            var query = HyperlinkCells.lazy_query.Value;
            return query.GetFormulas(page, shapeids);
        }

        public static List<List<HyperlinkCells>> GetResults(IVisio.Page page, IList<int> shapeids)
        {
            var query = HyperlinkCells.lazy_query.Value;
            return query.GetResults(page, shapeids);
        }

        public static List<HyperlinkCells> GetFormulas(IVisio.Shape shape)
        {
            var query = HyperlinkCells.lazy_query.Value;
            return query.GetFormulas(shape);
        }


        public static List<HyperlinkCells> GetResults(IVisio.Shape shape)
        {
            var query = HyperlinkCells.lazy_query.Value;
            return query.GetResults(shape);
        }

        private static readonly System.Lazy<HyperlinkCellsReader> lazy_query = new System.Lazy<HyperlinkCellsReader>();
    }
}