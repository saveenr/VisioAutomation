using System.Collections.Generic;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.ShapeSheet.Query;

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

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.HyperlinkAddress, this.Address.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.HyperlinkDescription, this.Description.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.HyperlinkExtraInfo, this.ExtraInfo.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.HyperlinkFrame, this.Frame.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.HyperlinkSortKey, this.SortKey.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.HyperlinkSubAddress, this.SubAddress.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.HyperlinkNewWindow, this.NewWindow.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.HyperlinkDefault, this.Default.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.HyperlinkInvisible, this.Invisible.Value);
            }
        }

        public static List<List<HyperlinkCells>> GetValues(IVisio.Page page, IList<int> shapeids, CellValueType cvt)
        {
            var query = HyperlinkCells.lazy_query.Value;
            return query.GetValues(page, shapeids, cvt);
        }

        public static List<HyperlinkCells> GetValues(IVisio.Shape shape, CellValueType cvt)
        {
            var query = HyperlinkCells.lazy_query.Value;
            return query.GetValues(shape, cvt);
        }

        private static readonly System.Lazy<HyperlinkCellsReader> lazy_query = new System.Lazy<HyperlinkCellsReader>();


        class HyperlinkCellsReader : ReaderMultiRow<HyperlinkCells>
        {

            public SectionQueryColumn Address { get; set; }
            public SectionQueryColumn Description { get; set; }
            public SectionQueryColumn ExtraInfo { get; set; }
            public SectionQueryColumn Frame { get; set; }
            public SectionQueryColumn SortKey { get; set; }
            public SectionQueryColumn SubAddress { get; set; }
            public SectionQueryColumn NewWindow { get; set; }
            public SectionQueryColumn Default { get; set; }
            public SectionQueryColumn Invisible { get; set; }

            public HyperlinkCellsReader()
            {
                var sec = this.query.SectionQueries.Add(IVisio.VisSectionIndices.visSectionHyperlink);

                this.Address = sec.Columns.Add(ShapeSheet.SrcConstants.HyperlinkAddress, nameof(ShapeSheet.SrcConstants.HyperlinkAddress));
                this.Default = sec.Columns.Add(ShapeSheet.SrcConstants.HyperlinkDefault, nameof(ShapeSheet.SrcConstants.HyperlinkDefault));
                this.Description = sec.Columns.Add(ShapeSheet.SrcConstants.HyperlinkDescription, nameof(ShapeSheet.SrcConstants.HyperlinkDescription));
                this.ExtraInfo = sec.Columns.Add(ShapeSheet.SrcConstants.HyperlinkExtraInfo, nameof(ShapeSheet.SrcConstants.HyperlinkExtraInfo));
                this.Frame = sec.Columns.Add(ShapeSheet.SrcConstants.HyperlinkFrame, nameof(ShapeSheet.SrcConstants.HyperlinkFrame));
                this.Invisible = sec.Columns.Add(ShapeSheet.SrcConstants.HyperlinkInvisible, nameof(ShapeSheet.SrcConstants.HyperlinkInvisible));
                this.NewWindow = sec.Columns.Add(ShapeSheet.SrcConstants.HyperlinkNewWindow, nameof(ShapeSheet.SrcConstants.HyperlinkNewWindow));
                this.SortKey = sec.Columns.Add(ShapeSheet.SrcConstants.HyperlinkSortKey, nameof(ShapeSheet.SrcConstants.HyperlinkSortKey));
                this.SubAddress = sec.Columns.Add(ShapeSheet.SrcConstants.HyperlinkSubAddress, nameof(ShapeSheet.SrcConstants.HyperlinkSubAddress));
            }

            public override HyperlinkCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<string> row)
            {
                var cells = new HyperlinkCells();

                cells.Address = row[this.Address];
                cells.Description = row[this.Description];
                cells.ExtraInfo = row[this.ExtraInfo];
                cells.Frame = row[this.Frame];
                cells.SortKey = row[this.SortKey];
                cells.SubAddress = row[this.SubAddress];
                cells.NewWindow = row[this.NewWindow];
                cells.Default = row[this.Default];
                cells.Invisible = row[this.Invisible];

                return cells;
            }
        }

    }
}