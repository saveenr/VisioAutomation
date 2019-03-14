using System.Collections.Generic;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Shapes
{
    public class HyperlinkCells : CellGroupBase
    {
        public CellValueLiteral Address { get; set; }
        public CellValueLiteral Description { get; set; }
        public CellValueLiteral ExtraInfo { get; set; }
        public CellValueLiteral Frame { get; set; }
        public CellValueLiteral SortKey { get; set; }
        public CellValueLiteral SubAddress { get; set; }
        public CellValueLiteral NewWindow { get; set; }
        public CellValueLiteral Default { get; set; }
        public CellValueLiteral Invisible { get; set; }

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return SrcValuePair.Create(SrcConstants.HyperlinkAddress, this.Address);
                yield return SrcValuePair.Create(SrcConstants.HyperlinkDescription, this.Description);
                yield return SrcValuePair.Create(SrcConstants.HyperlinkExtraInfo, this.ExtraInfo);
                yield return SrcValuePair.Create(SrcConstants.HyperlinkFrame, this.Frame);
                yield return SrcValuePair.Create(SrcConstants.HyperlinkSortKey, this.SortKey);
                yield return SrcValuePair.Create(SrcConstants.HyperlinkSubAddress, this.SubAddress);
                yield return SrcValuePair.Create(SrcConstants.HyperlinkNewWindow, this.NewWindow);
                yield return SrcValuePair.Create(SrcConstants.HyperlinkDefault, this.Default);
                yield return SrcValuePair.Create(SrcConstants.HyperlinkInvisible, this.Invisible);
            }
        }

        public static List<List<HyperlinkCells>> GetCells(IVisio.Page page, IList<int> shapeids, CellValueType type)
        {
            var reader = lazy_reader.Value;
            return reader.GetCellsMultiRow(page, shapeids, type);
        }

        public static List<HyperlinkCells> GetCells(IVisio.Shape shape, CellValueType type)
        {
            var reader = lazy_reader.Value;
            return reader.GetCellsMultiRow(shape, type);
        }

        private static readonly System.Lazy<HyperlinkCellsReader> lazy_reader = new System.Lazy<HyperlinkCellsReader>();


        class HyperlinkCellsReader : CellGroupReader<HyperlinkCells>
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

            public HyperlinkCellsReader() : base(new VisioAutomation.ShapeSheet.Query.SectionsQuery())
            {
                var sec = this.query_multirow.SectionQueries.Add(IVisio.VisSectionIndices.visSectionHyperlink);

                this.Address = sec.Columns.Add(SrcConstants.HyperlinkAddress, nameof(this.Address));
                this.Default = sec.Columns.Add(SrcConstants.HyperlinkDefault, nameof(this.Default));
                this.Description = sec.Columns.Add(SrcConstants.HyperlinkDescription, nameof(this.Description));
                this.ExtraInfo = sec.Columns.Add(SrcConstants.HyperlinkExtraInfo, nameof(this.ExtraInfo));
                this.Frame = sec.Columns.Add(SrcConstants.HyperlinkFrame, nameof(this.Frame));
                this.Invisible = sec.Columns.Add(SrcConstants.HyperlinkInvisible, nameof(this.Invisible));
                this.NewWindow = sec.Columns.Add(SrcConstants.HyperlinkNewWindow, nameof(this.NewWindow));
                this.SortKey = sec.Columns.Add(SrcConstants.HyperlinkSortKey, nameof(this.SortKey));
                this.SubAddress = sec.Columns.Add(SrcConstants.HyperlinkSubAddress, nameof(this.SubAddress));
            }

            public override HyperlinkCells ToCellGroup(ShapeSheet.Internal.ArraySegment<string> row)
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