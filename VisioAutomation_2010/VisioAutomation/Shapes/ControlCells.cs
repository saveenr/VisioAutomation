using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;


namespace VisioAutomation.Shapes
{
    public class ControlCells : ShapeSheet.CellGroups.CellGroupMultiRow
    {
        public VisioAutomation.ShapeSheet.CellValueLiteral CanGlue { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Tip { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral X { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Y { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral YBehavior { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral XBehavior { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral XDynamics { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral YDynamics { get; set; }

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ControlCanGlue, this.CanGlue.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ControlTip, this.Tip.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ControlX, this.X.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ControlY, this.Y.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ControlYBehavior, this.YBehavior.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ControlXBehavior, this.XBehavior.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ControlXDynamics, this.XDynamics.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ControlYDynamics, this.YDynamics.Value);
            }
        }

        public static List<List<ControlCells>> GetFormulas(IVisio.Page page, IList<int> shapeids)
        {
            var query = ControlCells.lazy_query.Value;
            return query.GetFormulas(page, shapeids);
        }

        public static List<List<ControlCells>> GetResults(IVisio.Page page, IList<int> shapeids)
        {
            var query = ControlCells.lazy_query.Value;
            return query.GetResults(page, shapeids);
        }


        public static List<ControlCells> GetFormulas(IVisio.Shape shape)
        {
            var query = ControlCells.lazy_query.Value;
            return query.GetFormulas(shape);
        }

        public static List<ControlCells> GetResults(IVisio.Shape shape)
        {
            var query = ControlCells.lazy_query.Value;
            return query.GetResults(shape);
        }

        private static readonly System.Lazy<ControlCellsReader> lazy_query = new System.Lazy<ControlCellsReader>();

        class ControlCellsReader : ReaderMultiRow<ControlCells>
        {
            public SectionQueryColumn CanGlue { get; set; }
            public SectionQueryColumn Tip { get; set; }
            public SectionQueryColumn X { get; set; }
            public SectionQueryColumn Y { get; set; }
            public SectionQueryColumn YBehavior { get; set; }
            public SectionQueryColumn XBehavior { get; set; }
            public SectionQueryColumn XDynamics { get; set; }
            public SectionQueryColumn YDynamics { get; set; }

            public ControlCellsReader()
            {
                var sec = this.query.SectionQueries.Add(IVisio.VisSectionIndices.visSectionControls);

                this.CanGlue = sec.Columns.Add(SrcConstants.ControlCanGlue, nameof(SrcConstants.ControlCanGlue));
                this.Tip = sec.Columns.Add(SrcConstants.ControlTip, nameof(SrcConstants.ControlTip));
                this.X = sec.Columns.Add(SrcConstants.ControlX, nameof(SrcConstants.ControlX));
                this.Y = sec.Columns.Add(SrcConstants.ControlY, nameof(SrcConstants.ControlY));
                this.YBehavior = sec.Columns.Add(SrcConstants.ControlYBehavior, nameof(SrcConstants.ControlYBehavior));
                this.XBehavior = sec.Columns.Add(SrcConstants.ControlXBehavior, nameof(SrcConstants.ControlXBehavior));
                this.XDynamics = sec.Columns.Add(SrcConstants.ControlXDynamics, nameof(SrcConstants.ControlXDynamics));
                this.YDynamics = sec.Columns.Add(SrcConstants.ControlYDynamics, nameof(SrcConstants.ControlYDynamics));

            }

            public override ControlCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<string> row)
            {
                var cells = new ControlCells();
                cells.CanGlue = row[this.CanGlue];
                cells.Tip = row[this.Tip];
                cells.X = row[this.X];
                cells.Y = row[this.Y];
                cells.YBehavior = row[this.YBehavior];
                cells.XBehavior = row[this.XBehavior];
                cells.XDynamics = row[this.XDynamics];
                cells.YDynamics = row[this.YDynamics];
                return cells;
            }
        }

    }
}