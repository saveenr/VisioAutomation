using System.Collections.Generic;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;



namespace VisioAutomation.Shapes
{
    public static class LockHelper
    {
        public static List<LockCells> GetLockCells(IVisio.Page page, IList<int> shapeids, CellValueType type)
        {
            var reader = LockCells_lazy_reader.Value;
            return reader.GetCellsSingleRow(page, shapeids, type);
        }

        public static LockCells GetLockCells(IVisio.Shape shape, CellValueType type)
        {
            var reader = LockCells_lazy_reader.Value;
            return reader.GetCellsSingleRow(shape, type);
        }

        private static readonly System.Lazy<LockCellsReader> LockCells_lazy_reader = new System.Lazy<LockCellsReader>();


        class LockCellsReader : CellGroupReader<LockCells>
        {
            public LockCellsReader() : base(VisioAutomation.ShapeSheet.CellGroups.CellGroupReaderType.SingleRow)
            {
                InitializeQuery();
            }

            public override LockCells ToCellGroup(ShapeSheet.Internal.ArraySegment<string> row)
            {
                var cells = new LockCells();
                var cols = this.query_singlerow.Columns;

                string getcellvalue(string name)
                {
                    return row[cols[name].Ordinal];
                }

                cells.Aspect = getcellvalue(nameof(LockCells.Aspect));
                cells.Begin = getcellvalue(nameof(LockCells.Begin));
                cells.CalcWH = getcellvalue(nameof(LockCells.CalcWH));
                cells.Crop = getcellvalue(nameof(LockCells.Crop));
                cells.CustProp = getcellvalue(nameof(LockCells.CustProp));
                cells.Delete = getcellvalue(nameof(LockCells.Delete));
                cells.End = getcellvalue(nameof(LockCells.End));
                cells.Format = getcellvalue(nameof(LockCells.Format));
                cells.FromGroupFormat = getcellvalue(nameof(LockCells.FromGroupFormat));
                cells.Group = getcellvalue(nameof(LockCells.Group));
                cells.Height = getcellvalue(nameof(LockCells.Height));
                cells.MoveX = getcellvalue(nameof(LockCells.MoveX));
                cells.MoveY = getcellvalue(nameof(LockCells.MoveY));
                cells.Rotate = getcellvalue(nameof(LockCells.Rotate));
                cells.Select = getcellvalue(nameof(LockCells.Select));
                cells.TextEdit = getcellvalue(nameof(LockCells.TextEdit));
                cells.ThemeColors = getcellvalue(nameof(LockCells.ThemeColors));
                cells.ThemeEffects = getcellvalue(nameof(LockCells.ThemeEffects));
                cells.VertexEdit = getcellvalue(nameof(LockCells.VertexEdit));
                cells.Width = getcellvalue(nameof(LockCells.Width));
                return cells;
            }
        }

    }
}