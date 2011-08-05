using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;

namespace VisioAutomation.Text
{
    public class CharacterFormatCells: VA.ShapeSheet.CellSectionDataGroup
    {
        public VA.ShapeSheet.CellData<int> Color { get; set; }
        public VA.ShapeSheet.CellData<int> Font { get; set; }
        public VA.ShapeSheet.CellData<double> Size { get; set; }
        public VA.ShapeSheet.CellData<VA.Text.CharStyle> Style { get; set; }
        public VA.ShapeSheet.CellData<double> Transparency { get; set; }

        protected override void _Apply(VA.ShapeSheet.CellSectionDataGroup.ApplyFormula func, short row)
        {
            func(VA.ShapeSheet.SRCConstants.Char_Color.ForRow(row), this.Color.Formula);
            func(VA.ShapeSheet.SRCConstants.Char_Font.ForRow(row), this.Font.Formula);
            func(VA.ShapeSheet.SRCConstants.Char_Size.ForRow(row), this.Size.Formula);
            func(VA.ShapeSheet.SRCConstants.Char_Style.ForRow(row), this.Style.Formula);
            func(VA.ShapeSheet.SRCConstants.Char_ColorTrans.ForRow(row), this.Transparency.Formula);
        }

        internal static IList<List<CharacterFormatCells>> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = new CharacterFormatQuery();
            return VA.ShapeSheet.CellSectionDataGroup._GetCells(page, shapeids, query, get_cells_from_row);
        }

        internal static IList<CharacterFormatCells> GetCells(IVisio.Shape shape)
        {
            var query = new CharacterFormatQuery();
            return VA.ShapeSheet.CellSectionDataGroup._GetCells(shape, query, get_cells_from_row);
        }

        private static CharacterFormatCells get_cells_from_row(CharacterFormatQuery query, VA.ShapeSheet.Query.QueryDataSet<double> qds, int row)
        {
            var cells = new CharacterFormatCells();
            cells.Color = qds.GetItem(row, query.Color).ToInt();
            cells.Transparency = qds.GetItem(row, query.Trans);
            cells.Font = qds.GetItem(row, query.Font).ToInt();
            cells.Size = qds.GetItem(row, query.Size);
            cells.Style = qds.GetItem(row, query.Style).Cast(v => (VA.Text.CharStyle) ((int) v));

            return cells;
        }
    }
}