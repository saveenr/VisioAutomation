using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Text
{
    public class TextCells : ShapeSheet.CellGroups.CellGroup
    {
        public ShapeSheet.CellData<double> BottomMargin { get; set; }
        public ShapeSheet.CellData<double> LeftMargin { get; set; }
        public ShapeSheet.CellData<double> RightMargin { get; set; }
        public ShapeSheet.CellData<double> TopMargin { get; set; }
        public ShapeSheet.CellData<double> DefaultTabStop { get; set; }
        public ShapeSheet.CellData<int> TextBkgnd { get; set; }
        public ShapeSheet.CellData<double> TextBkgndTrans { get; set; }
        public ShapeSheet.CellData<int> TextDirection { get; set; }
        public ShapeSheet.CellData<int> VerticalAlign { get; set; }
        public ShapeSheet.CellData<double> TxtAngle { get; set; }
        public ShapeSheet.CellData<double> TxtWidth { get; set; }
        public ShapeSheet.CellData<double> TxtHeight { get; set; }
        public ShapeSheet.CellData<double> TxtPinX { get; set; }
        public ShapeSheet.CellData<double> TxtPinY { get; set; }
        public ShapeSheet.CellData<double> TxtLocPinX { get; set; }
        public ShapeSheet.CellData<double> TxtLocPinY { get; set; }

        public override IEnumerable<SRCFormulaPair> Pairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SRCConstants.BottomMargin, this.BottomMargin.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LeftMargin, this.LeftMargin.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.RightMargin, this.RightMargin.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.TopMargin, this.TopMargin.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.DefaultTabStop, this.DefaultTabStop.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.TextBkgnd, this.TextBkgnd.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.TextBkgndTrans, this.TextBkgndTrans.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.TextDirection, this.TextDirection.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.VerticalAlign, this.VerticalAlign.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.TxtPinX, this.TxtPinX.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.TxtPinY, this.TxtPinY.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.TxtLocPinX, this.TxtLocPinX.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.TxtLocPinY, this.TxtLocPinY.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.TxtWidth, this.TxtWidth.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.TxtHeight, this.TxtHeight.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.TxtAngle, this.TxtAngle.Formula);
            }
        }

        public static IList<TextCells> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = TextCells.get_query();
            return ShapeSheet.CellGroups.CellGroup._GetCells<TextCells, double>(page, shapeids, query, query.GetCells);
        }

        public static TextCells GetCells(IVisio.Shape shape)
        {
            var query = TextCells.get_query();
            return ShapeSheet.CellGroups.CellGroup._GetCells<TextCells, double>(shape, query, query.GetCells);
        }

        private static TextBlockFormatCellQuery _mCellQuery;
        private static TextBlockFormatCellQuery get_query()
        {
            TextCells._mCellQuery= TextCells._mCellQuery ?? new TextBlockFormatCellQuery();
            return TextCells._mCellQuery;
        }

        private class TextBlockFormatCellQuery : CellQuery
        {
            public CellColumn BottomMargin { get; set; }
            public CellColumn LeftMargin { get; set; }
            public CellColumn RightMargin { get; set; }
            public CellColumn TopMargin { get; set; }
            public CellColumn DefaultTabStop { get; set; }
            public CellColumn TextBkgnd { get; set; }
            public CellColumn TextBkgndTrans { get; set; }
            public CellColumn TextDirection { get; set; }
            public CellColumn VerticalAlign { get; set; }
            public CellColumn TxtWidth { get; set; }
            public CellColumn TxtHeight { get; set; }
            public CellColumn TxtPinX { get; set; }
            public CellColumn TxtPinY { get; set; }
            public CellColumn TxtLocPinX { get; set; }
            public CellColumn TxtLocPinY { get; set; }
            public CellColumn TxtAngle { get; set; }

            public TextBlockFormatCellQuery() :
                base()
            {
                this.BottomMargin = this.AddCell(ShapeSheet.SRCConstants.BottomMargin, "BottomMargin");
                this.LeftMargin = this.AddCell(ShapeSheet.SRCConstants.LeftMargin, "LeftMargin");
                this.RightMargin = this.AddCell(ShapeSheet.SRCConstants.RightMargin, "RightMargin");
                this.TopMargin = this.AddCell(ShapeSheet.SRCConstants.TopMargin, "TopMargin");
                this.DefaultTabStop = this.AddCell(ShapeSheet.SRCConstants.DefaultTabStop, "DefaultTabStop");
                this.TextBkgnd = this.AddCell(ShapeSheet.SRCConstants.TextBkgnd, "TextBkgnd");
                this.TextBkgndTrans = this.AddCell(ShapeSheet.SRCConstants.TextBkgndTrans, "TextBkgndTrans");
                this.TextDirection = this.AddCell(ShapeSheet.SRCConstants.TextDirection, "TextDirection");
                this.VerticalAlign = this.AddCell(ShapeSheet.SRCConstants.VerticalAlign, "VerticalAlign");
                this.TxtPinX = this.AddCell(ShapeSheet.SRCConstants.TxtPinX, "TxtPinX");
                this.TxtPinY = this.AddCell(ShapeSheet.SRCConstants.TxtPinY, "TxtPinY");
                this.TxtLocPinX = this.AddCell(ShapeSheet.SRCConstants.TxtLocPinX, "TxtLocPinX");
                this.TxtLocPinY = this.AddCell(ShapeSheet.SRCConstants.TxtLocPinY, "TxtLocPinY");
                this.TxtWidth = this.AddCell(ShapeSheet.SRCConstants.TxtWidth, "TxtWidth");
                this.TxtHeight = this.AddCell(ShapeSheet.SRCConstants.TxtHeight, "TxtHeight");
                this.TxtAngle = this.AddCell(ShapeSheet.SRCConstants.TxtAngle, "TxtAngle");

            }

            public TextCells GetCells(IList<ShapeSheet.CellData<double>> row)
            {
                var cells = new TextCells();
                cells.BottomMargin = row[this.BottomMargin];
                cells.LeftMargin = row[this.LeftMargin];
                cells.RightMargin = row[this.RightMargin];
                cells.TopMargin = row[this.TopMargin];
                cells.DefaultTabStop = row[this.DefaultTabStop];
                cells.TextBkgnd = row[this.TextBkgnd].ToInt();
                cells.TextBkgndTrans = row[this.TextBkgndTrans];
                cells.TextDirection = row[this.TextDirection].ToInt();
                cells.VerticalAlign = row[this.VerticalAlign].ToInt();
                cells.TxtPinX = row[this.TxtPinX];
                cells.TxtPinY = row[this.TxtPinY];
                cells.TxtLocPinX = row[this.TxtLocPinX];
                cells.TxtLocPinY = row[this.TxtLocPinY];
                cells.TxtWidth = row[this.TxtWidth];
                cells.TxtHeight = row[this.TxtHeight];
                cells.TxtAngle = row[this.TxtAngle];
                return cells;
            }
        }
    }
}