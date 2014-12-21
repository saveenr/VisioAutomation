using System.Collections.Generic;
using System.Linq;
using VisioAutomation.ShapeSheet.CellGroups;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioAutomation.Shapes
{
    public class XFormCells : VA.ShapeSheet.CellGroups.CellGroup
    {
        public VA.ShapeSheet.CellData<double> PinX { get; set; }
        public VA.ShapeSheet.CellData<double> PinY { get; set; }
        public VA.ShapeSheet.CellData<double> LocPinX { get; set; }
        public VA.ShapeSheet.CellData<double> LocPinY { get; set; }
        public VA.ShapeSheet.CellData<double> Width { get; set; }
        public VA.ShapeSheet.CellData<double> Height { get; set; }
        public VA.ShapeSheet.CellData<double> Angle { get; set; }

        public override IEnumerable<VA.ShapeSheet.CellGroups.BaseCellGroup.SRCValuePair> EnumPairs()
        {
            yield return srcvaluepair(ShapeSheet.SRCConstants.PinX, this.PinX.Formula);
            yield return srcvaluepair(ShapeSheet.SRCConstants.PinY, this.PinY.Formula);
            yield return srcvaluepair(ShapeSheet.SRCConstants.LocPinX, this.LocPinX.Formula);
            yield return srcvaluepair(ShapeSheet.SRCConstants.LocPinY, this.LocPinY.Formula);
            yield return srcvaluepair(ShapeSheet.SRCConstants.Width, this.Width.Formula);
            yield return srcvaluepair(ShapeSheet.SRCConstants.Height, this.Height.Formula);
            yield return srcvaluepair(ShapeSheet.SRCConstants.Angle, this.Angle.Formula);
        }

        public static IList<XFormCells> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = get_query();
            return VA.ShapeSheet.CellGroups.CellGroup._GetCells<XFormCells, double>(page, shapeids, query, query.GetCells);
        }

        public static XFormCells GetCells(IVisio.Shape shape)
        {
            var query = get_query();
            return VA.ShapeSheet.CellGroups.CellGroup._GetCells<XFormCells, double>(shape, query, query.GetCells);
        }

        private static XFormCellQuery _mCellQuery;
        private static XFormCellQuery get_query()
        {
            _mCellQuery = _mCellQuery ?? new XFormCellQuery();
            return _mCellQuery;
        }

        class XFormCellQuery : VA.ShapeSheet.Query.CellQuery
        {
            public Column Width { get; set; }
            public Column Height { get; set; }
            public Column PinX { get; set; }
            public Column PinY { get; set; }
            public Column LocPinX { get; set; }
            public Column LocPinY { get; set; }
            public Column Angle { get; set; }

            public XFormCellQuery()
            {
                PinX = this.Columns.Add(VA.ShapeSheet.SRCConstants.PinX, "PinX");
                PinY = this.Columns.Add(VA.ShapeSheet.SRCConstants.PinY, "PinY");
                LocPinX = this.Columns.Add(VA.ShapeSheet.SRCConstants.LocPinX, "LocPinX");
                LocPinY = this.Columns.Add(VA.ShapeSheet.SRCConstants.LocPinY, "LocPinY");
                Width = this.Columns.Add(VA.ShapeSheet.SRCConstants.Width, "Width");
                Height = this.Columns.Add(VA.ShapeSheet.SRCConstants.Height, "Height");
                Angle = this.Columns.Add(VA.ShapeSheet.SRCConstants.Angle, "Angle");
            }

            public XFormCells GetCells(VA.ShapeSheet.CellData<double>[] row)
            {
                var cells = new XFormCells
                {
                    PinX = row[this.PinX.Ordinal],
                    PinY = row[this.PinY.Ordinal],
                    LocPinX = row[this.LocPinX.Ordinal],
                    LocPinY = row[this.LocPinY.Ordinal],
                    Width = row[this.Width.Ordinal],
                    Height = row[this.Height.Ordinal],
                    Angle = row[this.Angle.Ordinal]
                };
                return cells;
            }
        }
    }
}