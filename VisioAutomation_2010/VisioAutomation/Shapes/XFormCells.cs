using System.Collections.Generic;
using VisioAutomation.ShapeSheet.Query;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

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

        public override IEnumerable<SRCFormulaPair> Pairs
        {
            get
            {
                yield return newpair(ShapeSheet.SRCConstants.PinX, this.PinX.Formula);
                yield return newpair(ShapeSheet.SRCConstants.PinY, this.PinY.Formula);
                yield return newpair(ShapeSheet.SRCConstants.LocPinX, this.LocPinX.Formula);
                yield return newpair(ShapeSheet.SRCConstants.LocPinY, this.LocPinY.Formula);
                yield return newpair(ShapeSheet.SRCConstants.Width, this.Width.Formula);
                yield return newpair(ShapeSheet.SRCConstants.Height, this.Height.Formula);
                yield return newpair(ShapeSheet.SRCConstants.Angle, this.Angle.Formula);
            }
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
            public CellColumn Width { get; set; }
            public CellColumn Height { get; set; }
            public CellColumn PinX { get; set; }
            public CellColumn PinY { get; set; }
            public CellColumn LocPinX { get; set; }
            public CellColumn LocPinY { get; set; }
            public CellColumn Angle { get; set; }

            public XFormCellQuery()
            {
                PinX = this.AddCell(VA.ShapeSheet.SRCConstants.PinX, "PinX");
                PinY = this.AddCell(VA.ShapeSheet.SRCConstants.PinY, "PinY");
                LocPinX = this.AddCell(VA.ShapeSheet.SRCConstants.LocPinX, "LocPinX");
                LocPinY = this.AddCell(VA.ShapeSheet.SRCConstants.LocPinY, "LocPinY");
                Width = this.AddCell(VA.ShapeSheet.SRCConstants.Width, "Width");
                Height = this.AddCell(VA.ShapeSheet.SRCConstants.Height, "Height");
                Angle = this.AddCell(VA.ShapeSheet.SRCConstants.Angle, "Angle");
            }

            public XFormCells GetCells(IList<VA.ShapeSheet.CellData<double>> row)
            {
                var cells = new XFormCells
                {
                    PinX = row[this.PinX],
                    PinY = row[this.PinY],
                    LocPinX = row[this.LocPinX],
                    LocPinY = row[this.LocPinY],
                    Width = row[this.Width],
                    Height = row[this.Height],
                    Angle = row[this.Angle]
                };
                return cells;
            }
        }
    }
}