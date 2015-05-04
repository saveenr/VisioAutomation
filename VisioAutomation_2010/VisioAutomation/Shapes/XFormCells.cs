using System.Collections.Generic;
using VisioAutomation.ShapeSheet.Query;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public class XFormCells : ShapeSheet.CellGroups.CellGroup
    {
        public ShapeSheet.CellData<double> PinX { get; set; }
        public ShapeSheet.CellData<double> PinY { get; set; }
        public ShapeSheet.CellData<double> LocPinX { get; set; }
        public ShapeSheet.CellData<double> LocPinY { get; set; }
        public ShapeSheet.CellData<double> Width { get; set; }
        public ShapeSheet.CellData<double> Height { get; set; }
        public ShapeSheet.CellData<double> Angle { get; set; }

        public override IEnumerable<SRCFormulaPair> Pairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SRCConstants.PinX, this.PinX.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.PinY, this.PinY.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LocPinX, this.LocPinX.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LocPinY, this.LocPinY.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Width, this.Width.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Height, this.Height.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Angle, this.Angle.Formula);
            }
        }

        public static IList<XFormCells> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = XFormCells.get_query();
            return CellGroup._GetCells<XFormCells, double>(page, shapeids, query, query.GetCells);
        }

        public static XFormCells GetCells(IVisio.Shape shape)
        {
            var query = XFormCells.get_query();
            return CellGroup._GetCells<XFormCells, double>(shape, query, query.GetCells);
        }

        private static XFormCellQuery _mCellQuery;
        private static XFormCellQuery get_query()
        {
            XFormCells._mCellQuery = XFormCells._mCellQuery ?? new XFormCellQuery();
            return XFormCells._mCellQuery;
        }

        class XFormCellQuery : CellQuery
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
                this.PinX = this.AddCell(ShapeSheet.SRCConstants.PinX, "PinX");
                this.PinY = this.AddCell(ShapeSheet.SRCConstants.PinY, "PinY");
                this.LocPinX = this.AddCell(ShapeSheet.SRCConstants.LocPinX, "LocPinX");
                this.LocPinY = this.AddCell(ShapeSheet.SRCConstants.LocPinY, "LocPinY");
                this.Width = this.AddCell(ShapeSheet.SRCConstants.Width, "Width");
                this.Height = this.AddCell(ShapeSheet.SRCConstants.Height, "Height");
                this.Angle = this.AddCell(ShapeSheet.SRCConstants.Angle, "Angle");
            }

            public XFormCells GetCells(IList<ShapeSheet.CellData<double>> row)
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