using System.Collections.Generic;
using System.Linq;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioAutomation.Layout
{
    public class XFormCells : VA.ShapeSheet.CellGroups.CellGroupEx
    {
        public VA.ShapeSheet.CellData<double> PinX { get; set; }
        public VA.ShapeSheet.CellData<double> PinY { get; set; }
        public VA.ShapeSheet.CellData<double> LocPinX { get; set; }
        public VA.ShapeSheet.CellData<double> LocPinY { get; set; }
        public VA.ShapeSheet.CellData<double> Width { get; set; }
        public VA.ShapeSheet.CellData<double> Height { get; set; }
        public VA.ShapeSheet.CellData<double> Angle { get; set; }


        public override void ApplyFormulas(ApplyFormula func)
        {
            func(ShapeSheet.SRCConstants.PinX, this.PinX.Formula);
            func(ShapeSheet.SRCConstants.PinY, this.PinY.Formula);
            func(ShapeSheet.SRCConstants.LocPinX, this.LocPinX.Formula);
            func(ShapeSheet.SRCConstants.LocPinY, this.LocPinY.Formula);
            func(ShapeSheet.SRCConstants.Width, this.Width.Formula);
            func(ShapeSheet.SRCConstants.Height, this.Height.Formula);
            func(ShapeSheet.SRCConstants.Angle, this.Angle.Formula);
        }

        public static IList<XFormCells> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = get_query();
            return VA.ShapeSheet.CellGroups.CellGroupEx._GetCells(page, shapeids, query, query.GetCells);
        }

        public static XFormCells GetCells(IVisio.Shape shape)
        {
            var query = get_query();
            return VA.ShapeSheet.CellGroups.CellGroupEx._GetCells(shape, query, query.GetCells);
        }

        private static XFormQuery m_query;
        private static XFormQuery get_query()
        {
            m_query = m_query ?? new XFormQuery();
            return m_query;
        }

        class XFormQuery : VA.ShapeSheet.Query.QueryEx
        {
            public int Width { get; set; }
            public int Height { get; set; }
            public int PinX { get; set; }
            public int PinY { get; set; }
            public int LocPinX { get; set; }
            public int LocPinY { get; set; }
            public int Angle { get; set; }

            public XFormQuery()
            {
                PinX = this.AddCell(VA.ShapeSheet.SRCConstants.PinX, "PinX");
                PinY = this.AddCell(VA.ShapeSheet.SRCConstants.PinY, "PinY");
                LocPinX = this.AddCell(VA.ShapeSheet.SRCConstants.LocPinX, "LocPinX");
                LocPinY = this.AddCell(VA.ShapeSheet.SRCConstants.LocPinY, "LocPinY");
                Width = this.AddCell(VA.ShapeSheet.SRCConstants.Width, "Width");
                Height = this.AddCell(VA.ShapeSheet.SRCConstants.Height, "Height");
                Angle = this.AddCell(VA.ShapeSheet.SRCConstants.Angle, "Angle");
            }

            public  XFormCells GetCells(ExQueryResult<CellData<double>> data_for_shape)
            {
                var row = data_for_shape.Cells;

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