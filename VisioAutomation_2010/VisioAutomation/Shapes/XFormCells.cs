using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public class XFormCells : ShapeSheet.CellGroups.CellGroupSingleRow
    {
        public ShapeSheet.CellData PinX { get; set; }
        public ShapeSheet.CellData PinY { get; set; }
        public ShapeSheet.CellData LocPinX { get; set; }
        public ShapeSheet.CellData LocPinY { get; set; }
        public ShapeSheet.CellData Width { get; set; }
        public ShapeSheet.CellData Height { get; set; }
        public ShapeSheet.CellData Angle { get; set; }

        public override IEnumerable<SrcFormulaPair> SrcFormulaPairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SrcConstants.XFormPinX, this.PinX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.XFormPinY, this.PinY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.XFormLocPinX, this.LocPinX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.XFormLocPinY, this.LocPinY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.XFormWidth, this.Width.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.XFormHeight, this.Height.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.XFormAngle, this.Angle.Formula);
            }
        }

        public static List<XFormCells> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = XFormCells.lazy_query.Value;
            return query.GetCellGroups(page, shapeids);
        }

        public static XFormCells GetCells(IVisio.Shape shape)
        {
            var query = XFormCells.lazy_query.Value;
            return query.GetCellGroup(shape);
        }

        private static readonly System.Lazy<XFormCellsReader> lazy_query = new System.Lazy<XFormCellsReader>();
    }
}