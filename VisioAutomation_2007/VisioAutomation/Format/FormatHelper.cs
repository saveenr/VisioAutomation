using VA=VisioAutomation;
using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Format
{
    public static class FormatHelper
    {
        public static VA.Format.ShapeFormatCells GetShapeFormat(IVisio.Shape shape)
        {
            return ShapeFormatCells.GetCells(shape);
        }

        public static IList<VA.Format.ShapeFormatCells> GetShapeFormat(IVisio.Page page, IList<int> shapeids)
        {
            return ShapeFormatCells.GetCells(page, shapeids);
        }
    }
}