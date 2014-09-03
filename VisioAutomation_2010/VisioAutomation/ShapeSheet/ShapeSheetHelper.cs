using System;
using System.Collections.Generic;
using System.Linq;
using IVisio=Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.ShapeSheet
{
    public static class ShapeSheetHelper
    {
        public static string[] GetFormulasU(IVisio.Page page, short[] stream)
        {
            var surface = new VA.Drawing.DrawingSurface(page);
            return surface.GetFormulasU_4(stream);
        }

        public static string[] GetFormulasU(IVisio.Shape shape, short[] stream)
        {
            var surface = new VA.Drawing.DrawingSurface(shape);
            return surface.GetFormulasU_3(stream);
        }


        public static TResult[] GetResults<TResult>(IVisio.Page page, short[] stream, IList<IVisio.VisUnitCodes> unitcodes)
        {
            var surface = new VA.Drawing.DrawingSurface(page);
            return surface.GetResults_4<TResult>(stream,unitcodes);
        }

        public static TResult[] GetResults<TResult>(IVisio.Shape shape, short[] stream, IList<IVisio.VisUnitCodes> unitcodes)
        {
            var surface = new VA.Drawing.DrawingSurface(shape);
            return surface.GetResults_3<TResult>(stream, unitcodes);
        }
    }
}