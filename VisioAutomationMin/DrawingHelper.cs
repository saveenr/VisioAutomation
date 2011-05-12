using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using IVisio= Microsoft.Office.Interop.Visio;
using VAM=VisioAutomationMin;

namespace VisioAutomationMin
{
    public static class DrawingHelper
    {
        public static IVisio.Document OpenStencil(IVisio.Documents docs, string filename)
        {
            if (filename == null)
            {
                throw new System.ArgumentNullException("filename");
            }

            short flags = (short)IVisio.VisOpenSaveArgs.visOpenRO | (short)IVisio.VisOpenSaveArgs.visOpenDocked;
            var doc = docs.OpenEx(filename, flags);
            return doc;
        }

        public static short[] DropMaster(IVisio.Page page, IVisio.Master master, IList<Point> points)
        {
            var masters = new object[] { master };
            var xys = new double[points.Count];
            for (int i = 0; i < points.Count; i++)
            {
                xys[i + 0] = points[i].X;
                xys[i + 1] = points[i].Y;
            }

            System.Array out_ids_sa;
            page.DropManyU(masters, xys, out out_ids_sa);
            short[] out_ids = (short[])out_ids_sa;
            return out_ids;
        }

        public static short[] DropMaster(IVisio.Page page, IVisio.Master master, IList<Rectangle> rectangles)
        {
            var masters = new object[rectangles.Count];
            for (int i = 0; i < rectangles.Count; i++)
            {
                masters[i] = master;
            }

            var xys = new double[rectangles.Count*2];
            for (int i=0; i<rectangles.Count;i ++)
            {
                var r = rectangles[i];

                double x = (r.Right + r.Left)/2.0;
                double y = (r.Top + r.Bottom)/2.0;
                xys[(i*2) + 0] = x;
                xys[(i*2) + 1] = y;
            }
            System.Array out_ids_sa;
            page.DropManyU(masters, xys, out out_ids_sa);
            short[] out_ids = (short[])out_ids_sa;

            var update = new SIDSRCUpdate();

            for (int i = 0; i < rectangles.Count; i++ )
            {
                var rect = rectangles[i];
                update.SetFormula(out_ids[i], SRCConstants.Width, rect.Width);
                update.SetFormula(out_ids[i], SRCConstants.Height, rect.Height);
            }

            update.Execute(page, 0);
            
            return out_ids;
        }
    }
}
