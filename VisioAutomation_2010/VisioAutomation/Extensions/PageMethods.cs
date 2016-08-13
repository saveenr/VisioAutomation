using System.Collections.Generic;
using Microsoft.Office.Interop.Visio;
using VisioAutomation.Drawing;

namespace VisioAutomation.Extensions
{
    public static class PageMethods
    {
        public static void ResizeToFitContents(this Microsoft.Office.Interop.Visio.Page page, Drawing.Size padding)
        {
            Pages.PageHelper.ResizeToFitContents(page, padding);
        }

        public static Microsoft.Office.Interop.Visio.Shape DrawLine(this Microsoft.Office.Interop.Visio.Page page, Drawing.Point p1, Drawing.Point p2)
        {
            return VisioAutomation.Pages.PageHelper.DrawLine(page, p1, p2);
        }

        public static Microsoft.Office.Interop.Visio.Shape DrawOval(this Microsoft.Office.Interop.Visio.Page page, Drawing.Rectangle rect)
        {
            return VisioAutomation.Pages.PageHelper.DrawOval(page, rect);
        }

        public static Microsoft.Office.Interop.Visio.Shape DrawRectangle(this Microsoft.Office.Interop.Visio.Page page, Drawing.Rectangle rect)
        {
            return VisioAutomation.Pages.PageHelper.DrawRectangle(page, rect);
        }

        public static Microsoft.Office.Interop.Visio.Shape DrawBezier(this Microsoft.Office.Interop.Visio.Page page, IList<Drawing.Point> points)
        {
            return VisioAutomation.Pages.PageHelper.DrawBezier(page, points);
        }

        public static Microsoft.Office.Interop.Visio.Shape DrawBezier(this Microsoft.Office.Interop.Visio.Page page, IList<Drawing.Point> points, short degree, short flags)
        {
            return VisioAutomation.Pages.PageHelper.DrawBezier(page, points, degree, flags);
        }

        public static Microsoft.Office.Interop.Visio.Shape DrawPolyline(this Microsoft.Office.Interop.Visio.Page page, IList<Drawing.Point> points)
        {
            return VisioAutomation.Pages.PageHelper.DrawPolyline(page, points);
        }

        public static Microsoft.Office.Interop.Visio.Shape DrawNURBS(this Microsoft.Office.Interop.Visio.Page page, IList<Drawing.Point> controlpoints,
                                             IList<double> knots,
                                             IList<double> weights, int degree)
        {
            return VisioAutomation.Pages.PageHelper.DrawNURBS(page, controlpoints, knots, weights, degree);
        }

        public static Microsoft.Office.Interop.Visio.Shape Drop(
            this Microsoft.Office.Interop.Visio.Page page,
            Microsoft.Office.Interop.Visio.Master master,
            Drawing.Point point)
        {
            return VisioAutomation.Pages.PageHelper.Drop(page, master, point);
        }

        public static short[] DropManyU(
            this Microsoft.Office.Interop.Visio.Page page,
            IList<Master> masters,
            IEnumerable<Drawing.Point> points)
        {
            // TODO: Put this method in pagehelper class
            var surface = new Drawing.DrawingSurface(page);
            short[] shapeids = surface.DropManyU(masters, points);
            return shapeids;
        }

   	    public static short[] DropManyU(this Microsoft.Office.Interop.Visio.Page page, IList<Master> masters, IEnumerable<Point> points, IList<string> names)
   	    {
   	        return VisioAutomation.Pages.PageHelper.DropManyU(page, masters, points, names);
        }

        public static IEnumerable<Page> ToEnumerable(this Microsoft.Office.Interop.Visio.Pages pages)
        {
            return VisioAutomation.Pages.PageHelper.ToEnumerable(pages);
        }

        public static string[] GetNamesU(this Microsoft.Office.Interop.Visio.Pages pages)
        {
            return VisioAutomation.Pages.PageHelper.GetNamesU(pages);
        }
    }
}