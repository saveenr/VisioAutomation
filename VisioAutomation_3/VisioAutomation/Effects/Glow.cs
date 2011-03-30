using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Effects
{
    public class Glow
    {
        public double GlowWidth { get; set; }
        public VA.Drawing.ColorRGB GlowColor { get; set; }
        public int GlowTransparency { get; set; }

        private static Drawing.Rectangle[] GetInnerBorderRects(VA.Drawing.Rectangle R, double w)
        {
            VA.Drawing.Rectangle[] rects = {
                                    new VA.Drawing.Rectangle(R.LowerLeft.Add(0, 0), R.LowerLeft.Add(w, w)),
                                    new VA.Drawing.Rectangle(R.LowerRight.Add(-w, 0), R.LowerRight.Add(0, w)),
                                    new VA.Drawing.Rectangle(R.UpperLeft.Add(0, -w), R.UpperLeft.Add(w, 0)),
                                    new VA.Drawing.Rectangle(R.UpperRight.Add(-w, -w), R.UpperRight.Add(0, 0)),
                                    new VA.Drawing.Rectangle(R.LowerLeft.Add(0, w), R.UpperLeft.Add(w, -w)),
                                    new VA.Drawing.Rectangle(R.LowerRight.Add(-w, w), R.UpperRight.Add(0, -w)),
                                    new VA.Drawing.Rectangle(R.UpperLeft.Add(w, -w), R.UpperRight.Add(-w, 0)),
                                    new VA.Drawing.Rectangle(R.LowerLeft.Add(w, 0), R.LowerRight.Add(-w, w))
                                };
            return rects;
        }



        public IVisio.Shape Draw(IVisio.Page page, VA.Drawing.Rectangle rect)
        {
            var glow = this;
            double bg_trans = 1.0;
            var bg_color = glow.GlowColor;

            var rects = GetInnerBorderRects(rect, glow.GlowWidth);

            VA.Format.FillPattern[] grads = {
                                      VA.Format.FillPattern.RadialUpperRight,
                                      VA.Format.FillPattern.RadialUpperLeft,
                                      VA.Format.FillPattern.RadialLowerRight,
                                      VA.Format.FillPattern.RadialLowerLeft,
                                      VA.Format.FillPattern.LinearRightToLeft,
                                      VA.Format.FillPattern.LinearLeftToRight,
                                      VA.Format.FillPattern.LinearBottomToTop,
                                      VA.Format.FillPattern.LinearTopToBottom
                                  };

            var shapes = (from r in rects
                          select page.DrawRectangle(r))
                .ToArray();


            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();

            for (int i = 0; i < shapes.Length; i++)
            {
                short shapeid = (short)shapes[i].ID;
                update.SetFormula(shapeid, VA.ShapeSheet.SRCConstants.FillPattern, (int)grads[i]);
                update.SetFormula(shapeid, VA.ShapeSheet.SRCConstants.FillForegnd, VA.Convert.ColorToFormulaRGB(glow.GlowColor));
                update.SetFormula(shapeid, VA.ShapeSheet.SRCConstants.FillBkgnd, VA.Convert.ColorToFormulaRGB(bg_color));
                update.SetFormula(shapeid, VA.ShapeSheet.SRCConstants.FillForegndTrans, bg_trans);
                update.SetFormula(shapeid, VA.ShapeSheet.SRCConstants.FillBkgndTrans, glow.GlowTransparency);
                update.SetFormula(shapeid, VA.ShapeSheet.SRCConstants.LinePattern, 0);
            }

            update.Execute(page);

            var application = page.Application;
            var active_window = application.ActiveWindow;
            active_window.DeselectAll();
            var group = VA.SelectionHelper.SelectAndGroup(active_window, shapes);
            VA.ShapeHelper.SetGroupSelectMode(group, IVisio.VisCellVals.visGrpSelModeGroupOnly);

            return group;
        }

        public IVisio.Shape DrawOuter(IVisio.Page page, VA.Drawing.Rectangle rect)
        {
            var glow = this;
            const int bg_trans = 1;
            var bg_color = glow.GlowColor;

            var rects = GetOuterBorderRects(rect, glow.GlowWidth);

            VA.Format.FillPattern[] grads = {
                                      VA.Format.FillPattern.RadialUpperRight,
                                      VA.Format.FillPattern.RadialUpperLeft,
                                      VA.Format.FillPattern.RadialLowerRight,
                                      VA.Format.FillPattern.RadialLowerLeft,
                                      VA.Format.FillPattern.LinearRightToLeft,
                                      VA.Format.FillPattern.LinearLeftToRight,
                                      VA.Format.FillPattern.LinearBottomToTop,
                                      VA.Format.FillPattern.LinearTopToBottom
                                  };

            var shapes = (from r in rects
                          select page.DrawRectangle(r))
                .ToArray();

            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();


            for (int i = 0; i < shapes.Length; i++)
            {
                int shapeid = shapes[i].ID;

                update.SetFormula((short)shapeid, VA.ShapeSheet.SRCConstants.FillPattern, (int)grads[i]);
                update.SetFormula((short)shapeid, VA.ShapeSheet.SRCConstants.FillForegnd, VA.Convert.ColorToFormulaRGB(glow.GlowColor));
                update.SetFormula((short)shapeid, VA.ShapeSheet.SRCConstants.FillBkgnd, VA.Convert.ColorToFormulaRGB(bg_color));
                update.SetFormula((short)shapeid, VA.ShapeSheet.SRCConstants.FillForegndTrans, glow.GlowTransparency);
                update.SetFormula((short)shapeid, VA.ShapeSheet.SRCConstants.FillBkgndTrans, bg_trans);
                update.SetFormula((short)shapeid, VA.ShapeSheet.SRCConstants.LinePattern, 0);
            }

            update.Execute(page);

            var application = page.Application;
            var active_window = application.ActiveWindow;
            active_window.DeselectAll();
            var group = VA.SelectionHelper.SelectAndGroup(active_window, shapes);
            VA.ShapeHelper.SetGroupSelectMode(group, IVisio.VisCellVals.visGrpSelModeGroupOnly);

            return group;
        }

        private static VA.Drawing.Rectangle[] GetOuterBorderRects(VA.Drawing.Rectangle R, double w)
        {
            VA.Drawing.Rectangle[] rects = {
                                    new VA.Drawing.Rectangle(R.LowerLeft.Add(-w, -w), R.LowerLeft.Add(0, 0)),
                                    new VA.Drawing.Rectangle(R.LowerRight.Add(0, -w), R.LowerRight.Add(w, 0)),
                                    new VA.Drawing.Rectangle(R.UpperLeft.Add(-w, 0), R.UpperLeft.Add(0, w)),
                                    new VA.Drawing.Rectangle(R.UpperRight.Add(0, 0), R.UpperRight.Add(w, w)),
                                    new VA.Drawing.Rectangle(R.LowerLeft.Add(-w, 0), R.UpperLeft.Add(0, 0)),
                                    new VA.Drawing.Rectangle(R.LowerRight.Add(0, 0), R.UpperRight.Add(w, 0)),
                                    new VA.Drawing.Rectangle(R.UpperLeft.Add(0, 0), R.UpperRight.Add(0, w)),
                                    new VA.Drawing.Rectangle(R.LowerLeft.Add(0, -w), R.LowerRight.Add(0, 0))
                                };
            return rects;
        }


    }
}