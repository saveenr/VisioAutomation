using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Effects
{
    public class OuterGlow
    {
        public VA.Drawing.Rectangle Rectangle { get; set; }
        public double GlowWidth { get; set; }
        public VA.Drawing.ColorRGB GlowColor { get; set; }
        public int GlowTransparency { get; set; }

        public IVisio.Shape Draw(IVisio.Page page)
        {
            var glow = this;
            const int bg_trans = 1;
            var bg_color = glow.GlowColor;

            var rects = GetOuterBorderRects(glow.Rectangle, glow.GlowWidth);

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