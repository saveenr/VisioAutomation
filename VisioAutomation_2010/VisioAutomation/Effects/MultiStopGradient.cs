using System;
using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Linq;
using VisioAutomation.Extensions;

namespace VisioAutomation.Effects
{
    public class MultiStopGradient
    {
        public IList<GradientStop> Stops { get; private set; }
        public VA.Effects.MultiStopGradientDirection Direction { get; set; }

        public MultiStopGradient()
        {
            this.Stops = new List<GradientStop>();
            this.Direction = VA.Effects.MultiStopGradientDirection.LeftToRight;
        }

        public void Add(VA.Drawing.ColorRGB color, double trans, double pos)
        {
            if (pos < 0)
            {
                throw new ArgumentOutOfRangeException("pos");
            }

            if (pos > 1)
            {
                throw new ArgumentOutOfRangeException("pos");
            }

            var stop = new GradientStop(color, trans, pos);
            this.Stops.Add(stop);
        }

        public IVisio.Shape Draw( IVisio.Page page, VA.Drawing.Rectangle rect)
        {
            var gradient = this;

            var stops = gradient.Stops;

            if (stops.Count < 2)
            {
                throw new System.ArgumentException("gradient must have at least two stops");
            }

            var first_stop = stops[0];
            if (first_stop.Position != 0)
            {
                throw new System.ArgumentException("First stop must be at 0.0");
            }

            var last_stop = stops[stops.Count - 1];
            if (last_stop.Position != 1.0)
            {
                throw new System.ArgumentException("Last stop must be at 1.0");
            }

            int num_pairs = stops.Count - 1;

            // Check that they are increasing
            foreach (int i in Enumerable.Range(0, num_pairs))
            {
                var prev_stop = stops[i];
                var next_stop = stops[i + 1];

                if (next_stop.Position <= prev_stop.Position)
                {
                    throw new System.ArgumentException("Stop positions must monotonically increase from 0.0 to 1.0");
                }
            }

            double cur_pos;
            double physical_length;

            VA.Format.FillPattern fillpat;
            if (gradient.Direction == VA.Effects.MultiStopGradientDirection.BottomToTop)
            {
                physical_length = rect.Height;
                cur_pos = rect.Bottom;
                fillpat = VA.Format.FillPattern.LinearBottomToTop;
            }
            else
            {
                physical_length = rect.Width;
                cur_pos = rect.Left;
                fillpat = VA.Format.FillPattern.LinearLeftToRight;
            }

            double scale = physical_length;

            var grad_shapes = new IVisio.Shape[num_pairs];

            // Draw the shapes
            foreach (int i in Enumerable.Range(0, num_pairs))
            {
                var prev_stop = stops[i];
                var next_stop = stops[i + 1];

                double cur_length = (next_stop.Position - prev_stop.Position) * scale;

                VA.Drawing.Rectangle cur_rect;

                if (gradient.Direction == VA.Effects.MultiStopGradientDirection.BottomToTop)
                {
                    cur_rect = new VA.Drawing.Rectangle(rect.Left, cur_pos, rect.Right,
                                                        (cur_pos + cur_length));
                }
                else
                {
                    cur_rect = new VA.Drawing.Rectangle(cur_pos, rect.Bottom, (cur_pos + cur_length),
                                                        rect.Top);
                }

                var cur_shape = page.DrawRectangle(cur_rect);
                grad_shapes[i] = cur_shape;

                cur_pos += cur_length;
            }

            // Format the shapes
            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();
            foreach (int i in Enumerable.Range(0, num_pairs))
            {
                var shape = grad_shapes[i];
                short shapeid = (short)shape.ID;

                var prev_stop = stops[i];
                var next_stop = stops[i + 1];

                int linepat = 0;

                var gf = new VA.Effects.GradientFillDefinition();
                gf.StartColor = prev_stop.Color.ToFormula();
                gf.EndColor = next_stop.Color.ToFormula();
                gf.StartTransparency = prev_stop.Transparency.Value;
                gf.EndTransparency = next_stop.Transparency.Value;
                gf.FillPattern = (int)fillpat;

                gf.Apply(update, shapeid);
                update.SetFormula(shapeid, VA.ShapeSheet.SRCConstants.LinePattern, linepat);
            }
            update.Execute(page);

            var application = page.Application;
            var active_window = application.ActiveWindow;
            active_window.DeselectAll();
            var group = VA.Selection.SelectionHelper.SelectAndGroup(active_window, grad_shapes);
            VA.ShapeHelper.SetGroupSelectMode(group, VA.Selection.GroupSelectMode.GroupOnly);

            return group;
        }

    }
}