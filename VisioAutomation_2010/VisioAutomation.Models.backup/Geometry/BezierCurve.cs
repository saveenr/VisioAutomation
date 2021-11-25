﻿using System.Linq;

namespace VisioAutomation.Models.Geometry
{
    public class BezierCurve
    {
        public VisioAutomation.Geometry.Point[] ControlPoints { get; }
        public int Degree { get; }

        public BezierCurve(VisioAutomation.Geometry.Point[] controlpoints, int degree)
        {
            if (degree < 1)
            {
                throw new System.ArgumentOutOfRangeException(nameof(degree));                
            }

            this.ControlPoints = controlpoints ?? throw new System.ArgumentNullException(nameof(controlpoints));
            this.Degree = degree;
        }

        public static BezierCurve FromEllipse(VisioAutomation.Geometry.Point center, VisioAutomation.Geometry.Size radius)
        {
            var pt1 = new VisioAutomation.Geometry.Point(0, radius.Height); // top
            var pt2 = new VisioAutomation.Geometry.Point(radius.Width, 0); // right
            var pt3 = new VisioAutomation.Geometry.Point(0, -radius.Height); // bottom
            var pt4 = new VisioAutomation.Geometry.Point(-radius.Width, 0); // left

            double dx = radius.Width * 4.0 * (System.Math.Sqrt(2) - 1) / 3;
            double dy = radius.Height * 4.0 * (System.Math.Sqrt(2) - 1) / 3;

            var curve_control_points = new []
                                      {
                                          pt1,
                                          pt1.Add(dx, 0),
                                          pt2.Add(0, dy),
                                          pt2,
                                          pt2.Add(0, -dy),
                                          pt3.Add(dx, 0),
                                          pt3,
                                          pt3.Add(-dx, 0),
                                          pt4.Add(0, -dy),
                                          pt4,
                                          pt4.Add(0, dy),
                                          pt1.Add(-dx, 0),
                                          pt1
                                      }
                .Select(p => p + center).ToArray();
            var curve_Degree = 3;
            
            var curve = new BezierCurve(curve_control_points, curve_Degree);
            return curve;
        }
    }
}