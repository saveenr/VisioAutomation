using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VisioAutomation.UI
{
    class ColorUtil
    {
        public static void Normalize24BitRGB(byte r, byte g, byte b, out double R, out double G, out double B)
        {
            R = ((double)r) / 255.0;
            G = ((double)g) / 255.0;
            B = ((double)b) / 255.0;
        }

        public static void DeNormalize24BitRGB(double R, double G, double B, out byte r, out byte g, out byte b)
        {
            r = (byte)(R * 255.0);
            g = (byte)(G * 255.0);
            b = (byte)(B * 255.0);
        }


        public static void RGBToHSV(double R, double G, double B, out double H, out double S, out double V)
        {
            /* * CREDITS
             * -------
             * The HSV<->RGB Conversion code based on this source code: http://www.cs.rit.edu/~ncs/color/t_convert.html
             * from Eugene Vishnevsky*/

            double _max = System.Math.Max(R, System.Math.Max(G, B));
            double _min = System.Math.Min(R, System.Math.Min(G, B));

            double the_h = 0.0;
            double the_s = 0.0;
            double the_v = _max;

            double delta = _max - _min;

            if (_max == 0.0)
            {
                // this means r=g=b=0
                the_s = 0;
                the_h = 0;
                the_v = 0;

                H = the_h;
                S = the_s;
                V = the_v;
            }
            else
            {
                the_s = delta / _max;

                if (delta == 0.0)
                {
                    the_h = 1.0;
                }
                else
                {
                    if (R == _max)
                    {
                        the_h = (G - B) / delta;
                    }
                    else if (G == _max)
                    {
                        the_h = 2.0 + (B - R) / delta;
                    }
                    else
                    {
                        the_h = 4.0 + (R - G) / delta;
                    }
                }
                the_h *= 60.0;
                if (the_h < 0)
                {
                    the_h += 360;
                }

                the_h /= 360.0; // scale hue to between 0.0 and 1.0
            }

            H = the_h;
            S = the_s;
            V = the_v;
        }

        public static void RGBToHSV(System.Drawing.Color rgb, out double H, out double S, out double V)
        {
            double r;
            double g;
            double b;

            Normalize24BitRGB(rgb.R, rgb.G, rgb.B, out r, out g, out b);

            RGBToHSV(r, g, b, out H, out S, out V);
        }


        public static void HSVToRGB(double H, double S, double V, out double R, out double G, out double B)
        {
            // CREDITS
            // Based on code by Eugene Vishnevsky found at:
            // http://www.cs.rit.edu/~ncs/color/t_convert.html
            // 
            // Returns the RGBColor version of this HSVColor

            if (S == 0.0)
            {
                // Make it some kind of gray
                R = V;
                G = V;
                B = V;
            }
            else
            {
                if (H >= 1.0)
                {
                    H = 0.0;
                }

                double step = 1.0 / 6.0;

                double vh = H / step;
                int vi = (int)System.Math.Floor(vh);
                double v0 = vh - vi;
                double v1 = V * (1.0 - S);
                double v2 = V * (1.0 - (S * v0));
                double v3 = V * (1.0 - (S * (1.0 - v0)));

                switch (vi)
                {
                    case 0:
                        {
                            R = V;
                            G = v3;
                            B = v1;
                            break;
                        }
                    case 1:
                        {
                            R = v2;
                            G = V;
                            B = v1;
                            break;
                        }
                    case 2:
                        {
                            R = v1;
                            G = V;
                            B = v3;
                            break;
                        }
                    case 3:
                        {
                            R = v1;
                            G = v2;
                            B = V;
                            break;
                        }
                    case 4:
                        {
                            R = v3;
                            G = v1;
                            B = V;
                            break;
                        }
                    default:
                        {
                            R = V;
                            G = v1;
                            B = v2;
                            break;
                        }
                }
            }
        }

        public static System.Drawing.Color HSVToSystemDrawingColor(double H, double S, double V)
        {
            double R;
            double G;
            double B;
            HSVToRGB(H, S, V, out R, out G, out B);
            byte r;
            byte g;
            byte b;
            DeNormalize24BitRGB(R, G, B, out r, out g, out b);
            return System.Drawing.Color.FromArgb(r, g, b);
        }


    }
}
