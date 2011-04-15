using System;
using System.Collections.Generic;
using System.Linq;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioAutomation.Layout
{
    public partial class XFormCells
    {
        public VA.Drawing.Point Pin
        {
            get { return new VA.Drawing.Point(this.PinX.Result, this.PinY.Result); }
            set
            {
                this.PinX.SetResult(value.X);
                this.PinY.SetResult(value.Y);
            }
        }

        public VA.Drawing.Point LocPin
        {
            get { return new VA.Drawing.Point(this.LocPinX.Result, this.LocPinY.Result); }
            set
            {
                this.LocPinX.SetResult(value.X);
                this.LocPinY.SetResult(value.Y);
            }
        }

        public VA.Drawing.Size Size
        {
            get { return new VA.Drawing.Size(this.Width.Result, this.Height.Result); }
            set
            {
                this.Width.SetResult(value.Width);
                this.Height.SetResult(value.Height);
            }
        }

        public VA.Drawing.Rectangle Rectangle
        {
            get
            {
                var left = this.PinX.Result - this.LocPinX.Result;
                var bottom = this.PinY.Result - this.LocPinY.Result;
                var lowerleft = new VA.Drawing.Point(left, bottom);
                return new Drawing.Rectangle(lowerleft, this.Size);
            }
        }

        public override string ToString()
        {
            string s = string.Format("({0}, {1}, {2}, {3})", this.Pin, this.LocPin, this.Size, this.Angle);
            return s;
        }

        public VA.Drawing.Rectangle Rect
        {
            get { return new VA.Drawing.Rectangle(this.Pin - this.LocPin, this.Size); }
        }
    }
}