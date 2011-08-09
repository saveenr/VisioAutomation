using System;
using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Drawing;
using VA=VisioAutomation;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Infographics
{


    public class RenderContext
    {
        public IVisio.Page Page;
        public VA.Drawing.Point CurrentUpperLeft;
    }
}
