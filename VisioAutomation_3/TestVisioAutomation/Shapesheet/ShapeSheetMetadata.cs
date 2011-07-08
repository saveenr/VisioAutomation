using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Diagnostics;
using SEC = Microsoft.Office.Interop.Visio.VisSectionIndices;
using ROW = Microsoft.Office.Interop.Visio.VisRowIndices;

namespace TestVisioAutomation
{

    public class CellInfo
    {
        public string RealName;
        public VisioAutomation.ShapeSheet.SRC SRC;
        public string XName;
        public VisioAutomation.ShapeSheet.SRC XSRC;
        public string Formula;
        public double Result;
    }
}
