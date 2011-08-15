using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Visio;
using IVisio = Microsoft.Office.Interop.Visio;
using VAM = VisioAutomationMin;

namespace ContainerLayout
{
    class ContainerUtil
    {
        public static IVisio.Document LoadContainerStencil(IVisio.Documents docs)
        {
            // load the special container stencil
            var app = docs.Application;
            var measurement = IVisio.VisMeasurementSystem.visMSUS;
            var stenciltype = IVisio.VisBuiltInStencilTypes.visBuiltInStencilContainers;
            string stencilfile = app.GetBuiltInStencilFile(stenciltype, measurement);
            short flags = (short)IVisio.VisOpenSaveArgs.visOpenHidden;
            var stencil = docs.OpenEx(stencilfile, flags);
            return stencil;
        }
    }
}
