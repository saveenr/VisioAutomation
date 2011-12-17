using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using IVisio = Microsoft.Office.Interop.Visio;

namespace GetMasterPreviews
{
    class Program
    {
        static void Main(string[] args)
        {
            var app = new IVisio.Application();
            var doc = app.Documents.Add("");
            var flags = IVisio.VisOpenSaveArgs.visOpenDocked | IVisio.VisOpenSaveArgs.visOpenRO;
            var stencil = app.Documents.OpenEx( "BASIC_U.VSS", (short)flags);
            var page = app.ActivePage;
            var master = stencil.Masters["Rectangle"];

            var pic = master.Picture;
            var image = Microsoft.VisualBasic.Compatibility.VB6.Support.IPictureDispToImage(pic);

            image.Save("D:\\x.png");
            page.Drop(master, 4, 5);
        }
    }
}
