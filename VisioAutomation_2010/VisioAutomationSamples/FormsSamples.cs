using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using System.Linq;
using System.Collections.Generic;
using GRIDMODEL = VisioAutomation.Models.Grid;

namespace VisioAutomationSamples
{
    public class R
    {
        public string name;
        public string ar;
        public int width;
        public int height;

        public R(string name, string ar, int width, int height)
        {
            this.name = name;
            this.ar = ar;
            this.width = width;
            this.height = height;
        }
    }
    public static class FormsSamples
    {
        public static void MonitorResolutions()
        {
            var x0=new R("VGA","4:03",640,480);
            var x1=new R("SVGA","4:03",800,600);
            var x2=new R("WSVGA","~17:10",1024,600);
            var x3=new R("XGA","4:03",1024,768);
            var x4=new R("XGA+","4:03",1152,864);
            var x5=new R("WXGA","16:09",1280,720);
            var x6=new R("WXGA","5:03",1280,768);
            var x7=new R("WXGA","16:10",1280,800);
            var x8=new R("SXGA–(UVGA)","4:03",1280,960);
            var x9=new R("SXGA","5:04",1280,1024);
            var x10=new R("HD","~16:9",1360,768);
            var x11=new R("HD","~16:9",1366,768);
            var x12=new R("SXGA+","4:03",1400,1050);
            var x13=new R("WXGA+","16:10",1440,900);
            var x14=new R("HD+","16:09",1600,900);
            var x15=new R("UXGA","4:03",1600,1200);
            var x16=new R("WSXGA+","16:10",1680,1050);
            var x17=new R("FHD","16:09",1920,1080);
            var x18=new R("WUXGA","16:10",1920,1200);
            var x19=new R("QWXGA","16:09",2048,1152);
            var x20=new R("WQHD","16:09",2560,1440);
            var x21=new R("WQXGA","16:10",2560,1600);
            var x22=new R("Unknown","3:04",768,1024);
            var x23=new R("Unknown","16:09",1093,614);
            var x24=new R("Unknown","~16:9",1311,737);

            var reses = new List<R>
            {
                x0,
                x1,
                x2,
                x3,
                x4,
                x5,
                x6,
                x7,
                x8,
                x9,
                x10,
                x11,
                x12,
                x13,
                x14,
                x15,
                x16,
                x17,
                x18,
                x19,
                x20,
                x21,
                x22,
                x23,
                x24
            };

            var doc = SampleEnvironment.Application.ActiveDocument;
            var ir = new VA.Models.Forms.InteractiveRenderer(doc);

            var formpage = new VA.Models.Forms.FormPage();
            var page = ir.CreatePage(formpage);

            ir.AddShape(5, 0.5, "Resolutions");
            ir.Linefeed(0.5);

            foreach (var r in  reses)
            {
                double w = r.width/400.0;
                double h = r.height /400.0;
                ir.AddShape(w, h, r.name);               
                ir.Linefeed(0.5);
            }

            page.ResizeToFitContents();
        }
    }
}


























