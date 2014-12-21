using System.Text.RegularExpressions;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using System.Linq;
using System.Collections.Generic;
using GRIDMODEL = VisioAutomation.Models.Grid;

namespace VisioAutomationSamples
{
    public class ResolutionInfo
    {
        public string Name;
        public string AspectRatioName;
        public int Width;
        public int Height;
        public double AspectRatio;

        public ResolutionInfo(string name, string aspectrationame, int width, int height)
        {
            this.Name = name;
            this.AspectRatioName = aspectrationame;
            this.Width = width;
            this.Height = height;
            this.AspectRatio = this.Width *1.0 /this.Height;
        }
    }

    public static class FormsSamples
    {
        public static void MonitorResolutions()
        {
            var resolutions = new List<ResolutionInfo>
            {
                new ResolutionInfo("VGA","4:3",640,480),
                new ResolutionInfo("SVGA","4:3",800,600),
                new ResolutionInfo("WSVGA","~17:10",1024,600),
                new ResolutionInfo("XGA","4:3",1024,768),
                new ResolutionInfo("XGA+","4:3",1152,864),
                new ResolutionInfo("WXGA","16:9",1280,720),
                new ResolutionInfo("WXGA","5:3",1280,768),
                new ResolutionInfo("WXGA","16:10",1280,800),
                new ResolutionInfo("SXGA–(UVGA)","4:3",1280,960),
                new ResolutionInfo("SXGA","5:4",1280,1024),
                new ResolutionInfo("HD","~16:9",1360,768),
                new ResolutionInfo("HD","~16:9",1366,768),
                new ResolutionInfo("SXGA+","4:3",1400,1050),
                new ResolutionInfo("WXGA+","16:10",1440,900),
                new ResolutionInfo("HD+","16:9",1600,900),
                new ResolutionInfo("UXGA","4:3",1600,1200),
                new ResolutionInfo("WSXGA+","16:10",1680,1050),
                new ResolutionInfo("FHD","16:9",1920,1080),
                new ResolutionInfo("WUXGA","16:10",1920,1200),
                new ResolutionInfo("QWXGA","16:9",2048,1152),
                new ResolutionInfo("WQHD","16:9",2560,1440),
                new ResolutionInfo("WQXGA","16:10",2560,1600),
                new ResolutionInfo("Unknown","3:4",768,1024),
                new ResolutionInfo("Unknown","16:9",1093,614),
                new ResolutionInfo("Unknown","~16:9",1311,737)
            };

            var doc = SampleEnvironment.Application.ActiveDocument;

            var fonts = doc.Fonts;

            var segoe_ui_font = fonts["Segoe UI"];
            var segoe_ui__light_font = fonts["Segoe UI Light"];
            var segoe_ui_font_id = segoe_ui_font.ID;
            var segoe_ui__light_font_id = segoe_ui__light_font.ID;

            var renderer = new VA.Models.Forms.InteractiveRenderer(doc);

            var formpage = new VA.Models.Forms.FormPage();
            var page = renderer.CreatePage(formpage);

            double max_body_width = 30.0;
            var page_title = renderer.AddShape(max_body_width, 1.5, "Standard Resolutions by Aspect Ratio");
            page_title.CharacterCells.Font = segoe_ui__light_font_id;
            page_title.CharacterCells.Size = "100pt";
            page_title.ParagraphCells.HorizontalAlign = 0;
            page_title.FormatCells.LineWeight = 0;
            page_title.FormatCells.LinePattern = 0;
            //page_title.FormatCells.FillForegnd = "RGB(240,240,240)"; renderer.Linefeed(0.5);
            renderer.Linefeed(0);

            var grouped = resolutions.GroupBy(i => i.AspectRatioName).ToList();
            foreach (var group in grouped)
            {
                var group_title = renderer.AddShape(max_body_width, 1, group.Key);
               group_title.CharacterCells.Font = segoe_ui__light_font_id;
               group_title.CharacterCells.Size = "50pt";
               group_title.ParagraphCells.HorizontalAlign = 0;
               group_title.FormatCells.LineWeight = 0;
               group_title.FormatCells.LinePattern = 0;
               group_title.FormatCells.FillForegnd = "RGB(250,250,250)";

               renderer.Linefeed(0.5);

                foreach (var resolution in group)
                {
                    double w = resolution.Width / 400.0;
                    double h = resolution.Height / 400.0;

                    string label = string.Format("{0}\n{1}x{2}", resolution.Name, resolution.Width, resolution.Height);
                    var res_title = renderer.AddShape(w, h, label);
                    res_title.CharacterCells.Font = segoe_ui_font_id;
                    res_title.CharacterCells.Size = "25pt";
                    //res_title.ParagraphCells.HorizontalAlign = 0;
                    //res_title.FormatCells.LineWeight = 0;
                    //res_title.FormatCells.LinePattern = 0;
                    res_title.FormatCells.FillForegnd = "RGB(240,240,240)";
                    renderer.MoveRight(0.5);
                }
                renderer.Linefeed(1);
            }
            renderer.Finish();
            page.ResizeToFitContents();
        }
    }
}


























