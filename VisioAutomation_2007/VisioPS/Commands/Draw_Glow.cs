using VisioPS.Extensions;
using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    public enum EdgeGlowDirection
    {
        Inner,
        Outer
    }

    [SMA.Cmdlet("Draw", "Glow")]
    public class Draw_Glow : VisioPS.VisioPSCmdlet
    {

        [SMA.Parameter(Position = 0, Mandatory = true)]
        public double X0 { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = true)]
        public double Y0 { get; set; }

        [SMA.Parameter(Position = 2, Mandatory = true)]
        public double X1 { get; set; }

        [SMA.Parameter(Position = 3, Mandatory = true)]
        public double Y1 { get; set; }

        [SMA.Parameter(Mandatory = true)]
        public string Color { get; set; }

        [SMA.Parameter(Mandatory = true)]
        public double Transparency { get; set; }

        [SMA.Parameter(Mandatory = true)]
        public double Width { get; set; }

        [SMA.Parameter(Mandatory = true)]
        public EdgeGlowDirection Direction { get; set; }

        protected override void ProcessRecord()
        {
            var glow = new VA.Effects.EdgeGlow();
            var rect = new VA.Drawing.Rectangle(this.X0, this.Y0, this.X1, this.Y1);
            glow.GlowColor = VA.Drawing.ColorRGB.ParseWebColor(this.Color);
            glow.GlowTransparency = this.Transparency;
            glow.GlowWidth = this.Width;

            var scriptingsession = this.ScriptingSession;
            if (this.Direction == EdgeGlowDirection.Inner)
            {
                glow.DrawInner(scriptingsession.VisioApplication.ActivePage, rect);
            }
            else
            {
                glow.DrawOuter(scriptingsession.VisioApplication.ActivePage, rect);
                
            }
        }
    }
}