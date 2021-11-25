using System.Linq;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands.VisioRectangle
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, Nouns.VisioRectangle)]
    public class NewVisioRectangle : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public float Left { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = true)]
        public float Bottom { get; set; }

        [SMA.Parameter(Position = 2, Mandatory = true)]
        public float Right { get; set; }

        [SMA.Parameter(Position = 3, Mandatory = true)]
        public float Top { get; set; }

        protected override void ProcessRecord()
        {
            var rect = new VisioAutomation.Geometry.Rectangle(this.Left, this.Bottom, this.Right, this.Top);
            this.WriteObject(rect);
        }
    }
}