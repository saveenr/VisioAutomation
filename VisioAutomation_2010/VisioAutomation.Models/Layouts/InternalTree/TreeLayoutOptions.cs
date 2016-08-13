using VisioAutomation.Drawing;
using VisioAutomation.Drawing.Layout;

namespace VisioAutomation.Models.InternalTree
{
    internal class TreeLayoutOptions
    {
        public Drawing.Point TopAdjustment = new Point(0,0); // How to adjust the apex 

        public TreeLayoutOptions()
        {
            this.SubtreeSeparation = 1;
            this.SiblingSeparation = 1;
            this.Direction = LayoutDirection.Up;
            this.Alignment = AlignmentVertical.Top;
            this.MaximumDepth = 100;
            this.LevelSeparation = 1;
            this.DefaultNodeSize = new Drawing.Size(1, 1);
        }

        public Drawing.Size DefaultNodeSize { get; set; }
        public double LevelSeparation { get; set; }
        public int MaximumDepth { get; set; }
        public AlignmentVertical Alignment { get; set; }
        public LayoutDirection Direction { get; set; }
        public double SiblingSeparation { get; set; }
        public double SubtreeSeparation { get; set; }
    }
}