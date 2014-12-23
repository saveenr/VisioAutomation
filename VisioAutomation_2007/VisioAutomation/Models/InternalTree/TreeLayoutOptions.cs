using VA = VisioAutomation;

namespace VisioAutomation.Models.InternalTree
{
    internal class TreeLayoutOptions
    {
        public VA.Drawing.Point TopAdjustment; // How to adjust the apex 

        public TreeLayoutOptions()
        {
            SubtreeSeparation = 1;
            SiblingSeparation = 1;
            Direction = LayoutDirection.Up;
            Alignment = VA.Drawing.AlignmentVertical.Top;
            MaximumDepth = 100;
            LevelSeparation = 1;
            DefaultNodeSize = new VA.Drawing.Size(1, 1);
        }

        public VA.Drawing.Size DefaultNodeSize { get; set; }
        public double LevelSeparation { get; set; }
        public int MaximumDepth { get; set; }
        public VA.Drawing.AlignmentVertical Alignment { get; set; }
        public LayoutDirection Direction { get; set; }
        public double SiblingSeparation { get; set; }
        public double SubtreeSeparation { get; set; }
    }
}