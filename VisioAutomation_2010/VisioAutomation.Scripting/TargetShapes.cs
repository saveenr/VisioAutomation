using System.Collections.Generic;
using Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Scripting
{
    public class TargetShapes
    {
        public readonly IList<Shape> Shapes;
        public TargetShapes()
        {
            // This explicitly means that the current selection is intended to be used
            this.Shapes = null;
        }
        public TargetShapes(IList<Shape> shapes)
        {
            this.Shapes = shapes;
        }

    }
}