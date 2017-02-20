using System.Collections.Generic;
using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Windows
{
    public static class WindowHelper
    {
        public static void Select(
            IVisio.Window window,
            IEnumerable<IVisio.Shape> shapes,
            IVisio.VisSelectArgs selectargs)
        {
            if (shapes == null)
            {
                throw new System.ArgumentNullException(nameof(shapes));
            }

            foreach (var shape in shapes)
            {
                window.Select(shape, (short)selectargs);
            }
        }
    }
}