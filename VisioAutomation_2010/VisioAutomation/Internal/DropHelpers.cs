using System.Collections.Generic;

namespace VisioAutomation.Internal
{
    internal class DropHelpers
    {

        public static void ValidateDropManyParams(IList<Microsoft.Office.Interop.Visio.Master> masters, IEnumerable<Core.Point> points)
        {
            if (masters == null)
            {
                throw new System.ArgumentNullException(nameof(masters));
            }

            if (points == null)
            {
                throw new System.ArgumentNullException(nameof(points));
            }
        }

    }
}