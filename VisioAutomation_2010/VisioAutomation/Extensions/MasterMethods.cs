using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class MasterMethods
    {
        public static Drawing.Rectangle GetBoundingBox(this IVisio.Master master, IVisio.VisBoundingBoxArgs args)
        {
            // MSDN: http://msdn.microsoft.com/library/default.asp?url=/library/en-us/vissdk11/html/vimthBoundingBox_HV81900422.asp
            double bbx0, bby0, bbx1, bby1;
            master.BoundingBox((short) args, out bbx0, out bby0, out bbx1, out bby1);
            var r = new Drawing.Rectangle(bbx0, bby0, bbx1, bby1);
            return r;
        }

        public static IEnumerable<IVisio.Master> AsEnumerable(this IVisio.Masters masters)
        {
            short count = masters.Count;
            for (int i = 0; i < count; i++)
            {
                yield return masters[i + 1];
            }
        }

        public static string[] GetNamesU(this IVisio.Masters masters)
        {
            System.Array names_sa;
            masters.GetNamesU(out names_sa);
            string[] names = (string[])names_sa;
            return names;
        }
    }
}