using System.Collections.Generic;

namespace VisioAutomation.Masters
{
    public static class MasterHelper
    {
        public static Drawing.Rectangle GetBoundingBox(Microsoft.Office.Interop.Visio.Master master, Microsoft.Office.Interop.Visio.VisBoundingBoxArgs args)
        {
            // MSDN: http://msdn.microsoft.com/library/default.asp?url=/library/en-us/vissdk11/html/vimthBoundingBox_HV81900422.asp
            double bbx0, bby0, bbx1, bby1;
            master.BoundingBox((short)args, out bbx0, out bby0, out bbx1, out bby1);
            var r = new Drawing.Rectangle(bbx0, bby0, bbx1, bby1);
            return r;
        }

        public static IEnumerable<Microsoft.Office.Interop.Visio.Master> ToEnumerable(Microsoft.Office.Interop.Visio.Masters masters)
        {
            short count = masters.Count;
            for (int i = 0; i < count; i++)
            {
                yield return masters[i + 1];
            }
        }

        public static string[] GetNamesU(Microsoft.Office.Interop.Visio.Masters masters)
        {
            System.Array names_sa;
            masters.GetNamesU(out names_sa);
            string[] names = (string[])names_sa;
            return names;
        }
    }
}