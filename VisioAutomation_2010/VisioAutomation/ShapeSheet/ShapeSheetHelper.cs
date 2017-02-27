using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet
{
    internal static class ShapeSheetHelper
    {
        public static string GetSectionName(IVisio.VisSectionIndices value)
        {
            string s = value.ToString();
            const int start_index = 10; // Length of string "visSection"
            return s.Substring(start_index); // Get Rid of the visSection prefix
        }

        public static IEnumerable<IVisio.Row> ToEnumerable(IVisio.Section section)
        {
            // Section object: http://msdn.microsoft.com/en-us/library/ms408988(v=office.12).aspx

            int row_count = section.Count;

            for (int i = 0; i < row_count; i++)
            {
                var row = section[(short)i];
                yield return row;
            }
        }
    }
}