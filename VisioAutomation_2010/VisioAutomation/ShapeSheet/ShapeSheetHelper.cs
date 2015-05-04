namespace VisioAutomation.ShapeSheet
{
    internal static class ShapeSheetHelper
    {
        public static string GetSectionName(Microsoft.Office.Interop.Visio.VisSectionIndices value)
        {
            string s = value.ToString();
            return s.Substring(10); // Get Rid of the visSection prefix
        }

        public static string GetSectionName(int value, string defaultname)
        {
            if (System.Enum.IsDefined(typeof(Microsoft.Office.Interop.Visio.VisSectionIndices), value))
            {
                var a = (Microsoft.Office.Interop.Visio.VisSectionIndices)value;
                return ShapeSheetHelper.GetSectionName(a);
            }
            return defaultname;
        }
    }
}