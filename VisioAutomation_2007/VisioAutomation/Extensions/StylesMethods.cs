using System;
using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.Extensions
{
    public static partial class StylesMethods
    {
        public static string[] GetNamesU(this IVisio.Styles styles)
        {
            System.Array names_sa;
            styles.GetNamesU(out names_sa);
            string[] names = (string[])names_sa;
            return names;
        }
    }
}