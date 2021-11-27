using System.Collections.Generic;

namespace VTest.PowerShell.Framework
{
    public static class PsArray
    {
        public static T[] From<T>(params T[] items)
        {
            return items;
        }

        public static T[] From<T>(List<T> items)
        {
            return items.ToArray();
        }
    }
}