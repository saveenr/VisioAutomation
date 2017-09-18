namespace VisioPowerShell_Tests.Framework
{
    public static class PsArray
    {
        public static T[] From<T>(params T[] items)
        {
            return items;
        }
    }
}