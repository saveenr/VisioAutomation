using System.Collections.Generic;
using SMA = System.Management.Automation;

namespace VisioPS.Extensions
{
    public static class CmdletExtensions
    {
        public static void InvokeToEmpty(this SMA.Cmdlet cmdlet)
        {
            foreach (var i in cmdlet.Invoke())
            {
            }
        }

        public static IList<T> InvokeToList<T>(this SMA.Cmdlet cmdlet)
        {
            var list = new List<T>();
            foreach (var i in cmdlet.Invoke<T>())
            {
                list.Add(i);
            }
            return list;
        }

        public static void WriteVerbose(this SMA.Cmdlet cmdlet, string fmt, params object [] items)
        {
            string s = string.Format(fmt, items);
            cmdlet.WriteVerbose(s);
        }
    }
}