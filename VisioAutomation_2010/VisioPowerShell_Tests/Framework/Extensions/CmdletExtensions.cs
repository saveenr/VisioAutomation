using System.Linq;

namespace VisioPowerShell_Tests.Framework.Extensions
{
    public static class CmdletExtensions
    {
        public static void InvokeVoid(this System.Management.Automation.Cmdlet cmd)
        {
            var results = cmd.Invoke();
            foreach (object item in results) { }
        }

        public static T InvokeFirst<T>(this System.Management.Automation.Cmdlet cmd)
        {
            var results = cmd.Invoke<T>();
            T output = results.First();
            return output;
        }
    }
}