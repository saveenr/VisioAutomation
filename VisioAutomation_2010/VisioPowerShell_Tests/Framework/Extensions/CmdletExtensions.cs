using System.Linq;

namespace VisioPowerShell_Tests.Framework.Extensions
{
    public static class CmdletExtensions
    {
        public static void ExInvokeVoid(this System.Management.Automation.Cmdlet cmd)
        {
            var results = cmd.Invoke();
            foreach (var z in results)
            {

            }
        }

        public static T ExInvokeFirst<T>(this System.Management.Automation.Cmdlet cmd)
        {
            var results = cmd.Invoke<T>();
            T output = results.First();
            return output;
        }
    }
}