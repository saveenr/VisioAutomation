using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Interop
{
    public static class InteropHelper
    {
        public static List<Type> GetEnumTypes()
        {
            return typeof(IVisio.Application).Assembly.GetExportedTypes()
                .Where(t => t.IsEnum)
                .Where(t => t.IsPublic)
                .Where(t => !t.Name.StartsWith("tag"))
                .ToList();
        }
    }
}
