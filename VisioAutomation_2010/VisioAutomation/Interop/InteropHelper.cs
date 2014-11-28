using System.Collections.Generic;
using System.Linq;
using IVisio=Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.Interop
{
    public static class InteropHelper
    {
        private static bool _initialized=false;
        private static Dictionary<string, EnumType> g_name_to_enum;
        private static List<System.Type> g_types; 

        public static void init()
        {
            if (!_initialized)
            {
                g_types = typeof(IVisio.Application).Assembly.GetExportedTypes()
                    .Where(t => t.IsPublic)
                    .Where(t => !t.Name.StartsWith("tag"))
                    .OrderBy(t=>t.Name)
                    .ToList();
                g_name_to_enum = g_types
                    .Where(t => t.IsEnum)
                    .Select(i => new EnumType(i))
                    .ToDictionary(i => i.Name, i => i);
                _initialized = true;
            }
        }

        public static List<EnumType> GetEnums()
        {
            init();
            return g_name_to_enum.Values.ToList();
        }

        public static EnumType GetEnum(string name)
        {
            init();
            return g_name_to_enum[name];
        }

        public static EnumType GetEnum(System.Type t)
        {
            init();
            return g_name_to_enum[t.Name];
        }
    }
}
