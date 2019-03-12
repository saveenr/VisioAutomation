using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Helpers
{
    public static class InteropHelper
    {
        private static bool _initialized = false;
        private static Dictionary<string, Models.EnumType> _gNameToEnum;
        private static List<System.Type> _gTypes;

        private static void initialize()
        {
            if (!InteropHelper._initialized)
            {
                InteropHelper._gTypes = typeof(IVisio.Application).Assembly.GetExportedTypes()
                    .Where(t => t.IsPublic)
                    .Where(t => !t.Name.StartsWith("tag"))
                    .OrderBy(t => t.Name)
                    .ToList();
                InteropHelper._gNameToEnum = InteropHelper._gTypes
                    .Where(t => t.IsEnum)
                    .Select(i => new Models.EnumType(i))
                    .ToDictionary(i => i.Name, i => i);
                InteropHelper._initialized = true;
            }
        }

        public static List<Models.EnumType> GetEnums()
        {
            InteropHelper.initialize();
            return InteropHelper._gNameToEnum.Values.ToList();
        }

        public static Models.EnumType GetEnum(string name)
        {
            InteropHelper.initialize();
            return InteropHelper._gNameToEnum[name];
        }

        public static Models.EnumType GetEnum(System.Type t)
        {
            InteropHelper.initialize();
            return InteropHelper._gNameToEnum[t.Name];
        }
    }
}