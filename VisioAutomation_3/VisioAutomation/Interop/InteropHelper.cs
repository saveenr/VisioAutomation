using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using IVisio=Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.Interop
{
    public class EnumValue
    {
        public string Name { get; private set; }
        public int Value { get; private set; }

        public EnumValue(string name, int value)
        {
            this.Name = name;
            this.Value = value;
        }
    }

    public class EnumType
    {
        public System.Type Type { get; private set; }
        public string Name { get; private set; }
        public List<EnumValue> Values { get; private set; }
        public Dictionary<string,int> NameToValue { get; private set; }
        

        internal EnumType(System.Type t)
        {
            this.Type = t;
            this.Name = t.Name;
            this.Values = this.GetEnumValues().ToList();
            this.NameToValue = this.Values.ToDictionary(i => i.Name, i => i.Value);
        }

        public IEnumerable<EnumValue> GetEnumValues()
        {
            string[] names = System.Enum.GetNames(this.Type);
            System.Array avalues = System.Enum.GetValues(this.Type);
            for (int i = 0; i < avalues.Length; i++)
            {
                var item = new EnumValue(names[i],(int) avalues.GetValue(i));
                yield return item;
            }
        }

    }

    public static class InteropHelper
    {

        private static bool inited=false;
        private static Dictionary<string, EnumType> g_name_to_enum;
        private static List<System.Type> g_types; 

        private static List<Type> GetTypes()
        {
            init();
            return g_types;
        }


        public static void init()
        {
            if (!inited)
            {
                g_types = typeof(IVisio.Application).Assembly.GetExportedTypes()
                    .Where(t => t.IsPublic)
                    .Where(t => !t.Name.StartsWith("tag"))
                    .ToList();
                g_name_to_enum = g_types
                    .Where(t => t.IsEnum)
                    .Select(i => new EnumType(i))
                    .ToDictionary(i => i.Name, i => i);
                inited = true;
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
