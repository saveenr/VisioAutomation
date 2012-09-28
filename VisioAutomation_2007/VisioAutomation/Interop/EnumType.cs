using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.Interop
{
    public class EnumType
    {
        public System.Type Type { get; private set; }
        public string Name { get; private set; }
        public List<EnumValue> Values { get; private set; }
        public Dictionary<string,int> NameToValue { get; private set; }
        
        public EnumType(System.Type t)
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
                object o = avalues.GetValue(i);
                System.Type t = o.GetType();
                var item = new EnumValue(names[i],(int) o);
                yield return item;
            }
        }
    }
}