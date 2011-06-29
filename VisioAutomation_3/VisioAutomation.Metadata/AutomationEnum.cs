using System.Collections.Generic;

namespace VisioAutomation.Metadata
{
    public class AutomationEnumItem
    {
        public string Name;
        public int Value;

        internal AutomationEnumItem(string name, int value)
        {
            this.Name = name;
            this.Value = value;
        }
    }

    public class AutomationEnum
    {
        public string Name;
        public List<AutomationEnumItem> Items;
        private Dictionary<string, int> _dic;

        public AutomationEnum(string name)
        {
            this.Name = name;
        }

        internal void Add(string name, int value)
        {
            if (this.Items == null)
            {
                this.Items = new List<AutomationEnumItem>();
            }

            var i = new AutomationEnumItem(name, value);
            this.Items.Add(i);
        }

        public int this[string index]
        {
            get
            {
                checkdic();
                return this._dic[index];
            }
        }

        public bool HasItem(string name)
        {
            this.checkdic();
            return this._dic.ContainsKey(name);
        }

        private void checkdic()
        {
            if (this._dic == null)
            {
                this._dic = new Dictionary<string, int>();
                foreach (var i in this.Items)
                {
                    this._dic[i.Name] = i.Value;
                }
            }
        }
    }
}