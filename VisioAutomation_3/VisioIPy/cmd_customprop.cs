using VAS = VisioAutomation.Scripting;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;

namespace VisioIPy
{
    public partial class VisioIPySession
    {
        public IDictionary<IVisio.Shape, Dictionary<string,VA.CustomProperties.CustomPropertyCells>> GetCustomProperties()
        {
            var prop_dic = this.ScriptingSession.CustomProp.GetCustomProperties();
            return prop_dic;
        }

        public void SetCustomProperty(string name, string value)
        {
            var cp = new VA.CustomProperties.CustomPropertyCells();
            cp.Value = value;
            this.ScriptingSession.CustomProp.SetCustomProperty(name, cp);
        }

        public void SetCustomProperty(string name, VA.CustomProperties.CustomPropertyCells customprop)
        {
            if (customprop == null)
            {
                throw new System.ArgumentNullException("customprop");
            }

            this.ScriptingSession.CustomProp.SetCustomProperty(name,customprop);
        }

        public void DeleteCustomProperty(string name)
        {
            this.ScriptingSession.CustomProp.DeleteCustomProperty(name);
        }

    }
}