using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Scripting.Commands
{
    public class CustomPropCommands : CommandSet
    {
        public CustomPropCommands(Session session) :
            base(session)
        {

        }

        public IDictionary<IVisio.Shape, Dictionary<string,VA.CustomProperties.CustomPropertyCells>> GetCustomProperties()
        {
            var prop_dic = new Dictionary<IVisio.Shape, Dictionary<string,VA.CustomProperties.CustomPropertyCells>>();
            if (!this.Session.HasSelectedShapes())
            {
                return prop_dic;
            }

            var shapes = this.Session.Selection.EnumShapes().ToList();
            var application = this.Session.VisioApplication;
            var page = application.ActivePage;
            var list_custom_props = VA.CustomProperties.CustomPropertyHelper.GetCustomProperties(page, shapes);

            for (int i = 0; i < shapes.Count; i++)
            {
                var shape = shapes[i];
                var props = list_custom_props[i];
                prop_dic[shape] = props;
            }

            return prop_dic;
        }

        public IList<bool> HasCustomProperty(string name)
        {
            if (name == null)
            {
                throw new System.ArgumentNullException("name");
            }

            if (!this.Session.HasSelectedShapes())
            {
                return new List<bool>();
            }

            var results = (from s in this.Session.Selection.EnumShapes()
                           select VA.CustomProperties.CustomPropertyHelper.HasCustomProperty(s, name))
                .ToList();

            return results;
        }

        public void DeleteCustomProperty(string name)
        {
            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            if (name == null)
            {
                throw new System.ArgumentNullException("name");
            }

            if (name.Length < 1)
            {
                throw new System.ArgumentException("name");
            }

            var shapes = this.Session.Selection.EnumShapes().ToList();

            var application = this.Session.VisioApplication;
            using (var undoscope = application.CreateUndoScope())
            {
                foreach (var shape in shapes)
                {
                    VA.CustomProperties.CustomPropertyHelper.DeleteCustomProperty(shape, name);
                }
            }
        }

        public void SetCustomProperty(string name, VA.CustomProperties.CustomPropertyCells customprop)
        {
            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            if (customprop == null)
            {
                throw new System.ArgumentNullException("customprop");
            }

            var shapes = this.Session.Selection.EnumShapes().ToList();

            var application = this.Session.VisioApplication;
            using (var undoscope = application.CreateUndoScope())
            {
                foreach (var shape in shapes)
                {
                    VA.CustomProperties.CustomPropertyHelper.SetCustomProperty(shape, name, customprop);
                }
            }
        }

        /// <summary>
        /// Given a set of shapes will will enumerate each shape and set the selection to that shape
        /// </summary>
        /// <param name="scripting_session"></param>
        /// <param name="shapes"></param>
        /// <returns></returns>
        public IEnumerable<IVisio.Shape> EnumerateAndSelect(IEnumerable<IVisio.Shape> shapes)
        {
            if (shapes == null)
            {
                throw new System.ArgumentNullException("shapes");
            }

            foreach (var shape in shapes)
            {
                this.Session.Selection.SelectNone();
                this.Session.Selection.Select(shape);
                yield return shape;
            }
        }
    }
}