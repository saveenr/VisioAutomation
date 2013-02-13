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

        public IDictionary<IVisio.Shape, Dictionary<string, VA.CustomProperties.CustomPropertyCells>> Get()
        {
            return this.Get(null);
        }

        public IDictionary<IVisio.Shape, Dictionary<string,VA.CustomProperties.CustomPropertyCells>> Get(IList<IVisio.Shape> target_shapes)
        {
            var prop_dic = new Dictionary<IVisio.Shape, Dictionary<string, VA.CustomProperties.CustomPropertyCells>>();
            var shapes = get_target_shapes(target_shapes);
            if (shapes.Count < 1)
            {
                return prop_dic;
            }

            var application = this.Session.VisioApplication;
            var page = application.ActivePage;
            var list_custom_props = VA.CustomProperties.CustomPropertyHelper.Get(page, shapes);

            for (int i = 0; i < shapes.Count; i++)
            {
                var shape = shapes[i];
                var props = list_custom_props[i];
                prop_dic[shape] = props;
            }

            return prop_dic;
        }

        public IList<bool> Contains(string name)
        {
            return this.Contains(null,name);
        }


        public IList<bool> Contains(IList<IVisio.Shape> target_shapes, string name)
        {
            if (name == null)
            {
                throw new System.ArgumentNullException("name");
            }

            var shapes = get_target_shapes(target_shapes);
            if (shapes.Count < 1)
            {
                return new List<bool>();
            }

            var results = (from s in this.Session.Selection.EnumShapes()
                           select VA.CustomProperties.CustomPropertyHelper.Contains(s, name))
                .ToList();

            return results;
        }

        public void Delete(string name)
        {
            this.Delete(null,name);
        }


        public void Delete(IList<IVisio.Shape> target_shapes, string name)
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

            var shapes = get_target_shapes(target_shapes);
            if (shapes.Count < 1)
            {
                return;
            }

            var application = this.Session.VisioApplication;
            using (var undoscope = new VA.Application.UndoScope(this.Session.VisioApplication, "Delete Custom Property"))
            {
                foreach (var shape in shapes)
                {
                    VA.CustomProperties.CustomPropertyHelper.Delete(shape, name);
                }
            }
        }

        public void Set(string name, VA.CustomProperties.CustomPropertyCells customprop)
        {
            this.Set(null,name,customprop);
        }

        public void Set(IList<IVisio.Shape> target_shapes, string name, VA.CustomProperties.CustomPropertyCells customprop)
        {
            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            if (customprop == null)
            {
                throw new System.ArgumentNullException("customprop");
            }

            var shapes = get_target_shapes(target_shapes);
            if (shapes.Count < 1)
            {
                return;
            }

            var application = this.Session.VisioApplication;
            using (var undoscope = new VA.Application.UndoScope(this.Session.VisioApplication, "Set Custom Property"))
            {
                foreach (var shape in shapes)
                {
                    VA.CustomProperties.CustomPropertyHelper.Set(shape, name, customprop);
                }
            }
        }

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