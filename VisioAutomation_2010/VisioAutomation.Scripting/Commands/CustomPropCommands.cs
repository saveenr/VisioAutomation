using System.Collections.Generic;
using System.Linq;
using CP=VisioAutomation.Shapes.CustomProperties;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Scripting.Commands
{
    public class CustomPropCommands : CommandSet
    {
        public CustomPropCommands(Client client) :
            base(client)
        {

        }

        public IDictionary<IVisio.Shape, Dictionary<string,CP.CustomPropertyCells>> Get(IList<IVisio.Shape> target_shapes)
        {
            this.AssertApplicationAvailable();
            this.AssertDocumentAvailable();

            var prop_dic = new Dictionary<IVisio.Shape, Dictionary<string, CP.CustomPropertyCells>>();
            var shapes = GetTargetShapes(target_shapes);
            if (shapes.Count < 1)
            {
                return prop_dic;
            }

            var application = this.Client.VisioApplication;
            var page = application.ActivePage;

            var list_custom_props = CP.CustomPropertyHelper.Get(page, shapes);

            for (int i = 0; i < shapes.Count; i++)
            {
                var shape = shapes[i];
                var props = list_custom_props[i];
                prop_dic[shape] = props;
            }

            return prop_dic;
        }

        public IList<bool> Contains(IList<IVisio.Shape> target_shapes, string name)
        {
            if (name == null)
            {
                throw new System.ArgumentNullException("name");
            }

            var shapes = GetTargetShapes(target_shapes);
            if (shapes.Count < 1)
            {
                return new List<bool>();
            }

            var results = this.Client.Selection.GetShapes().Select(s => CP.CustomPropertyHelper.Contains(s, name)).ToList();

            return results;
        }

        public void Delete(IList<IVisio.Shape> target_shapes, string name)
        {
            this.AssertApplicationAvailable();
            this.AssertDocumentAvailable();
            
            if (name == null)
            {
                throw new System.ArgumentNullException("name");
            }

            if (name.Length < 1)
            {
                throw new System.ArgumentException("name");
            }

            var shapes = GetTargetShapes(target_shapes);
            if (shapes.Count < 1)
            {
                return;
            }

            var application = this.Client.VisioApplication;
            using (var undoscope = new VA.Application.UndoScope(this.Client.VisioApplication, "Delete Custom Property"))
            {
                foreach (var shape in shapes)
                {
                    CP.CustomPropertyHelper.Delete(shape, name);
                }
            }
        }

        public void Set(IList<IVisio.Shape> target_shapes, string name, CP.CustomPropertyCells customprop)
        {
            this.AssertApplicationAvailable();
            this.AssertDocumentAvailable();
            
            if (customprop == null)
            {
                throw new System.ArgumentNullException("customprop");
            }

            var shapes = GetTargetShapes(target_shapes);
            if (shapes.Count < 1)
            {
                return;
            }

            var application = this.Client.VisioApplication;
            using (var undoscope = new VA.Application.UndoScope(this.Client.VisioApplication, "Set Custom Property"))
            {
                foreach (var shape in shapes)
                {
                    CP.CustomPropertyHelper.Set(shape, name, customprop);
                }
            }
        }

        public IEnumerable<IVisio.Shape> EnumerateAndSelect(IEnumerable<IVisio.Shape> shapes)
        {
            this.AssertApplicationAvailable();
            this.AssertDocumentAvailable();
            
            if (shapes == null)
            {
                throw new System.ArgumentNullException("shapes");
            }

            foreach (var shape in shapes)
            {
                this.Client.Selection.None();
                this.Client.Selection.Select(shape);
                yield return shape;
            }
        }
    }
}