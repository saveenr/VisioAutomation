using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Scripting.Commands
{
    public class UserDefinedCellCommands : CommandSet
    {
        public UserDefinedCellCommands(Session session) :
            base(session)
        {

        }

        public IDictionary<IVisio.Shape, IList<VA.UserDefinedCells.UserDefinedCell>> Get(IList<IVisio.Shape> target_shapes)
        {
            this.CheckVisioApplicationAvailable();
            this.CheckActiveDrawingAvailable();

            var prop_dic = new Dictionary<IVisio.Shape, IList<VA.UserDefinedCells.UserDefinedCell>>();

            var shapes = GetTargetShapes(target_shapes);
            if (shapes.Count < 1)
            {
                return prop_dic;
            } 

            var application = this.Session.VisioApplication;
            var page = application.ActivePage;
            var list_user_props = VA.UserDefinedCells.UserDefinedCellsHelper.Get(page, shapes);

            for (int i = 0; i < shapes.Count; i++)
            {
                var shape = shapes[i];
                var props = list_user_props[i];
                prop_dic[shape] = props;
            }

            return prop_dic;
        }

        public IList<bool> Contains(IList<IVisio.Shape> target_shapes, string name)
        {
            this.CheckVisioApplicationAvailable();
            this.CheckActiveDrawingAvailable();

            if (name == null)
            {
                throw new System.ArgumentNullException("name");
            }

            var shapes = GetTargetShapes(target_shapes);
            if (shapes.Count < 1)
            {
                return new List<bool>();
            } 

            var results = (from s in this.Session.Selection.EnumShapes().ToList()
                           select VA.UserDefinedCells.UserDefinedCellsHelper.Contains(s, name))
                .ToList();

            return results;
        }
       
        public void Delete(IList<IVisio.Shape> target_shapes, string name)
        {
            this.CheckVisioApplicationAvailable();
            this.CheckActiveDrawingAvailable();

            var shapes = GetTargetShapes(target_shapes);
            if (shapes.Count < 1)
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

            using (var undoscope = new VA.Application.UndoScope(this.Session.VisioApplication,"Delete User-Defined Cell"))
            {
                foreach (var shape in shapes)
                {
                    VA.UserDefinedCells.UserDefinedCellsHelper.Delete(shape, name);
                }
            }
        }
      
        public void Set(IList<IVisio.Shape> target_shapes, VA.UserDefinedCells.UserDefinedCell userdefinedcell)
        {
            this.CheckVisioApplicationAvailable();
            this.CheckActiveDrawingAvailable();

            var shapes = GetTargetShapes(target_shapes);
            if (shapes.Count < 1)
            {
                return;
            } 

            if (userdefinedcell == null)
            {
                throw new System.ArgumentNullException("userdefinedcell");
            }

            var application = this.Session.VisioApplication;
            using (var undoscope = new VA.Application.UndoScope(this.Session.VisioApplication,"Set User-Defined Cell"))
            {
                foreach (var shape in shapes)
                {
                    VA.UserDefinedCells.UserDefinedCellsHelper.Set(shape, userdefinedcell.Name, userdefinedcell.Value, userdefinedcell.Prompt);
                }
            }
        }
    }
}