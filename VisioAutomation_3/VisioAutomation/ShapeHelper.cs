using System.Collections.Generic;
using System.Xml.Linq;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation
{
    public static class ShapeHelper
    {
        /// <summary>
        /// Enumerates all shapes contained by a set of shapes recursively
        /// </summary>
        /// <param name="shapes">the set of shapes to start the enumeration</param>
        /// <returns>The enumeration</returns>
        public static IList<IVisio.Shape> GetNestedShapes(IEnumerable<IVisio.Shape> shapes)
        {
            if (shapes == null)
            {
                throw new System.ArgumentNullException("shapes");
            }

            var result = new List<IVisio.Shape>();
            var stack = new Stack<IVisio.Shape>(shapes);

            while (stack.Count > 0)
            {
                var s = stack.Pop();
                var subshapes = s.Shapes;
                if (subshapes.Count > 0)
                {
                    foreach (var child in subshapes.AsEnumerable())
                    {
                        stack.Push(child);
                    }
                }

                result.Add(s);
            }

            return result;
        }

        public static IList<IVisio.Shape> GetNestedShapes(IVisio.Shape shape)
        {
            if (shape== null)
            {
                throw new System.ArgumentNullException("shape");
            }

            var shapes = new[] {shape};

            return GetNestedShapes(shapes);
        }

        public static IList<IVisio.Shape> GetShapesFromIDs(IVisio.Shapes shapes, IList<short> shapeids)
        {
            var shape_objs = new List<IVisio.Shape>(shapeids.Count);
            foreach (short shapeid in shapeids)
            {
                var shape = shapes.ItemFromID16[shapeid];
                shape_objs.Add(shape);
            }
            return shape_objs;
        }

        public static void SetGroupSelectMode(IVisio.Shape shape, IVisio.VisCellVals mode)
        {
            var src_selectmode = VA.ShapeSheet.SRCConstants.SelectMode;
            var select_mode_cell = shape.CellsSRC[src_selectmode.Section, src_selectmode.Row, src_selectmode.Cell];
            select_mode_cell.FormulaU = ((int)mode).ToString();
        }

        public static System.Xml.Linq.XElement GetShapeDescriptionXML(IVisio.Shape shape)
        {
            var culture = System.Globalization.CultureInfo.InvariantCulture;
            var xform = VA.Layout.LayoutHelper.GetXForm(shape);
            var fmt = VA.Format.FormatHelper.GetShapeFormat(shape);
            string text = shape.Text;
            
            XElement el_shape = _CreateShapeXMLDesc(text, xform, fmt);

            return el_shape;
        }

        public static IList<System.Xml.Linq.XElement> GetShapeDescriptionXML(IVisio.Page page, IList<int> shapeids)
        {
            var culture = System.Globalization.CultureInfo.InvariantCulture;
            var xform = VA.Layout.LayoutHelper.GetXForm(page, shapeids);
            var fmt = VA.Format.FormatHelper.GetShapeFormat(page, shapeids);
            var pageshapes = page.Shapes;
            var text = shapeids.Select(id => pageshapes.ItemFromID[id].Text).ToList();


            var list =  new List<System.Xml.Linq.XElement>();

            for (int i = 0; i < shapeids.Count; i++)
            {

                var e = _CreateShapeXMLDesc(text[i], xform[i], fmt[i]);

                list.Add(e);
            }

             return list;
        }

        private static XElement _CreateShapeXMLDesc(string text, VA.Layout.XFormCells xform, VA.Format.ShapeFormatCells fmt)
        {
            var el_shape = new System.Xml.Linq.XElement("Shape");

            var el_text = new System.Xml.Linq.XText(text);
            el_shape.Add(el_text);

            var el_xform= new System.Xml.Linq.XElement("XForm");
            el_shape.Add(el_xform);

            var el_fill = new System.Xml.Linq.XElement("Fill");
            el_shape.Add(el_fill);

            var el_line = new System.Xml.Linq.XElement("Line");
            el_shape.Add(el_line);

            var el_shadow = new System.Xml.Linq.XElement("Shadow");
            el_shape.Add(el_shadow);

            var el_textbox = new System.Xml.Linq.XElement("TextBox");
            el_shape.Add(el_textbox);

            el_xform.SetAttributeValue("PinX", xform.PinX.Formula.ToDisplayString());
            el_xform.SetAttributeValue("PinY", xform.PinY.Formula.ToDisplayString());
            el_xform.SetAttributeValue("LocPinX", xform.LocPinX.Formula.ToDisplayString());
            el_xform.SetAttributeValue("LocPinY", xform.LocPinY.Formula.ToDisplayString());
            el_xform.SetAttributeValue("Width", xform.Width.Formula.ToDisplayString());
            el_xform.SetAttributeValue("Height", xform.Height.Formula.ToDisplayString());
            el_xform.SetAttributeValue("Angle", xform.Angle.Formula.ToDisplayString());


            el_fill.SetAttributeValue("FillBkgnd", fmt.FillBkgnd.Formula.ToDisplayString());
            el_fill.SetAttributeValue("FillBkgndTrans", fmt.FillBkgndTrans.Formula.ToDisplayString());
            el_fill.SetAttributeValue("FillForegnd", fmt.FillForegnd.Formula.ToDisplayString());
            el_fill.SetAttributeValue("FillForegndTrans", fmt.FillForegndTrans.Formula.ToDisplayString());
            el_fill.SetAttributeValue("FillPattern", fmt.FillPattern.Formula.ToDisplayString());
            
            el_line.SetAttributeValue("LineCap", fmt.LineCap.Formula.ToDisplayString());
            el_line.SetAttributeValue("LineColor", fmt.LineColor.Formula.ToDisplayString());
            el_line.SetAttributeValue("LineColorTrans", fmt.LineColorTrans.Formula.ToDisplayString());
            el_line.SetAttributeValue("LinePattern", fmt.LinePattern.Formula.ToDisplayString());
            el_line.SetAttributeValue("LineWeight", fmt.LineWeight.Formula.ToDisplayString());
            el_line.SetAttributeValue("Rounding", fmt.Rounding.Formula.ToDisplayString());
            el_line.SetAttributeValue("BeginArrow", fmt.BeginArrow.Formula.ToDisplayString());
            el_line.SetAttributeValue("BeginArrowSize", fmt.BeginArrowSize.Formula.ToDisplayString());
            el_line.SetAttributeValue("EndArrow", fmt.EndArrow.Formula.ToDisplayString());
            el_line.SetAttributeValue("EndArrowSize", fmt.EndArrowSize.Formula.ToDisplayString());

            el_shadow.SetAttributeValue("ShapeShdwObliqueAngle", fmt.ShapeShdwObliqueAngle.Formula.ToDisplayString());
            el_shadow.SetAttributeValue("ShapeShdwOffsetX", fmt.ShapeShdwOffsetX.Formula.ToDisplayString());
            el_shadow.SetAttributeValue("ShapeShdwOffsetY", fmt.ShapeShdwOffsetY.Formula.ToDisplayString());
            el_shadow.SetAttributeValue("ShapeShdwScaleFactor", fmt.ShapeShdwScaleFactor.Formula.ToDisplayString());
            el_shadow.SetAttributeValue("ShapeShdwType", fmt.ShapeShdwType.Formula.ToDisplayString());
            el_shadow.SetAttributeValue("ShdwBkgnd", fmt.ShdwBkgnd.Formula.ToDisplayString());
            el_shadow.SetAttributeValue("ShdwBkgndTrans", fmt.ShdwBkgndTrans.Formula.ToDisplayString());
            el_shadow.SetAttributeValue("ShdwForegnd", fmt.ShdwForegnd.Formula.ToDisplayString());
            el_shadow.SetAttributeValue("ShdwForegndTrans", fmt.ShdwForegndTrans.Formula.ToDisplayString());
            el_shadow.SetAttributeValue("ShdwPattern", fmt.ShdwPattern.Formula.ToDisplayString());

            el_textbox.SetAttributeValue("TextBkgnd", fmt.TextBkgnd.Formula.ToDisplayString());
            el_textbox.SetAttributeValue("TextBkgndTrans", fmt.TextBkgndTrans.Formula.ToDisplayString());

            return el_shape;
        }
    }
}