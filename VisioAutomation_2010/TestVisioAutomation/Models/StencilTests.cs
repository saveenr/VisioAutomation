using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using VisioAutomation.Shapes.CustomProperties;
using IVisio=Microsoft.Office.Interop.Visio;
using System.Linq;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class StencilTests : VisioAutomationTest
    {
        [TestMethod]
        public void RackMountedEquipment_1()
        {
            var page1 = GetNewPage();

            short flags = (short)IVisio.VisOpenSaveArgs.visOpenRO | (short)IVisio.VisOpenSaveArgs.visOpenDocked;
            var application = page1.Application;
            var documents = application.Documents;
            var rack_eq = documents.OpenEx("RCKEQP_U.VSS", flags);

            var items = new[]
                            {
                                new
                                    {
                                        master = "Server",
                                        ucount = 1,
                                        name = "foo1"
                                    },
                                new
                                    {
                                        master = "Power strip",
                                        ucount = 2,
                                        name = "foo2"
                                    },
                                new
                                    {
                                        master = "Power supply/UPS",
                                        ucount = 3,
                                        name = "foo3"
                                    },
                                new
                                    {
                                        master = "Router 1",
                                        ucount = 4,
                                        name = "foo4"
                                    },
                                new
                                    {
                                        master = "RAID array",
                                        ucount = 5,
                                        name = "foo5"
                                    }
                            };

            double eq_width = 1.5833; // inches

            double cx = 0 + eq_width/2.0;
            double cy = 0;

            double total_height = items.Select(i => i.ucount).Sum();
            var raq_eq_masters = rack_eq.Masters;
            var rack_master = raq_eq_masters["Rack"];
            var drop_pos_rack = new VA.Drawing.Point(cx, cy);
            var rack_shape = page1.Drop(rack_master, drop_pos_rack);
            CustomPropertyHelper.Set(rack_shape, "UCount", total_height.ToString());

            foreach (var item in items)
            {
                var master = raq_eq_masters[item.master];
                var drop_pos1 = new VA.Drawing.Point(cx, cy + 0.25);
                var shape1 = page1.Drop(master, drop_pos1);
                CustomPropertyHelper.Update(shape1, "UCount", item.ucount.ToString());
                shape1.Text = item.name;

                var src_height = VisioAutomation.ShapeSheet.SRCConstants.Height;
                var height_cell = shape1.CellsSRC[src_height.Section, src_height.Row, src_height.Cell];
                var actual_height = height_cell.Result[IVisio.VisUnitCodes.visNumber];

                cy += actual_height;
            }

            page1.Delete(0);
            rack_eq.Close(true);
        }
    }
}