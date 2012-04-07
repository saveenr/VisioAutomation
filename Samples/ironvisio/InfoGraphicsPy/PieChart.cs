using System;
using System.Collections;
using System.Collections.Generic;
using IVisio=Microsoft.Office.Interop.Visio;
using IG=InfoGraphicsPy;
using System.Linq;
using VA=VisioAutomation;
using VisioAutomation.Extensions;

namespace InfoGraphicsPy
{
    public class PieChart : Chart
    {
        public PieChart(DataPoints dps, string [] cats) 
            : base(dps,cats)
        {
        }

        public void Draw(Session session)
        {
            var pc = new VA.Layout.Models.Pie.PieLayout();
            foreach (var dp in this.DataPoints)
            {
                pc.Add(dp.Value, dp.Text);
            }

            pc.DrawBackground = true;

            pc.Radius = 2.0;
            pc.Center = new VA.Drawing.Point(3,3);

            pc.Render(session.Page);
        }
    }
}
