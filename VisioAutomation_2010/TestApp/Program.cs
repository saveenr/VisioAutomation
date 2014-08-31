using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Application;
using VA=VisioAutomation;
using VisioAutomation.Extensions;

namespace TestApp
{
    class Program
    {
        static void Main(string[] args)
        {

            // Create a new app and doc
            var app = new Microsoft.Office.Interop.Visio.ApplicationClass();
            var docs = app.Documents;

            // Create a new doc
            var doc = docs.Add("");

            // Create a new master (it will contain no shapes)
            var masters = doc.Masters;
            IVisio.Master m = masters.Add();
            m.Name = "MyMasterShape";


            // Edit the master by adding a shape
            var mdraw_window = m.OpenDrawWindow();
            mdraw_window.Activate();
            var st = mdraw_window.SubType;
            if (st == 64)
            {
                // Master Property: http://msdn.microsoft.com/en-us/library/ms426178(v=office.12).aspx
                // SubType Property: http://msdn.microsoft.com/en-us/library/office/ff766045(v=office.15).aspx

                var master = (IVisio.Master) mdraw_window.Master;

                var shape = master.DrawRectangle(0, 0, 1, 1);
                shape.Cells["FillForegnd"].Formula = "=rgb(0,255,0)";
                master.Close();
            }
            
            // Done with the master

            // Now drop it twice
            var m1 = masters[m.Name];
            var ds1 = app.ActivePage.Drop(m1, 5, 5);
            var ds2 = app.ActivePage.Drop(m1, 2, 2);

            // And go back to edit it by add
            mdraw_window = m.OpenDrawWindow();
            mdraw_window.Activate();
            st = mdraw_window.SubType;
            if (st == 64)
            {
                // Master Property: http://msdn.microsoft.com/en-us/library/ms426178(v=office.12).aspx
                IVisio.Master o = (IVisio.Master)mdraw_window.Master;

                var s = o.DrawRectangle(1, 1, 2, 2);
                s.Cells["FillForegnd"].Formula = "=rgb(0,0,255)";
                o.Close();
            }

            // Now the two drop shapes should both be update




        }
    }
}
