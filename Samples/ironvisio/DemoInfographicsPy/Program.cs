using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using IG = InfoGraphicsPy;
using VA=VisioAutomation;

namespace DemoInfographicsPy
{
    class Program
    {
        static void Main(string[] args)
        {

            var igs = new InfoGraphicsPy.Session();

            igs.NewDocument();
            igs.NewPage();




            igs.TestDraw();
        }

        
    }
}
