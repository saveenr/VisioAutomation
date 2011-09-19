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
    public class Chart
    {
        public IEnumerable<T> SkipOdd<T>(IEnumerable<T> items)
        {
            int i = 0;
            foreach (var item in items)
            {
                if (i % 2 == 1)
                {
                    //
                }
                else
                {
                    yield return item;
                }
                i++;
            }

        }
       
    }

}
