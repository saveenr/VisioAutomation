using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;

namespace TestVisioAutomation
{
    public class VisioApplicationSafeReference
    {
        // this class ensures that a valid application instance is always available

        private IVisio.Application app;

        public IVisio.Application GetVisioApplication()
        {
            if (this.app == null)
            {
                this.app = new IVisio.Application();
            }
            else
            {
                // we have an instance, but it may not be valid
                try
                {
                    string s = app.Name;
                }
                catch (System.Runtime.InteropServices.COMException e)
                {
                    this.app = new IVisio.Application();
                }
            }

            return this.app;
        }
    }
}