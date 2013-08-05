using VA = VisioAutomation;

namespace VisioAutomation.Internal.Interop
{
    internal static class COMInterop
    {
        public static System.Collections.Generic.IList<VA.Internal.Interop.RunningObject> GetRunningObjects()
        {
            // Based on:
            // http://blocko.blogspot.com/2006/10/driving-excel-and-powerpoint-with-c.html
            // http://www.codeproject.com/KB/COM/ROTStuff.aspx

            var results = new System.Collections.Generic.List<RunningObject>();

            System.Runtime.InteropServices.ComTypes.IBindCtx bindctx;
            NativeMethods.CreateBindCtx(0, out bindctx);

            System.Runtime.InteropServices.ComTypes.IRunningObjectTable rot;
            bindctx.GetRunningObjectTable(out rot);
            
            System.Runtime.InteropServices.ComTypes.IEnumMoniker enum_mon;
            rot.EnumRunning(out enum_mon);
            enum_mon.Reset();

            var monikers = new System.Runtime.InteropServices.ComTypes.IMoniker[1];

            System.IntPtr numFetched = System.IntPtr.Zero;

            while (enum_mon.Next(1, monikers, numFetched) == 0)
            {
                var moniker = monikers[0];

                string name;
                moniker.GetDisplayName(bindctx, null, out name);

                object obj;
                rot.GetObject(moniker, out obj);

                System.Guid classid;
                moniker.GetClassID(out classid);
                var ro = new RunningObject(name,obj,classid); 
                results.Add(ro);
            }

            return results;
        }
    }
}