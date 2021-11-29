﻿using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class MasterMethods_Drop
    {

        public static IVisio.Shape Drop(
            this IVisio.Master master1,
            IVisio.Master master2,
            Core.Point point)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(master1);
            return VisioAutomation.Internal.Extensions.ExtensionHelpers._Drop(visobjtarget, master2, point);
        }

        public static short[] DropManyU(
            this IVisio.Master master,
            IList<IVisio.Master> masters,
            IEnumerable<Core.Point> points)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(master);
            return VisioAutomation.Internal.Extensions.ExtensionHelpers._DropManyU(visobjtarget, masters, points);
        }
    }
}