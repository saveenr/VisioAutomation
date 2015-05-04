﻿using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.CmdletAttribute(SMA.VerbsCommon.Redo, "Visio")]
    public class Redo_Visio : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            this.client.Application.Redo();
        }
    }
}