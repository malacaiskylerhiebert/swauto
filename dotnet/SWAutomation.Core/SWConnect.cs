using System;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;

namespace SWAutomation.Core
{
    public static class SWConnect
    {
        public static SldWorks AttachOrLaunch(bool visible = true)
        {
            try
            {
                var running = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
                running.Visible = visible;
                return running;
            }
            catch
            {
                var t = Type.GetTypeFromProgID("SldWorks.Application", throwOnError: true);
                var app = (SldWorks)Activator.CreateInstance(t);
                app.Visible = visible;
                return app;
            }
        }
    }
}