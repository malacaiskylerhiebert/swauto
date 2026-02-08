using System;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SWAutomation.Core
{
    public sealed class SWSession : IDisposable
    {
        private readonly STADispatcher _sta;
        private SldWorks _app;
        private bool _connected;

        public SWSession()
        {
            _sta = new STADispatcher();
        }

        public bool IsConnected => _connected;

        private T InvokeSW<T>(Func<SldWorks, T> fn)
        {
            if (!_connected || _app == null)
                throw new InvalidOperationException("Call Connect() first.");

            return _sta.Invoke(() => fn(_app));
        }

        public void Connect(bool visible = true, bool attachIfRunning = true)
        {
            if (_connected) return;

            _sta.Invoke(() =>
            {
                if (attachIfRunning)
                {
                    try { _app = (SldWorks)Marshal.GetActiveObject("SldWorks.Application"); }
                    catch (COMException) { _app = null; }
                }

                if (_app == null)
                {
                    var t = Type.GetTypeFromProgID("SldWorks.Application");
                    _app = (SldWorks)Activator.CreateInstance(t);
                }

                _app.Visible = visible;
                _connected = true;
                return 0;
            });
        }

        public OpenDocResult OpenAssembly(string path, bool silent = true)
        {
            if (string.IsNullOrWhiteSpace(path))
                throw new ArgumentException("path is required", nameof(path));

            return InvokeSW(app =>
            {
                int errs = 0, warns = 0;

                app.OpenDoc6(
                    path,
                    (int)swDocumentTypes_e.swDocASSEMBLY,
                    silent ? (int)swOpenDocOptions_e.swOpenDocOptions_Silent : 0,
                    "",
                    ref errs,
                    ref warns
                );

                return new OpenDocResult { Errors = errs, Warnings = warns };
            });
        }

        public string GetRevision()
            => InvokeSW(app => app.RevisionNumber());

        public void Dispose()
        {
            _sta.Invoke(() =>
            {
                if (_app != null)
                {
                    try { Marshal.FinalReleaseComObject(_app); } catch { }
                    _app = null;
                }
                _connected = false;
                return 0;
            });

            _sta.Dispose();
        }
    }

    public sealed class OpenDocResult
    {
        public int Errors {  get; set; }
        public int Warnings { get; set; }
    }
}