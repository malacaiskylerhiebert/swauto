using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace SWAutomation.Core
{
    public sealed class SWSession : IDisposable
    {
        private readonly STADispatcher _sta;
        private SldWorks _app;
        private bool _connected;
        private bool _launchedByUs;

        private readonly Dictionary<string, ModelDoc2> _docs = new Dictionary<string, ModelDoc2>();

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
                    _launchedByUs = true;
                }
                else
                {
                    _launchedByUs = false;
                }

                    _app.Visible = visible;
                _connected = true;
                return 0;
            });
        }

        private string TrackDoc(ModelDoc2 doc)
        {
            var id = Guid.NewGuid().ToString("N");
            _docs[id] = doc;
            return id;
        }

        private ModelDoc2 GetDoc(string docId)
        {
            if (string.IsNullOrWhiteSpace(docId) || !_docs.TryGetValue(docId, out var doc) || doc == null)
                throw new KeyNotFoundException($"Unknown docId: {docId}");
            return doc;
        }

        public string OpenPart(string path, bool silent = true)
        {
            return InvokeSW(app =>
            {
                int e = 0, w = 0;
                var doc = (ModelDoc2)app.OpenDoc6(
                    path,
                    (int)swDocumentTypes_e.swDocPART,
                    silent ? (int)swOpenDocOptions_e.swOpenDocOptions_Silent : 0,
                    "",
                    ref e,
                    ref w
                );
                if (doc == null) throw new InvalidOperationException($"Failed to open part: {path} (errors={e}, warnings={w})");
                return TrackDoc(doc);
            });
        }

        public string OpenAssembly(string path, bool silent = true)
        {
            return InvokeSW(app =>
            {
                int e = 0, w = 0;
                var doc = (ModelDoc2)app.OpenDoc6(
                    path,
                    (int)swDocumentTypes_e.swDocASSEMBLY,
                    silent ? (int)swOpenDocOptions_e.swOpenDocOptions_Silent : 0,
                    "",
                    ref e,
                    ref w
                );
                if (doc == null) throw new InvalidOperationException($"Failed to open assembly: {path} (errors={e}, warnings={w})");
                return TrackDoc(doc);
            });
        }

        public string AssemblyAddComponent(
            string assemblyId,
            string partRef,   // can be docId OR file path
            double x = 0.0,
            double y = 0.0,
            double z = 0.0
        )
        {
            if (string.IsNullOrWhiteSpace(partRef))
                throw new ArgumentException("partRef is required", nameof(partRef));

            return InvokeSW(app =>
            {
                var asmDoc = GetDoc(assemblyId) as AssemblyDoc;
                if (asmDoc == null)
                    throw new InvalidOperationException("docId is not an assembly.");

                Component2 comp = null;

                // Case 1: partRef is a tracked docId
                ModelDoc2 partDoc;
                if (_docs.TryGetValue(partRef, out partDoc))
                {
                    if (partDoc == null)
                        throw new InvalidOperationException("Referenced docId is null.");

                    var partPath = partDoc.GetPathName();
                    if (string.IsNullOrWhiteSpace(partPath))
                        throw new InvalidOperationException("Referenced part doc has no saved path. Save it first.");

                    comp = (Component2)asmDoc.AddComponent5(
                        partPath,
                        (int)swAddComponentConfigOptions_e.swAddComponentConfigOptions_CurrentSelectedConfig,
                        "",
                        false,
                        "",
                        x, y, z
                    );
                }
                else
                {
                    // Case 2: partRef is a file path
                    comp = (Component2)asmDoc.AddComponent5(
                        partRef,
                        (int)swAddComponentConfigOptions_e.swAddComponentConfigOptions_CurrentSelectedConfig,
                        "",
                        false,
                        "",
                        x, y, z
                    );
                }

                if (comp == null)
                    throw new InvalidOperationException("Failed to add component.");

                return comp.Name2;
            });
        }

        public void AssemblyRemoveComponent(string assemblyId, string componentRef)
        {
            if (string.IsNullOrWhiteSpace(componentRef))
                throw new ArgumentException("componentRef is required", nameof(componentRef));

            InvokeSW(app =>
            {
                var doc = GetDoc(assemblyId);
                var asm = doc as AssemblyDoc;
                if (asm == null)
                    throw new InvalidOperationException("docId is not an assembly.");

                var comp = ResolveComponent(asm, componentRef);

                // Select component and delete
                doc.ClearSelection2(true);
                bool selOk = comp.Select4(false, null, false);
                if (!selOk)
                    throw new InvalidOperationException("Failed to select component.");

                bool delOk = doc.Extension.DeleteSelection2(0);
                if (!delOk)
                    throw new InvalidOperationException("Failed to delete component.");

                return 0;
            });
        }

        public void AssemblySetComponentFixed(string assemblyId, string componentRef, bool fixedInPlace)
        {
            if (string.IsNullOrWhiteSpace(componentRef))
                throw new ArgumentException("componentRef is required", nameof(componentRef));

            InvokeSW(app =>
            {
                var doc = GetDoc(assemblyId);
                var asm = doc as AssemblyDoc;
                if (asm == null)
                    throw new InvalidOperationException("docId is not an assembly.");

                var comp = ResolveComponent(asm, componentRef);

                doc.ClearSelection2(true);
                if (!comp.Select4(false, null, false))
                    throw new InvalidOperationException("Failed to select component.");

                if (fixedInPlace)
                    asm.FixComponent();
                else
                    asm.UnfixComponent();

                return 0;
            });
        }

        public void AssemblyTranslateComponent(
            string assemblyId,
            string componentRef,   // name OR docId
            double dx, double dy, double dz
        )
        {
            InvokeSW(app =>
            {
                var doc = GetDoc(assemblyId);
                var asm = doc as AssemblyDoc;
                if (asm == null)
                    throw new InvalidOperationException("docId is not an assembly.");

                var comp = ResolveComponent(asm, componentRef);

                var mu = (MathUtility)app.GetMathUtility();

                var cur = comp.Transform2;
                if (cur == null)
                    throw new InvalidOperationException("Component has no transform.");

                var curArr = (double[])cur.ArrayData;

                var tArr = new double[]
                {
                    1,0,0,0,
                    0,1,0,0,
                    0,0,1,0,
                    dx,dy,dz,1
                };

                var delta = (MathTransform)mu.CreateTransform(tArr);
                var newT = (MathTransform)delta.Multiply(cur);

                comp.Transform2 = newT;
                return 0;
            });
        }


        public void AssemblyRotateComponent(
            string assemblyId,
            string componentRef,   // name OR docId
            double rx, double ry, double rz
        )
        {
            InvokeSW(app =>
            {
                var doc = GetDoc(assemblyId);
                var asm = doc as AssemblyDoc;
                if (asm == null)
                    throw new InvalidOperationException("docId is not an assembly.");

                var comp = ResolveComponent(asm, componentRef);

                var mu = (MathUtility)app.GetMathUtility();
                var cur = comp.Transform2;
                if (cur == null)
                    throw new InvalidOperationException("Component has no transform.");

                double cx = Math.Cos(rx), sx = Math.Sin(rx);
                double cy = Math.Cos(ry), sy = Math.Sin(ry);
                double cz = Math.Cos(rz), sz = Math.Sin(rz);

                double r11 = cz * cy;
                double r12 = cz * sy * sx - sz * cx;
                double r13 = cz * sy * cx + sz * sx;

                double r21 = sz * cy;
                double r22 = sz * sy * sx + cz * cx;
                double r23 = sz * sy * cx - cz * sx;

                double r31 = -sy;
                double r32 = cy * sx;
                double r33 = cy * cx;

                var rArr = new double[]
                {
                    r11, r21, r31, 0,
                    r12, r22, r32, 0,
                    r13, r23, r33, 0,
                    0,   0,   0,   1
                };

                var delta = (MathTransform)mu.CreateTransform(rArr);
                var newT = (MathTransform)delta.Multiply(cur);

                comp.Transform2 = newT;
                return 0;
            });
        }

        private Component2 ResolveComponent(AssemblyDoc asm, string compRef)
        {
            // Case 1: compRef is a component name
            var comp = (Component2)asm.GetComponentByName(compRef);
            if (comp != null)
                return comp;

            // Case 2: compRef is a tracked docId
            ModelDoc2 doc;
            if (_docs.TryGetValue(compRef, out doc) && doc != null)
            {
                var path = doc.GetPathName();
                if (!string.IsNullOrWhiteSpace(path))
                {
                    var comps = (object[])asm.GetComponents(false);
                    if (comps != null)
                    {
                        foreach (Component2 c in comps)
                        {
                            if (string.Equals(c.GetPathName(), path, StringComparison.OrdinalIgnoreCase))
                                return c;
                        }
                    }
                }
            }

            throw new InvalidOperationException("Component not found: " + compRef);
        }

        public void SaveDocument(string docId, bool silent = true)
        {
            InvokeSW(app =>
            {
                var doc = GetDoc(docId);
                int errs = 0, warns = 0;

                var opts = silent ? (int)swSaveAsOptions_e.swSaveAsOptions_Silent : 0;
                doc.Save3(opts, ref errs, ref warns);

                if (errs != 0)
                    throw new InvalidOperationException("Save failed. Errors: " + errs + " Warnings: " + warns);

                return 0;
            });
        }

        public void SaveAsDocument(string docId, string outPath, bool silent = true)
        {
            if (string.IsNullOrWhiteSpace(outPath))
                throw new ArgumentException("outPath is required", nameof(outPath));

            InvokeSW(app =>
            {
                var doc = GetDoc(docId);

                int errors = 0;
                int warnings = 0;

                int opts = 0;
                if (silent)
                    opts |= (int)swSaveAsOptions_e.swSaveAsOptions_Silent;

                bool ok = doc.Extension.SaveAs(
                    outPath,
                    (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                    opts,
                    null,
                    ref errors,
                    ref warnings
                );

                if (!ok || errors != 0)
                    throw new InvalidOperationException(
                        $"SaveAs failed. ok={ok}, errors={errors}, warnings={warnings}, path={outPath}"
                    );

                return 0;
            });
        }

        public void RebuildDocument(string docId, bool topOnly = false)
        {
            InvokeSW(app =>
            {
                var doc = GetDoc(docId);
                doc.ForceRebuild3(topOnly);
                return 0;
            });
        }

        public void CloseDocument(string docId, bool save = false, bool silentSave = true)
        {
            InvokeSW(app =>
            {
                var doc = GetDoc(docId);

                if (save)
                    SaveDocument(docId, silentSave); // runs on same STA thread via InvokeSw

                var title = doc.GetTitle();
                app.CloseDoc(title);

                _docs.Remove(docId);
                return 0;
            });
        }

        public string GetRevision()
            => InvokeSW(app => app.RevisionNumber());
        public void Shutdown(bool force = false)
        {
            if (!_connected) return;

            _sta.Invoke(() =>
            {
                if (_app != null)
                {
                    try
                    {
                        if (force || _launchedByUs)
                            _app.ExitApp();
                    }
                    catch { /* ignore */ }

                    try { Marshal.FinalReleaseComObject(_app); } catch { }
                    _app = null;
                }

                _connected = false;
                return 0;
            });
        }

        public void Dispose()
        {
            Shutdown(force: false);   // only closes if we launched it
            _sta.Dispose();
        }
    }
}