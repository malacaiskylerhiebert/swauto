using System;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SWAutomation.Core
{
    public static class SWDocuments
    {
        /// <summary>
        /// Opens an assembly document from a file path.
        /// </summary>
        /// <param name="app">Active SolidWorks application.</param>
        /// <param name="assemblyPath">Full path to .sldasm file.</param>
        /// <returns>True if opened successfully.</returns>
        public static bool OpenAssembly(SldWorks app, string assemblyPath)
        {
            if (app == null)
                throw new ArgumentNullException(nameof(app));

            if (string.IsNullOrWhiteSpace(assemblyPath))
                throw new ArgumentException("Assembly path is null or empty.");

            int errors = 0;
            int warnings = 0;

            var model = app.OpenDoc6(
                assemblyPath,
                (int)swDocumentTypes_e.swDocASSEMBLY,
                (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                "",
                ref errors,
                ref warnings
            );

            return model != null && errors == 0;
        }
    }
}
