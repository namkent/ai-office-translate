using System;
using System.IO;
using Microsoft.Office.Tools.Excel;

namespace AITranslateExcel
{
    [global::Microsoft.VisualStudio.Tools.Applications.Runtime.StartupObjectAttribute(0)]
    public partial class ThisAddIn : Microsoft.Office.Tools.AddInBase
    {
        private AITranslateCore.AITranslateCore core;

        public ThisAddIn(global::Microsoft.Office.Tools.Excel.ApplicationFactory factory, global::System.IServiceProvider serviceProvider) 
            : base(factory, serviceProvider, "AddIn", "ThisAddIn")
        {
        }

        public global::Microsoft.Office.Interop.Excel.Application Application { get; private set; }

        private static void Log(string message)
        {
            try
            {
                string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string logDir = Path.Combine(appData, "AITranslateAddin");
                if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                string logFile = Path.Combine(logDir, "vsto_log.txt");
                File.AppendAllText(logFile, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}\r\n");
            }
            catch { }
        }

        protected override void Initialize()
        {
            Log("Initialize() entered");
            try
            {
                base.Initialize();
                Log("base.Initialize() completed");
            }
            catch (Exception ex)
            {
                Log($"base.Initialize() failed: {ex}");
            }

            try
            {
                this.Application = this.GetHostItem<global::Microsoft.Office.Interop.Excel.Application>(typeof(global::Microsoft.Office.Interop.Excel.Application), "Application");
                Log($"GetHostItem Application check: {(this.Application != null ? "Not Null" : "Null")}");
            }
            catch (Exception ex)
            {
                Log($"GetHostItem Application failed: {ex}");
            }

            try
            {
                Globals.ThisAddIn = this;
                Log("Globals.ThisAddIn assigned");
            }
            catch (Exception ex)
            {
                Log($"Assigning Globals.ThisAddIn failed: {ex}");
            }
        }

        protected override void FinishInitialization()
        {
            Log("FinishInitialization() entered");
            try
            {
                this.InternalStartup();
                Log("InternalStartup() completed");
                this.OnStartup();
                Log("OnStartup() completed");
            }
            catch (Exception ex)
            {
                Log($"FinishInitialization failed: {ex}");
            }
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Log($"ThisAddIn_Startup entered. Application check: {(this.Application != null ? "Not Null" : "Null")}");
            if (this.Application == null)
            {
                try
                {
                    // Fallback: try to query from HostContext service provider
                    if (this.HostContext != null)
                    {
                        Log("Attempting to get Application from HostContext...");
                        var appType = typeof(global::Microsoft.Office.Interop.Excel.Application);
                        this.Application = this.HostContext.GetService(appType) as global::Microsoft.Office.Interop.Excel.Application;
                        Log($"HostContext Application check: {(this.Application != null ? "Not Null" : "Null")}");
                    }
                }
                catch (Exception ex)
                {
                    Log($"HostContext GetService failed: {ex}");
                }
            }

            if (core == null)
            {
                Log("Startup: Creating core...");
                core = new AITranslateCore.AITranslateCore(this.Application, "Excel");
            }
            else
            {
                Log("Startup: SetApplication on core...");
                core.SetApplication(this.Application);
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            Log("ThisAddIn_Shutdown entered");
            core = null;
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            Log($"CreateRibbonExtensibilityObject entered. Application check: {(this.Application != null ? "Not Null" : "Null")}");
            if (core == null)
            {
                Log("CreateRibbonExtensibilityObject: Creating core...");
                core = new AITranslateCore.AITranslateCore(this.Application, "Excel");
            }
            return core;
        }

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
    }

    internal sealed partial class Globals
    {
        private static ThisAddIn _ThisAddIn;
        internal static ThisAddIn ThisAddIn
        {
            get { return _ThisAddIn; }
            set
            {
                if (_ThisAddIn == null) _ThisAddIn = value;
                else throw new System.NotSupportedException();
            }
        }
    }
}
