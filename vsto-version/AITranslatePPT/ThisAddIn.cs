using System;
using Microsoft.Office.Tools;

namespace AITranslatePPT
{
    [global::Microsoft.VisualStudio.Tools.Applications.Runtime.StartupObjectAttribute(0)]
    public partial class ThisAddIn : Microsoft.Office.Tools.AddInBase
    {
        private AITranslateCore.AITranslateCore core;

        public ThisAddIn(global::Microsoft.Office.Tools.Factory factory, global::System.IServiceProvider serviceProvider) 
            : base(factory, serviceProvider, "AddIn", "ThisAddIn")
        {
        }

        public global::Microsoft.Office.Interop.PowerPoint.Application Application { get; private set; }

        protected override void Initialize()
        {
            base.Initialize();
            this.Application = this.GetHostItem<global::Microsoft.Office.Interop.PowerPoint.Application>(typeof(global::Microsoft.Office.Interop.PowerPoint.Application), "Application");
            Globals.ThisAddIn = this;
        }

        protected override void FinishInitialization()
        {
            this.InternalStartup();
            this.OnStartup();
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            if (core == null)
            {
                core = new AITranslateCore.AITranslateCore(this.Application, "PPT");
            }
            else
            {
                core.SetApplication(this.Application);
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            core = null;
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            if (core == null)
            {
                core = new AITranslateCore.AITranslateCore(this.Application, "PPT");
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
