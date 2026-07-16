using System;
using Microsoft.Office.Tools.Word;

namespace AITranslateWord
{
    public partial class ThisAddIn : Microsoft.Office.Tools.Word.WordAddInBase
    {
        private AITranslateCore.AITranslateCore core;

        public ThisAddIn(global::Microsoft.Office.Tools.Word.ApplicationFactory factory, global::System.IServiceProvider serviceProvider) 
            : base(factory, serviceProvider, "AddIn", "ThisAddIn")
        {
        }

        protected override void Initialize()
        {
            base.Initialize();
            Globals.ThisAddIn = this;
        }

        protected override void FinishInitialization()
        {
            base.FinishInitialization();
            this.InternalStartup();
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            core = new AITranslateCore.AITranslateCore(this.Application, "Word");
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            core = null;
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            if (core == null)
            {
                core = new AITranslateCore.AITranslateCore(this.Application, "Word");
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
