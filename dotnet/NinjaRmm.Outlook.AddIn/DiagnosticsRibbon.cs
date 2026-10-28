using System.Runtime.InteropServices;
using System.Windows.Forms;
using NinjaRmm.Outlook.Core;
using Office = Microsoft.Office.Core;

namespace NinjaRmm.Outlook.AddIn
{
    [ComVisible(true)]
    public sealed class DiagnosticsRibbon : Office.IRibbonExtensibility
    {
        public string GetCustomUI(string ribbonId)
        {
            if (ribbonId != "Microsoft.Outlook.Explorer") return null;
            HostDiagnostics.Record("ribbon-requested");
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnLoad'>
  <ribbon><tabs><tab id='NinjaM9Tab' label='NinjaRMM M9'>
    <group id='NinjaM9Diagnostics' label='Development host'>
      <button id='NinjaM9Status' label='Host status' size='large' imageMso='Info' onAction='OnStatus'/>
    </group>
  </tab></tabs></ribbon>
</customUI>";
        }

        public void OnLoad(Office.IRibbonUI ribbon) { HostDiagnostics.Record("ribbon-loaded"); }
        public void OnStatus(Office.IRibbonControl control)
        {
            bool coreReady = ArchivePolicy.Decide(TicketState.Open, true, true) == ArchiveDecision.Keep;
            HostDiagnostics.Record(coreReady ? "core-check-passed" : "core-check-failed");
            MessageBox.Show("VSTO host and Ribbon callback are loaded.\r\nCore reference: " +
                (coreReady ? "OK" : "FAILED") + "\r\n\r\nMail processing and API access are not connected.",
                "NinjaRMM M9", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
