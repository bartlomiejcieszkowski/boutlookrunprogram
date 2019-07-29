using System;
using Microsoft.Office.Interop.Outlook;
using bLogger;
using System.IO;

namespace OutlookRunProgram
{
	public partial class ThisAddIn
	{
		private readonly Ruler ruler = new Ruler();
		private bool enabled = false;
		private Logger logger;

		private NameSpace ns;

		public Logger GetLogger() { return logger; }

		private string GetPluginDirectory()
		{
			return Environment.GetFolderPath(
				Environment.SpecialFolder.LocalApplicationData) + "\\bcieszko\\OutlookRunProgram";
		}

		private void ThisAddIn_Startup(object sender, EventArgs e)
		{
			if (!Directory.Exists(GetPluginDirectory()))
			{
				Directory.CreateDirectory(GetPluginDirectory());
			}

			logger = new Logger(GetPluginDirectory() + "\\bOutlookRunProgram.log");
			Reload();

			ns = Application.GetNamespace("MAPI");
			Application.NewMailEx += Application_NewMailEx;
		}

		internal void Auto(bool @checked)
		{
			enabled = @checked;
		}

		private void Application_NewMailEx(string EntryIDCollection)
		{
			if (enabled)
			{
				MailItem item = ns.GetItemFromID(EntryIDCollection);
				ruler.ApplyRules(item);
			}
		}

		internal void RunRules(MailItem mailItem)
		{
			ruler.ApplyRules(mailItem);
		}

		private void ThisAddIn_Shutdown(object sender, EventArgs e)
		{
			// Note: Outlook no longer raises this event. If you have code that
			//    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
		}

		internal void Reload()
		{
			ruler.ClearRules();

			bool success = ruler.ReadRules(GetPluginDirectory());
			logger.Log($"Load succes? {success}");
		}

		#region VSTO generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InternalStartup()
		{
			this.Startup += ThisAddIn_Startup;
			this.Shutdown += ThisAddIn_Shutdown;
		}

		#endregion
	}
}
