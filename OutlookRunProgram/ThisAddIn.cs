using System;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.Outlook;

namespace OutlookRunProgram
{
	public partial class ThisAddIn
	{
		Ruler ruler = new Ruler();
		bool enabled = false;

		NameSpace ns;

		private void ThisAddIn_Startup(object sender, EventArgs e)
		{
			Reload();

			ns = this.Application.GetNamespace("MAPI");
			this.Application.NewMailEx += Application_NewMailEx;
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

			bool success = ruler.ReadRules(Environment.GetFolderPath(
				Environment.SpecialFolder.LocalApplicationData) + "\\bcieszko\\OutlookRunProgram"
				);
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
