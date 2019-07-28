using System;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.Outlook;

namespace OutlookRunProgram
{
	public partial class ThisAddIn
	{
		Ruler ruler = new Ruler();

		Outlook.NameSpace ns;

		private void ThisAddIn_Startup(object sender, System.EventArgs e)
		{
			bool success = ruler.ReadRules(Environment.GetFolderPath(
				Environment.SpecialFolder.LocalApplicationData) + "\\bcieszko\\OutlookRunProgram"
				);

			if (!success)
			{
				// no rules, go to void
				return;
			}

			// 1. read from xml
			ReadFromXml();
			/*
			 * (?!) match nothing
			 * (?=) match everything
			 * <entry> (?!)
			 * <regex_subject>
			 * <regex_body>
			 * <regex_mail>
			 * <actions>
			 * <run>
			 *
			 */


			this.Application.NewMailEx += Application_NewMailEx;
			ns = this.Application.GetNamespace("MAPI");
		}

		private void ReadFromXml()
		{
			throw new NotImplementedException();
		}

		private void Application_NewMailEx(string EntryIDCollection)
		{
			MailItem item = ns.GetItemFromID(EntryIDCollection);
			ruler.ApplyRules(item);
		}

		private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
		{
			// Note: Outlook no longer raises this event. If you have code that
			//    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
		}

		#region VSTO generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InternalStartup()
		{
			this.Startup += new System.EventHandler(ThisAddIn_Startup);
			this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
		}

		#endregion
	}
}
