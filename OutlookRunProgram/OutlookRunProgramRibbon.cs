using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;

namespace OutlookRunProgram
{
	public partial class OutlookRunProgramRibbon
	{
		private void OutlookRunProgramRibbon_Load(object sender, RibbonUIEventArgs e)
		{
			Globals.ThisAddIn.Auto(toggleButtonAuto.Checked);

		}

		private void toggleButton1_Click(object sender, RibbonControlEventArgs e)
		{
			Globals.ThisAddIn.Auto(toggleButtonAuto.Checked);
		}

		private void buttonRun_Click(object sender, RibbonControlEventArgs e)
		{
			var activeExplorer = Globals.ThisAddIn.Application.ActiveExplorer();
			foreach (var selectedItem in activeExplorer.Selection)
			{
				var mailItem = selectedItem as MailItem;
				if (mailItem != null)
				{
					Globals.ThisAddIn.RunRules(selectedItem as MailItem);
				}
			}

		}

		private void buttonReload_Click(object sender, RibbonControlEventArgs e)
		{
			Globals.ThisAddIn.Reload();
		}
	}
}
