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

		private void ToggleButton1_Click(object sender, RibbonControlEventArgs e)
		{
			Globals.ThisAddIn.Auto(toggleButtonAuto.Checked);
		}

		private void ButtonRun_Click(object sender, RibbonControlEventArgs e)
		{
			var activeExplorer = Globals.ThisAddIn.Application.ActiveExplorer();
			foreach (var selectedItem in activeExplorer.Selection)
			{
				if (selectedItem is MailItem mailItem)
				{
					Globals.ThisAddIn.RunRules(selectedItem as MailItem);
				}
			}
		}

		private void ButtonReload_Click(object sender, RibbonControlEventArgs e)
		{
			Globals.ThisAddIn.Reload();
		}

		private void ButtonLog_Click(object sender, RibbonControlEventArgs e)
		{
			Globals.ThisAddIn.GetLogger().Show();
		}
	}
}
