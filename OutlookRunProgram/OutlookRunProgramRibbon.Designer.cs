namespace OutlookRunProgram
{
	partial class OutlookRunProgramRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		public OutlookRunProgramRibbon()
			: base(Globals.Factory.GetRibbonFactory())
		{
			InitializeComponent();
		}

		/// <summary> 
		/// Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Component Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.tab1 = this.Factory.CreateRibbonTab();
			this.group1 = this.Factory.CreateRibbonGroup();
			this.toggleButtonAuto = this.Factory.CreateRibbonToggleButton();
			this.buttonRun = this.Factory.CreateRibbonButton();
			this.buttonReload = this.Factory.CreateRibbonButton();
			this.tab1.SuspendLayout();
			this.group1.SuspendLayout();
			this.SuspendLayout();
			// 
			// tab1
			// 
			this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
			this.tab1.Groups.Add(this.group1);
			this.tab1.Label = "TabAddIns";
			this.tab1.Name = "tab1";
			// 
			// group1
			// 
			this.group1.Items.Add(this.toggleButtonAuto);
			this.group1.Items.Add(this.buttonReload);
			this.group1.Items.Add(this.buttonRun);
			this.group1.Label = "bORP";
			this.group1.Name = "group1";
			// 
			// toggleButtonAuto
			// 
			this.toggleButtonAuto.Label = "Apply on new mail";
			this.toggleButtonAuto.Name = "toggleButtonAuto";
			this.toggleButtonAuto.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButton1_Click);
			// 
			// buttonRun
			// 
			this.buttonRun.Label = "Run on current mail";
			this.buttonRun.Name = "buttonRun";
			this.buttonRun.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonRun_Click);
			// 
			// buttonReload
			// 
			this.buttonReload.Label = "Reload";
			this.buttonReload.Name = "buttonReload";
			this.buttonReload.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonReload_Click);
			// 
			// OutlookRunProgramRibbon
			// 
			this.Name = "OutlookRunProgramRibbon";
			this.RibbonType = "Microsoft.Outlook.Explorer, Microsoft.Outlook.Mail.Read";
			this.Tabs.Add(this.tab1);
			this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.OutlookRunProgramRibbon_Load);
			this.tab1.ResumeLayout(false);
			this.tab1.PerformLayout();
			this.group1.ResumeLayout(false);
			this.group1.PerformLayout();
			this.ResumeLayout(false);

		}

		#endregion

		internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
		internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
		internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButtonAuto;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonRun;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonReload;
	}

	partial class ThisRibbonCollection
	{
		internal OutlookRunProgramRibbon OutlookRunProgramRibbon
		{
			get { return this.GetRibbon<OutlookRunProgramRibbon>(); }
		}
	}
}
