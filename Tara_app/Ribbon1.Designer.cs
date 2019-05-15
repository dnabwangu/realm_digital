using Microsoft.Office.Interop.Excel;

namespace Tara_app
{
    partial class AI_Spark : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public AI_Spark()
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
            this.data = this.Factory.CreateRibbonGroup();
            this.view_form = this.Factory.CreateRibbonButton();
            this.view = this.Factory.CreateRibbonGroup();
            this.ask_tara_ = this.Factory.CreateRibbonButton();
            this.illuminate_ = this.Factory.CreateRibbonButton();
            this.comparison = this.Factory.CreateRibbonGroup();
            this.compare_dropdown = this.Factory.CreateRibbonDropDown();
            this.compare_button = this.Factory.CreateRibbonButton();
            this.authentication = this.Factory.CreateRibbonGroup();
            this.userLabel = this.Factory.CreateRibbonLabel();
            this.loginButton = this.Factory.CreateRibbonButton();
            this.logoutButton = this.Factory.CreateRibbonButton();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.directorySearcher1 = new System.DirectoryServices.DirectorySearcher();
            this.model_governance = this.Factory.CreateRibbonGroup();
            this.dropDown3 = this.Factory.CreateRibbonDropDown();
            this.dropDown4 = this.Factory.CreateRibbonDropDown();
            this.dropDown5 = this.Factory.CreateRibbonDropDown();
            this.tab1.SuspendLayout();
            this.data.SuspendLayout();
            this.view.SuspendLayout();
            this.comparison.SuspendLayout();
            this.authentication.SuspendLayout();
            this.model_governance.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.data);
            this.tab1.Groups.Add(this.view);
            this.tab1.Groups.Add(this.comparison);
            this.tab1.Groups.Add(this.authentication);
            this.tab1.Label = "ai-Spark";
            this.tab1.Name = "tab1";
            // 
            // data
            // 
            this.data.Items.Add(this.view_form);
            this.data.Name = "data";
            // 
            // view_form
            // 
            this.view_form.Label = "Intex";
            this.view_form.Name = "view_form";
            this.view_form.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.view_form_Click);
            // 
            // view
            // 
            this.view.Items.Add(this.ask_tara_);
            this.view.Items.Add(this.illuminate_);
            this.view.Name = "view";
            // 
            // ask_tara_
            // 
            this.ask_tara_.Description = "load TARA risk ratings";
            this.ask_tara_.Image = global::Tara_app.Properties.Resources.favicon__1_;
            this.ask_tara_.Label = "ask tara";
            this.ask_tara_.Name = "ask_tara_";
            this.ask_tara_.ShowImage = true;
            this.ask_tara_.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Ask_tara__Click);
            // 
            // illuminate_
            // 
            this.illuminate_.Image = global::Tara_app.Properties.Resources.favicon__1_;
            this.illuminate_.Label = "illuminate";
            this.illuminate_.Name = "illuminate_";
            this.illuminate_.ScreenTip = "TARA sensitvity analysis";
            this.illuminate_.ShowImage = true;
            this.illuminate_.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Illuminate__Click);
            // 
            // comparison
            // 
            this.comparison.Items.Add(this.compare_dropdown);
            this.comparison.Items.Add(this.compare_button);
            this.comparison.Name = "comparison";
            // 
            // compare_dropdown
            // 
            this.compare_dropdown.Label = " comparison";
            this.compare_dropdown.Name = "compare_dropdown";
            // 
            // compare_button
            // 
            this.compare_button.Description = "compare";
            this.compare_button.Image = global::Tara_app.Properties.Resources.favicon__1_;
            this.compare_button.Label = "comparison";
            this.compare_button.Name = "compare_button";
            this.compare_button.ShowImage = true;
            this.compare_button.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Compare_button_Click);
            // 
            // authentication
            // 
            this.authentication.Items.Add(this.userLabel);
            this.authentication.Items.Add(this.loginButton);
            this.authentication.Items.Add(this.logoutButton);
            this.authentication.Name = "authentication";
            // 
            // userLabel
            // 
            this.userLabel.Label = "user";
            this.userLabel.Name = "userLabel";
            // 
            // loginButton
            // 
            this.loginButton.Label = "Sign in";
            this.loginButton.Name = "loginButton";
            this.loginButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.LoginButton_Click);
            // 
            // logoutButton
            // 
            this.logoutButton.Label = "Sign out";
            this.logoutButton.Name = "logoutButton";
            this.logoutButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.LogoutButton_Click);
            // 
            // directorySearcher1
            // 
            this.directorySearcher1.ClientTimeout = System.TimeSpan.Parse("-00:00:01");
            this.directorySearcher1.ServerPageTimeLimit = System.TimeSpan.Parse("-00:00:01");
            this.directorySearcher1.ServerTimeLimit = System.TimeSpan.Parse("-00:00:01");
            // 
            // model_governance
            // 
            this.model_governance.Items.Add(this.dropDown3);
            this.model_governance.Items.Add(this.dropDown4);
            this.model_governance.Items.Add(this.dropDown5);
            this.model_governance.Label = "model governance";
            this.model_governance.Name = "model_governance";
            // 
            // dropDown3
            // 
            this.dropDown3.Label = "analyst assumptions";
            this.dropDown3.Name = "dropDown3";
            this.dropDown3.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DropDown3_SelectionChanged);
            // 
            // dropDown4
            // 
            this.dropDown4.Label = "TARA version";
            this.dropDown4.Name = "dropDown4";
            this.dropDown4.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DropDown4_SelectionChanged);
            // 
            // dropDown5
            // 
            this.dropDown5.Label = "save analysis";
            this.dropDown5.Name = "dropDown5";
            this.dropDown5.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DropDown5_SelectionChanged);
            // 
            // AI_Spark
            // 
            this.Name = "AI_Spark";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.data.ResumeLayout(false);
            this.data.PerformLayout();
            this.view.ResumeLayout(false);
            this.view.PerformLayout();
            this.comparison.ResumeLayout(false);
            this.comparison.PerformLayout();
            this.authentication.ResumeLayout(false);
            this.authentication.PerformLayout();
            this.model_governance.ResumeLayout(false);
            this.model_governance.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.DirectoryServices.DirectorySearcher directorySearcher1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup view;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup model_governance;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDown3;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDown4;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDown5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ask_tara_;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton illuminate_;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton compare_button;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown compare_dropdown;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup authentication;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton loginButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton logoutButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel userLabel;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup data;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton view_form;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup comparison;
    }

    partial class ThisRibbonCollection
    {
        internal AI_Spark Ribbon1
        {
            get { return this.GetRibbon<AI_Spark>(); }
        }
    }
}
