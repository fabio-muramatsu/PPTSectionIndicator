namespace PPT_Section_Indicator
{
    partial class MainRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MainRibbon()
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
            this.slideMarkerCheckBox = this.Factory.CreateRibbonCheckBox();
            this.slideRangeEditBox = this.Factory.CreateRibbonEditBox();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.cleanPresentationButton = this.Factory.CreateRibbonButton();
            this.startButton = this.Factory.CreateRibbonButton();
            this.stepOneNextButton = this.Factory.CreateRibbonButton();
            this.stepOneAboutButton = this.Factory.CreateRibbonButton();
            this.stepTwoDoneButton = this.Factory.CreateRibbonButton();
            this.stepTwoAboutButton = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Label = "Section Indicator";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.slideMarkerCheckBox);
            this.group1.Items.Add(this.slideRangeEditBox);
            this.group1.Items.Add(this.cleanPresentationButton);
            this.group1.Items.Add(this.startButton);
            this.group1.Label = "Settings";
            this.group1.Name = "group1";
            // 
            // slideMarkerCheckBox
            // 
            this.slideMarkerCheckBox.Label = "Include slide markers";
            this.slideMarkerCheckBox.Name = "slideMarkerCheckBox";
            this.slideMarkerCheckBox.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SlideMarkerCheckBox_Click);
            // 
            // slideRangeEditBox
            // 
            this.slideRangeEditBox.Label = "Slides to edit:";
            this.slideRangeEditBox.Name = "slideRangeEditBox";
            this.slideRangeEditBox.ScreenTip = "Slides to edit";
            this.slideRangeEditBox.SizeString = "WWWWWWWWWW";
            this.slideRangeEditBox.SuperTip = "Specify the slides where section progress indicators are to be inserted. Separate" +
    " pages or ranges with \";\" and use \"-\" to indicate page ranges.";
            this.slideRangeEditBox.Text = null;
            // 
            // group2
            // 
            this.group2.Items.Add(this.stepOneNextButton);
            this.group2.Items.Add(this.stepOneAboutButton);
            this.group2.Label = "Step 1";
            this.group2.Name = "group2";
            // 
            // group3
            // 
            this.group3.Items.Add(this.stepTwoDoneButton);
            this.group3.Items.Add(this.stepTwoAboutButton);
            this.group3.Label = "Step 2";
            this.group3.Name = "group3";
            // 
            // cleanPresentationButton
            // 
            this.cleanPresentationButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.cleanPresentationButton.Image = global::PPT_Section_Indicator.Properties.Resources.cancel;
            this.cleanPresentationButton.Label = "Cleanup";
            this.cleanPresentationButton.Name = "cleanPresentationButton";
            this.cleanPresentationButton.ShowImage = true;
            this.cleanPresentationButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CleanupButton_Click);
            // 
            // startButton
            // 
            this.startButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.startButton.Image = global::PPT_Section_Indicator.Properties.Resources.start;
            this.startButton.Label = "Start";
            this.startButton.Name = "startButton";
            this.startButton.ShowImage = true;
            this.startButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.StartButton_Click);
            // 
            // stepOneNextButton
            // 
            this.stepOneNextButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.stepOneNextButton.Image = global::PPT_Section_Indicator.Properties.Resources.next;
            this.stepOneNextButton.Label = "Next";
            this.stepOneNextButton.Name = "stepOneNextButton";
            this.stepOneNextButton.ShowImage = true;
            this.stepOneNextButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.StepOneNextButton_Click);
            // 
            // stepOneAboutButton
            // 
            this.stepOneAboutButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.stepOneAboutButton.Image = global::PPT_Section_Indicator.Properties.Resources.info;
            this.stepOneAboutButton.Label = "About this step";
            this.stepOneAboutButton.Name = "stepOneAboutButton";
            this.stepOneAboutButton.ShowImage = true;
            // 
            // stepTwoDoneButton
            // 
            this.stepTwoDoneButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.stepTwoDoneButton.Image = global::PPT_Section_Indicator.Properties.Resources.done;
            this.stepTwoDoneButton.Label = "Done";
            this.stepTwoDoneButton.Name = "stepTwoDoneButton";
            this.stepTwoDoneButton.ShowImage = true;
            this.stepTwoDoneButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.StepTwoDoneButton_Click);
            // 
            // stepTwoAboutButton
            // 
            this.stepTwoAboutButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.stepTwoAboutButton.Image = global::PPT_Section_Indicator.Properties.Resources.info;
            this.stepTwoAboutButton.Label = "About this step";
            this.stepTwoAboutButton.Name = "stepTwoAboutButton";
            this.stepTwoAboutButton.ShowImage = true;
            // 
            // MainRibbon
            // 
            this.Name = "MainRibbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MainRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox slideMarkerCheckBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox slideRangeEditBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton startButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton stepOneNextButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton stepOneAboutButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton stepTwoDoneButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton stepTwoAboutButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton cleanPresentationButton;
    }

    partial class ThisRibbonCollection
    {
        internal MainRibbon MainRibbon
        {
            get { return this.GetRibbon<MainRibbon>(); }
        }
    }
}
