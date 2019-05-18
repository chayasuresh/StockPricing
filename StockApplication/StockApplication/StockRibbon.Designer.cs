namespace StockApplication
{
    partial class StockRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public StockRibbon()
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
            this.btnAllStocks = this.Factory.CreateRibbonButton();
            this.btnIndvStocks = this.Factory.CreateRibbonButton();
            this.btnClear = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "StocksAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnAllStocks);
            this.group1.Items.Add(this.btnIndvStocks);
            this.group1.Items.Add(this.btnClear);
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // btnAllStocks
            // 
            this.btnAllStocks.Label = "Get All Stock Prices";
            this.btnAllStocks.Name = "btnAllStocks";
            this.btnAllStocks.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAllStocks_Click);
            // 
            // btnIndvStocks
            // 
            this.btnIndvStocks.Label = "Get Selected Stock Value";
            this.btnIndvStocks.Name = "btnIndvStocks";
            this.btnIndvStocks.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnIndvStocks_Click);
            // 
            // btnClear
            // 
            this.btnClear.Label = "Clear All Stock Prices";
            this.btnClear.Name = "btnClear";
            this.btnClear.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnClear_Click);
            // 
            // StockRibbon
            // 
            this.Name = "StockRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.StockRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAllStocks;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnIndvStocks;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnClear;
    }

    partial class ThisRibbonCollection
    {
        internal StockRibbon StockRibbon
        {
            get { return this.GetRibbon<StockRibbon>(); }
        }
    }
}
