namespace WordPlusAndMinus
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.tabPlusAndMinus = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.buttonGenerate = this.Factory.CreateRibbonButton();
            this.tabPlusAndMinus.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabPlusAndMinus
            // 
            this.tabPlusAndMinus.Groups.Add(this.group1);
            this.tabPlusAndMinus.Label = "加减法生成";
            this.tabPlusAndMinus.Name = "tabPlusAndMinus";
            // 
            // group1
            // 
            this.group1.Items.Add(this.buttonGenerate);
            this.group1.Label = "生成规则";
            this.group1.Name = "group1";
            // 
            // buttonGenerate
            // 
            this.buttonGenerate.Label = "生成";
            this.buttonGenerate.Name = "buttonGenerate";
            this.buttonGenerate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonGenerate_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tabPlusAndMinus);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tabPlusAndMinus.ResumeLayout(false);
            this.tabPlusAndMinus.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabPlusAndMinus;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonGenerate;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
