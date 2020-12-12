
namespace MathTexPPT {
    partial class Designer : Microsoft.Office.Tools.Ribbon.RibbonBase {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Designer()
            : base(Globals.Factory.GetRibbonFactory()) {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing) {
            if(disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent() {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.butInsert = this.Factory.CreateRibbonButton();
            this.butEdit = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabHome";
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabHome";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.butInsert);
            this.group1.Items.Add(this.butEdit);
            this.group1.Label = "MathTex";
            this.group1.Name = "group1";
            // 
            // butInsert
            // 
            this.butInsert.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.butInsert.Image = global::MathTexPPT.Properties.Resources.sigma;
            this.butInsert.Label = "添加公式";
            this.butInsert.Name = "butInsert";
            this.butInsert.ShowImage = true;
            this.butInsert.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.butInsert_Click);
            // 
            // butEdit
            // 
            this.butEdit.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.butEdit.Image = global::MathTexPPT.Properties.Resources.edit;
            this.butEdit.Label = "编辑公式";
            this.butEdit.Name = "butEdit";
            this.butEdit.ShowImage = true;
            this.butEdit.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.butEdit_Click);
            // 
            // Designer
            // 
            this.Name = "Designer";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Designer_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton butInsert;
        private Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton butEdit;
    }

    partial class ThisRibbonCollection {
        internal Designer Designer {
            get { return this.GetRibbon<Designer>(); }
        }
    }
}
