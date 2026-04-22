namespace FunctionBox
{
    partial class FunctionBoxRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public FunctionBoxRibbon()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FunctionBoxRibbon));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.核对工具 = this.Factory.CreateRibbonGroup();
            this.btnValidateVerticalTop = this.Factory.CreateRibbonButton();
            this.btnValidateHorizontal = this.Factory.CreateRibbonButton();
            this.btnValidateVerticalDown = this.Factory.CreateRibbonButton();
            this.btnClearSelectionBackground = this.Factory.CreateRibbonButton();
            this.btnClearDocumentBackground = this.Factory.CreateRibbonButton();
            this.btnCheckSumDebug = this.Factory.CreateRibbonToggleButton();
            this.文字处理 = this.Factory.CreateRibbonGroup();
            this.btnAddThousand = this.Factory.CreateRibbonButton();
            this.btnBracketConvert = this.Factory.CreateRibbonButton();
            this.bthNegativeFormat = this.Factory.CreateRibbonButton();
            this.收纳箱 = this.Factory.CreateRibbonGroup();
            this.btnToolBox = this.Factory.CreateRibbonButton();
            this.btnToolList = this.Factory.CreateRibbonDropDown();
            this.btnExecuteVba = this.Factory.CreateRibbonButton();
            this.其他 = this.Factory.CreateRibbonGroup();
            this.btnQuestion = this.Factory.CreateRibbonButton();
            this.btnUpdate = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.核对工具.SuspendLayout();
            this.文字处理.SuspendLayout();
            this.收纳箱.SuspendLayout();
            this.其他.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.核对工具);
            this.tab1.Groups.Add(this.文字处理);
            this.tab1.Groups.Add(this.收纳箱);
            this.tab1.Groups.Add(this.其他);
            this.tab1.Label = "便利店";
            this.tab1.Name = "tab1";
            // 
            // 核对工具
            // 
            this.核对工具.Items.Add(this.btnValidateVerticalTop);
            this.核对工具.Items.Add(this.btnValidateHorizontal);
            this.核对工具.Items.Add(this.btnValidateVerticalDown);
            this.核对工具.Items.Add(this.btnClearSelectionBackground);
            this.核对工具.Items.Add(this.btnClearDocumentBackground);
            this.核对工具.Items.Add(this.btnCheckSumDebug);
            this.核对工具.Label = "核对工具";
            this.核对工具.Name = "核对工具";
            // 
            // btnValidateVerticalTop
            // 
            this.btnValidateVerticalTop.Image = ((System.Drawing.Image)(resources.GetObject("btnValidateVerticalTop.Image")));
            this.btnValidateVerticalTop.Label = "竖向加总检查";
            this.btnValidateVerticalTop.Name = "btnValidateVerticalTop";
            this.btnValidateVerticalTop.ScreenTip = "选中表格数据检验竖向加总-自下向上加总";
            this.btnValidateVerticalTop.ShowImage = true;
            this.btnValidateVerticalTop.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnValidateVerticalTop_Click);
            // 
            // btnValidateHorizontal
            // 
            this.btnValidateHorizontal.Image = ((System.Drawing.Image)(resources.GetObject("btnValidateHorizontal.Image")));
            this.btnValidateHorizontal.Label = "横向加总检查";
            this.btnValidateHorizontal.Name = "btnValidateHorizontal";
            this.btnValidateHorizontal.ScreenTip = "选中表格数据检验横向加总";
            this.btnValidateHorizontal.ShowImage = true;
            this.btnValidateHorizontal.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnValidateHorizontal_Click);
            // 
            // btnValidateVerticalDown
            // 
            this.btnValidateVerticalDown.Image = ((System.Drawing.Image)(resources.GetObject("btnValidateVerticalDown.Image")));
            this.btnValidateVerticalDown.Label = "竖向加总检查";
            this.btnValidateVerticalDown.Name = "btnValidateVerticalDown";
            this.btnValidateVerticalDown.ScreenTip = "选中表格数据检验竖向加总-自上向下加总";
            this.btnValidateVerticalDown.ShowImage = true;
            this.btnValidateVerticalDown.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnValidateVerticalDown_Click);
            // 
            // btnClearSelectionBackground
            // 
            this.btnClearSelectionBackground.Label = "清除选中高亮";
            this.btnClearSelectionBackground.Name = "btnClearSelectionBackground";
            this.btnClearSelectionBackground.SuperTip = "清除选中背景色及高亮";
            this.btnClearSelectionBackground.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnClearSelectionBackground_Click);
            // 
            // btnClearDocumentBackground
            // 
            this.btnClearDocumentBackground.Label = "清除全文高亮";
            this.btnClearDocumentBackground.Name = "btnClearDocumentBackground";
            this.btnClearDocumentBackground.SuperTip = "清除全文背景色及高亮";
            this.btnClearDocumentBackground.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnClearDocumentBackground_Click);
            // 
            // btnCheckSumDebug
            // 
            this.btnCheckSumDebug.Label = "加总调试模式";
            this.btnCheckSumDebug.Name = "btnCheckSumDebug";
            this.btnCheckSumDebug.ScreenTip = "开启后会输出加总检查调试信息";
            this.btnCheckSumDebug.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCheckSumDebug_Click);
            // 
            // 文字处理
            // 
            this.文字处理.Items.Add(this.btnAddThousand);
            this.文字处理.Items.Add(this.btnBracketConvert);
            this.文字处理.Items.Add(this.bthNegativeFormat);
            this.文字处理.Label = "文字处理";
            this.文字处理.Name = "文字处理";
            // 
            // btnAddThousand
            // 
            this.btnAddThousand.Description = "添加千分符";
            this.btnAddThousand.Label = "添加千分符号";
            this.btnAddThousand.Name = "btnAddThousand";
            this.btnAddThousand.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAddThousand_Click);
            // 
            // btnBracketConvert
            // 
            this.btnBracketConvert.Description = "中英括号转换";
            this.btnBracketConvert.Label = "中英括号转换";
            this.btnBracketConvert.Name = "btnBracketConvert";
            this.btnBracketConvert.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnBracketConvert_Click);
            // 
            // bthNegativeFormat
            // 
            this.bthNegativeFormat.Description = "负号格式转换";
            this.bthNegativeFormat.Label = "负号格式转换";
            this.bthNegativeFormat.Name = "bthNegativeFormat";
            this.bthNegativeFormat.ScreenTip = "负号及括号互转";
            this.bthNegativeFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bthNegativeFormat_Click);
            // 
            // 收纳箱
            // 
            this.收纳箱.Items.Add(this.btnToolBox);
            this.收纳箱.Items.Add(this.btnToolList);
            this.收纳箱.Items.Add(this.btnExecuteVba);
            this.收纳箱.Label = "VBA工具箱";
            this.收纳箱.Name = "收纳箱";
            // 
            // btnToolBox
            // 
            this.btnToolBox.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnToolBox.Image = global::FunctionBox.Properties.Resources.app;
            this.btnToolBox.Label = "VBA工具箱";
            this.btnToolBox.Name = "btnToolBox";
            this.btnToolBox.ShowImage = true;
            this.btnToolBox.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnToolBox_Click);
            // 
            // btnToolList
            // 
            this.btnToolList.Label = " ";
            this.btnToolList.Name = "btnToolList";
            // 
            // btnExecuteVba
            // 
            this.btnExecuteVba.Label = "运行程序";
            this.btnExecuteVba.Name = "btnExecuteVba";
            this.btnExecuteVba.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExecuteVba_Click);
            // 
            // 其他
            // 
            this.其他.Items.Add(this.btnQuestion);
            this.其他.Items.Add(this.btnUpdate);
            this.其他.Label = "帮助";
            this.其他.Name = "其他";
            // 
            // btnQuestion
            // 
            this.btnQuestion.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnQuestion.Image = global::FunctionBox.Properties.Resources.帮助_问号;
            this.btnQuestion.Label = "Help";
            this.btnQuestion.Name = "btnQuestion";
            this.btnQuestion.ScreenTip = "点我没用，我也不知道。";
            this.btnQuestion.ShowImage = true;
            this.btnQuestion.Tag = "";
            this.btnQuestion.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnQuestion_Click);
            // 
            // btnUpdate
            // 
            this.btnUpdate.Label = "检查更新";
            this.btnUpdate.Name = "btnUpdate";
            this.btnUpdate.OfficeImageId = "Refresh";
            this.btnUpdate.ShowImage = true;
            this.btnUpdate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdate_Click);
            // 
            // FunctionBoxRibbon
            // 
            this.Name = "FunctionBoxRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.核对工具.ResumeLayout(false);
            this.核对工具.PerformLayout();
            this.文字处理.ResumeLayout(false);
            this.文字处理.PerformLayout();
            this.收纳箱.ResumeLayout(false);
            this.收纳箱.PerformLayout();
            this.其他.ResumeLayout(false);
            this.其他.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup 核对工具;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnValidateHorizontal;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnValidateVerticalDown;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup 其他;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnClearDocumentBackground;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnValidateVerticalTop;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnClearSelectionBackground;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnCheckSumDebug;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnQuestion;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdate;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup 收纳箱;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnToolBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown btnToolList;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExecuteVba;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnBracketConvert;
        public Microsoft.Office.Tools.Ribbon.RibbonButton btnAddThousand;
        public Microsoft.Office.Tools.Ribbon.RibbonButton bthNegativeFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup 文字处理;
    }


}
