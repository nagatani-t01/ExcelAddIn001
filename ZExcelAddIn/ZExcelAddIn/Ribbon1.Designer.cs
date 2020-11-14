namespace ZExcelAddIn
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージ リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region コンポーネント デザイナーで生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// このメソッドの内容をコード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.button1_1 = this.Factory.CreateRibbonButton();
            this.button1_2 = this.Factory.CreateRibbonButton();
            this.button1_3 = this.Factory.CreateRibbonButton();
            this.button1_4 = this.Factory.CreateRibbonButton();
            this.button1_5 = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.button2_1 = this.Factory.CreateRibbonButton();
            this.zoom1_5 = this.Factory.CreateRibbonEditBox();
            this.cursor1_5 = this.Factory.CreateRibbonEditBox();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.button1_6 = this.Factory.CreateRibbonButton();
            this.button1_7 = this.Factory.CreateRibbonButton();
            this.editBox1_7 = this.Factory.CreateRibbonEditBox();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.button1_8 = this.Factory.CreateRibbonButton();
            this.editBox1_8 = this.Factory.CreateRibbonEditBox();
            this.button1_9 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.group4.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group4);
            this.tab1.Label = "Zツール";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.button1_1);
            this.group1.Items.Add(this.button1_2);
            this.group1.Items.Add(this.button1_3);
            this.group1.Items.Add(this.button1_4);
            this.group1.Items.Add(this.button1_5);
            this.group1.Label = "グループ1";
            this.group1.Name = "group1";
            // 
            // button1_1
            // 
            this.button1_1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1_1.KeyTip = "1";
            this.button1_1.Label = "ユーザ設定のビューを削除";
            this.button1_1.Name = "button1_1";
            this.button1_1.ShowImage = true;
            this.button1_1.SuperTip = "ユーザ設定のビューを削除";
            this.button1_1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_1_Click);
            // 
            // button1_2
            // 
            this.button1_2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1_2.KeyTip = "2";
            this.button1_2.Label = "全シートのオートフィルターと非表示行列を解除する";
            this.button1_2.Name = "button1_2";
            this.button1_2.ShowImage = true;
            this.button1_2.SuperTip = "全シートのオートフィルターと非表示行列を解除する";
            this.button1_2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_2_Click);
            // 
            // button1_3
            // 
            this.button1_3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1_3.KeyTip = "3";
            this.button1_3.Label = "全シートのウィンドウ枠の固定と分割を解除する";
            this.button1_3.Name = "button1_3";
            this.button1_3.ShowImage = true;
            this.button1_3.SuperTip = "全シートのウィンドウ枠の固定と分割を解除する";
            this.button1_3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_3_Click);
            // 
            // button1_4
            // 
            this.button1_4.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1_4.KeyTip = "4";
            this.button1_4.Label = "全シートの枠線を非表示にする";
            this.button1_4.Name = "button1_4";
            this.button1_4.ShowImage = true;
            this.button1_4.SuperTip = "全シートの枠線を非表示にする";
            this.button1_4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_4_Click);
            // 
            // button1_5
            // 
            this.button1_5.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1_5.Label = "全シートのグループを解除する";
            this.button1_5.Name = "button1_5";
            this.button1_5.ShowImage = true;
            this.button1_5.SuperTip = "全シートのグループを解除する";
            this.button1_5.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_5_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.button2_1);
            this.group2.Items.Add(this.zoom1_5);
            this.group2.Items.Add(this.cursor1_5);
            this.group2.Label = "グループ2";
            this.group2.Name = "group2";
            // 
            // button2_1
            // 
            this.button2_1.KeyTip = "5";
            this.button2_1.Label = "全シートのズームとカーソル位置を揃える";
            this.button2_1.Name = "button2_1";
            this.button2_1.SuperTip = "※集計と更新履歴シートはズーム率100%固定です";
            this.button2_1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_1_Click);
            // 
            // zoom1_5
            // 
            this.zoom1_5.Label = "ズーム率";
            this.zoom1_5.MaxLength = 3;
            this.zoom1_5.Name = "zoom1_5";
            this.zoom1_5.SuperTip = "ズーム率";
            this.zoom1_5.Text = "70";
            this.zoom1_5.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.zoom1_5_TextChanged);
            // 
            // cursor1_5
            // 
            this.cursor1_5.Label = "カーソル位置";
            this.cursor1_5.MaxLength = 4;
            this.cursor1_5.Name = "cursor1_5";
            this.cursor1_5.SuperTip = "カーソル位置";
            this.cursor1_5.Text = "A1";
            this.cursor1_5.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cursor1_5_TextChanged);
            // 
            // group3
            // 
            this.group3.Items.Add(this.button1_6);
            this.group3.Items.Add(this.button1_7);
            this.group3.Items.Add(this.editBox1_7);
            this.group3.Label = "グループ3";
            this.group3.Name = "group3";
            // 
            // button1_6
            // 
            this.button1_6.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1_6.Image = ((System.Drawing.Image)(resources.GetObject("button1_6.Image")));
            this.button1_6.KeyTip = "N";
            this.button1_6.Label = "改行を追加";
            this.button1_6.Name = "button1_6";
            this.button1_6.ShowImage = true;
            this.button1_6.SuperTip = "改行を追加";
            this.button1_6.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_6_Click);
            // 
            // button1_7
            // 
            this.button1_7.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1_7.Image = ((System.Drawing.Image)(resources.GetObject("button1_7.Image")));
            this.button1_7.KeyTip = "K";
            this.button1_7.Label = "手順番号振りなおし";
            this.button1_7.Name = "button1_7";
            this.button1_7.ShowImage = true;
            this.button1_7.SuperTip = "手順番号振りなおし";
            this.button1_7.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_7_Click);
            // 
            // editBox1_7
            // 
            this.editBox1_7.Label = "「上記に続けて」文言";
            this.editBox1_7.Name = "editBox1_7";
            this.editBox1_7.Text = null;
            this.editBox1_7.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.editBox1_7_TextChanged);
            // 
            // group4
            // 
            this.group4.Items.Add(this.button1_8);
            this.group4.Items.Add(this.editBox1_8);
            this.group4.Items.Add(this.button1_9);
            this.group4.Label = "グループ4";
            this.group4.Name = "group4";
            // 
            // button1_8
            // 
            this.button1_8.Label = "列を記憶";
            this.button1_8.Name = "button1_8";
            this.button1_8.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_8_Click);
            // 
            // editBox1_8
            // 
            this.editBox1_8.Enabled = false;
            this.editBox1_8.Label = " ";
            this.editBox1_8.MaxLength = 2;
            this.editBox1_8.Name = "editBox1_8";
            this.editBox1_8.Text = null;
            // 
            // button1_9
            // 
            this.button1_9.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1_9.Image = ((System.Drawing.Image)(resources.GetObject("button1_9.Image")));
            this.button1_9.KeyTip = "Y";
            this.button1_9.Label = "列へジャンプ";
            this.button1_9.Name = "button1_9";
            this.button1_9.ShowImage = true;
            this.button1_9.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_9_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1_1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1_2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1_3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1_4;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2_1;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox zoom1_5;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox cursor1_5;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1_6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1_7;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox1_7;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1_8;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1_9;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox1_8;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1_5;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
