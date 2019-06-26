namespace BestU8
{
    partial class MainForm
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.MainmenuStrip = new System.Windows.Forms.MenuStrip();
            this.系统ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.reloginMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exitMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.帮助ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.关于ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.formstatusStrip = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripStatusLabel7 = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripStatususerid = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripStatususeridtext = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripStatusLabelcompany = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripStatuscompanytext = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripStatusLabeloperationdate = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripStatusoperationdatetext = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.u8toolBox = new Silver.UI.ToolBox();
            this.U8tabCtl = new System.Windows.Forms.TabControl();
            this.U8dataimporttabPage = new System.Windows.Forms.TabPage();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.toolStripButton1 = new System.Windows.Forms.ToolStripButton();
            this.toolStripButton3 = new System.Windows.Forms.ToolStripButton();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.GLdataimportBT = new System.Windows.Forms.Button();
            this.receiptnoteBT = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.MainmenuStrip.SuspendLayout();
            this.formstatusStrip.SuspendLayout();
            this.toolStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.U8tabCtl.SuspendLayout();
            this.U8dataimporttabPage.SuspendLayout();
            this.SuspendLayout();
            // 
            // MainmenuStrip
            // 
            this.MainmenuStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.系统ToolStripMenuItem,
            this.帮助ToolStripMenuItem});
            this.MainmenuStrip.Location = new System.Drawing.Point(0, 0);
            this.MainmenuStrip.Name = "MainmenuStrip";
            this.MainmenuStrip.Size = new System.Drawing.Size(1152, 25);
            this.MainmenuStrip.TabIndex = 0;
            this.MainmenuStrip.Text = "MainmenuStrip";
            // 
            // 系统ToolStripMenuItem
            // 
            this.系统ToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.reloginMenuItem,
            this.exitMenuItem});
            this.系统ToolStripMenuItem.Name = "系统ToolStripMenuItem";
            this.系统ToolStripMenuItem.Size = new System.Drawing.Size(44, 21);
            this.系统ToolStripMenuItem.Text = "系统";
            // 
            // reloginMenuItem
            // 
            this.reloginMenuItem.Name = "reloginMenuItem";
            this.reloginMenuItem.Size = new System.Drawing.Size(152, 22);
            this.reloginMenuItem.Text = "重新登陆";
            this.reloginMenuItem.Click += new System.EventHandler(this.reloginMenuItem_Click);
            // 
            // exitMenuItem
            // 
            this.exitMenuItem.Name = "exitMenuItem";
            this.exitMenuItem.Size = new System.Drawing.Size(152, 22);
            this.exitMenuItem.Text = "退出系统";
            this.exitMenuItem.Click += new System.EventHandler(this.exitMenuItem_Click);
            // 
            // 帮助ToolStripMenuItem
            // 
            this.帮助ToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.关于ToolStripMenuItem});
            this.帮助ToolStripMenuItem.Name = "帮助ToolStripMenuItem";
            this.帮助ToolStripMenuItem.Size = new System.Drawing.Size(44, 21);
            this.帮助ToolStripMenuItem.Text = "帮助";
            // 
            // 关于ToolStripMenuItem
            // 
            this.关于ToolStripMenuItem.Name = "关于ToolStripMenuItem";
            this.关于ToolStripMenuItem.Size = new System.Drawing.Size(100, 22);
            this.关于ToolStripMenuItem.Text = "关于";
            // 
            // formstatusStrip
            // 
            this.formstatusStrip.AutoSize = false;
            this.formstatusStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel,
            this.toolStripStatusLabel7,
            this.toolStripStatususerid,
            this.toolStripStatususeridtext,
            this.toolStripStatusLabelcompany,
            this.toolStripStatuscompanytext,
            this.toolStripStatusLabeloperationdate,
            this.toolStripStatusoperationdatetext});
            this.formstatusStrip.Location = new System.Drawing.Point(0, 674);
            this.formstatusStrip.Name = "formstatusStrip";
            this.formstatusStrip.Size = new System.Drawing.Size(1152, 26);
            this.formstatusStrip.SizingGrip = false;
            this.formstatusStrip.TabIndex = 1;
            this.formstatusStrip.Text = "formstatusStrip";
            // 
            // toolStripStatusLabel
            // 
            this.toolStripStatusLabel.BorderSides = System.Windows.Forms.ToolStripStatusLabelBorderSides.Right;
            this.toolStripStatusLabel.Name = "toolStripStatusLabel";
            this.toolStripStatusLabel.Size = new System.Drawing.Size(204, 21);
            this.toolStripStatusLabel.Text = "状态。。。。。。。。。。。。。。";
            // 
            // toolStripStatusLabel7
            // 
            this.toolStripStatusLabel7.Name = "toolStripStatusLabel7";
            this.toolStripStatusLabel7.Size = new System.Drawing.Size(543, 21);
            this.toolStripStatusLabel7.Spring = true;
            // 
            // toolStripStatususerid
            // 
            this.toolStripStatususerid.BorderSides = System.Windows.Forms.ToolStripStatusLabelBorderSides.Right;
            this.toolStripStatususerid.Name = "toolStripStatususerid";
            this.toolStripStatususerid.Size = new System.Drawing.Size(36, 21);
            this.toolStripStatususerid.Text = "用户";
            // 
            // toolStripStatususeridtext
            // 
            this.toolStripStatususeridtext.BorderSides = System.Windows.Forms.ToolStripStatusLabelBorderSides.Right;
            this.toolStripStatususeridtext.Name = "toolStripStatususeridtext";
            this.toolStripStatususeridtext.Size = new System.Drawing.Size(92, 21);
            this.toolStripStatususeridtext.Text = "XXXXXXXXXX";
            // 
            // toolStripStatusLabelcompany
            // 
            this.toolStripStatusLabelcompany.BorderSides = System.Windows.Forms.ToolStripStatusLabelBorderSides.Right;
            this.toolStripStatusLabelcompany.Name = "toolStripStatusLabelcompany";
            this.toolStripStatusLabelcompany.Size = new System.Drawing.Size(36, 21);
            this.toolStripStatusLabelcompany.Text = "账套";
            // 
            // toolStripStatuscompanytext
            // 
            this.toolStripStatuscompanytext.BorderSides = System.Windows.Forms.ToolStripStatusLabelBorderSides.Right;
            this.toolStripStatuscompanytext.Name = "toolStripStatuscompanytext";
            this.toolStripStatuscompanytext.Size = new System.Drawing.Size(88, 21);
            this.toolStripStatuscompanytext.Text = "XXXXX某公司";
            // 
            // toolStripStatusLabeloperationdate
            // 
            this.toolStripStatusLabeloperationdate.BorderSides = System.Windows.Forms.ToolStripStatusLabelBorderSides.Right;
            this.toolStripStatusLabeloperationdate.Name = "toolStripStatusLabeloperationdate";
            this.toolStripStatusLabeloperationdate.Size = new System.Drawing.Size(60, 21);
            this.toolStripStatusLabeloperationdate.Text = "操作日期";
            // 
            // toolStripStatusoperationdatetext
            // 
            this.toolStripStatusoperationdatetext.BorderSides = System.Windows.Forms.ToolStripStatusLabelBorderSides.Right;
            this.toolStripStatusoperationdatetext.Name = "toolStripStatusoperationdatetext";
            this.toolStripStatusoperationdatetext.Size = new System.Drawing.Size(78, 21);
            this.toolStripStatusoperationdatetext.Text = "2019-06-28";
            // 
            // toolStrip1
            // 
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripButton1,
            this.toolStripButton3});
            this.toolStrip1.Location = new System.Drawing.Point(0, 25);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(1152, 25);
            this.toolStrip1.TabIndex = 2;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // splitContainer1
            // 
            this.splitContainer1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(0, 50);
            this.splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.u8toolBox);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.U8tabCtl);
            this.splitContainer1.Size = new System.Drawing.Size(1152, 624);
            this.splitContainer1.SplitterDistance = 203;
            this.splitContainer1.TabIndex = 3;
            // 
            // u8toolBox
            // 
            this.u8toolBox.AllowSwappingByDragDrop = false;
            this.u8toolBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.u8toolBox.InitialScrollDelay = 500;
            this.u8toolBox.ItemBackgroundColor = System.Drawing.Color.Empty;
            this.u8toolBox.ItemBorderColor = System.Drawing.Color.Empty;
            this.u8toolBox.ItemHeight = 20;
            this.u8toolBox.ItemHoverColor = System.Drawing.SystemColors.Control;
            this.u8toolBox.ItemHoverTextColor = System.Drawing.SystemColors.ControlText;
            this.u8toolBox.ItemNormalColor = System.Drawing.SystemColors.Control;
            this.u8toolBox.ItemNormalTextColor = System.Drawing.SystemColors.ControlText;
            this.u8toolBox.ItemSelectedColor = System.Drawing.Color.White;
            this.u8toolBox.ItemSelectedTextColor = System.Drawing.SystemColors.ControlText;
            this.u8toolBox.ItemSpacing = 2;
            this.u8toolBox.LargeItemSize = new System.Drawing.Size(64, 64);
            this.u8toolBox.LayoutDelay = 10;
            this.u8toolBox.Location = new System.Drawing.Point(0, 0);
            this.u8toolBox.Name = "u8toolBox";
            this.u8toolBox.ScrollDelay = 60;
            this.u8toolBox.SelectAllTextWhileRenaming = true;
            this.u8toolBox.SelectedTabIndex = -1;
            this.u8toolBox.ShowOnlyOneItemPerRow = false;
            this.u8toolBox.Size = new System.Drawing.Size(199, 620);
            this.u8toolBox.SmallImageList = this.imageList1;
            this.u8toolBox.SmallItemSize = new System.Drawing.Size(16, 32);
            this.u8toolBox.TabHeight = 18;
            this.u8toolBox.TabHoverTextColor = System.Drawing.SystemColors.ControlText;
            this.u8toolBox.TabIndex = 0;
            this.u8toolBox.TabNormalTextColor = System.Drawing.SystemColors.ControlText;
            this.u8toolBox.TabSelectedTextColor = System.Drawing.SystemColors.ControlText;
            this.u8toolBox.TabSpacing = 1;
            this.u8toolBox.UseItemColorInRename = false;
            this.u8toolBox.ItemSelectionChanged += new Silver.UI.ItemSelectionChangedHandler(this.u8toolBox_ItemSelectionChanged);
            this.u8toolBox.ItemMouseDown += new Silver.UI.ItemMouseEventHandler(this.u8toolBox_ItemMouseDown);
            // 
            // U8tabCtl
            // 
            this.U8tabCtl.Controls.Add(this.U8dataimporttabPage);
            this.U8tabCtl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.U8tabCtl.ImageList = this.imageList1;
            this.U8tabCtl.Location = new System.Drawing.Point(0, 0);
            this.U8tabCtl.Name = "U8tabCtl";
            this.U8tabCtl.SelectedIndex = 0;
            this.U8tabCtl.Size = new System.Drawing.Size(941, 620);
            this.U8tabCtl.TabIndex = 0;
            // 
            // U8dataimporttabPage
            // 
            this.U8dataimporttabPage.BackColor = System.Drawing.Color.Transparent;
            this.U8dataimporttabPage.Controls.Add(this.groupBox2);
            this.U8dataimporttabPage.Controls.Add(this.groupBox1);
            this.U8dataimporttabPage.Controls.Add(this.receiptnoteBT);
            this.U8dataimporttabPage.Controls.Add(this.GLdataimportBT);
            this.U8dataimporttabPage.Controls.Add(this.label2);
            this.U8dataimporttabPage.Controls.Add(this.label1);
            this.U8dataimporttabPage.ForeColor = System.Drawing.Color.Black;
            this.U8dataimporttabPage.ImageIndex = 0;
            this.U8dataimporttabPage.Location = new System.Drawing.Point(4, 23);
            this.U8dataimporttabPage.Name = "U8dataimporttabPage";
            this.U8dataimporttabPage.Padding = new System.Windows.Forms.Padding(3);
            this.U8dataimporttabPage.Size = new System.Drawing.Size(933, 593);
            this.U8dataimporttabPage.TabIndex = 0;
            this.U8dataimporttabPage.Text = "U8数据导入";
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "hideproduct_16x16.png");
            this.imageList1.Images.SetKeyName(1, "contentarrangeinrows_16x161.png");
            this.imageList1.Images.SetKeyName(2, "addnewdatasource_16x16.png");
            this.imageList1.Images.SetKeyName(3, "highlightactiveelements_16x16.png");
            // 
            // toolStripButton1
            // 
            this.toolStripButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripButton1.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton1.Image")));
            this.toolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton1.Name = "toolStripButton1";
            this.toolStripButton1.Size = new System.Drawing.Size(23, 22);
            this.toolStripButton1.Text = "toolStripButton1";
            // 
            // toolStripButton3
            // 
            this.toolStripButton3.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripButton3.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton3.Image")));
            this.toolStripButton3.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton3.Name = "toolStripButton3";
            this.toolStripButton3.Size = new System.Drawing.Size(23, 22);
            this.toolStripButton3.Text = "toolStripButton3";
            // 
            // label1
            // 
            this.label1.ForeColor = System.Drawing.SystemColors.Highlight;
            this.label1.Location = new System.Drawing.Point(68, 28);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(200, 23);
            this.label1.TabIndex = 0;
            this.label1.Text = "财务单据导入";
            this.label1.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // label2
            // 
            this.label2.ForeColor = System.Drawing.SystemColors.Highlight;
            this.label2.Location = new System.Drawing.Point(356, 28);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(153, 23);
            this.label2.TabIndex = 1;
            this.label2.Text = "库存单据导入";
            this.label2.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // GLdataimportBT
            // 
            this.GLdataimportBT.Cursor = System.Windows.Forms.Cursors.Hand;
            this.GLdataimportBT.FlatAppearance.BorderColor = System.Drawing.SystemColors.Control;
            this.GLdataimportBT.FlatAppearance.MouseDownBackColor = System.Drawing.Color.OldLace;
            this.GLdataimportBT.FlatAppearance.MouseOverBackColor = System.Drawing.Color.OldLace;
            this.GLdataimportBT.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.GLdataimportBT.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.GLdataimportBT.ImageIndex = 3;
            this.GLdataimportBT.ImageList = this.imageList1;
            this.GLdataimportBT.Location = new System.Drawing.Point(102, 68);
            this.GLdataimportBT.Name = "GLdataimportBT";
            this.GLdataimportBT.Size = new System.Drawing.Size(131, 23);
            this.GLdataimportBT.TabIndex = 2;
            this.GLdataimportBT.Text = "总账凭证导入";
            this.GLdataimportBT.UseVisualStyleBackColor = true;
            this.GLdataimportBT.Click += new System.EventHandler(this.GLdataimportBT_Click);
            // 
            // receiptnoteBT
            // 
            this.receiptnoteBT.Cursor = System.Windows.Forms.Cursors.Hand;
            this.receiptnoteBT.FlatAppearance.BorderColor = System.Drawing.SystemColors.Control;
            this.receiptnoteBT.FlatAppearance.MouseDownBackColor = System.Drawing.Color.OldLace;
            this.receiptnoteBT.FlatAppearance.MouseOverBackColor = System.Drawing.Color.OldLace;
            this.receiptnoteBT.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.receiptnoteBT.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.receiptnoteBT.ImageIndex = 3;
            this.receiptnoteBT.ImageList = this.imageList1;
            this.receiptnoteBT.Location = new System.Drawing.Point(369, 68);
            this.receiptnoteBT.Name = "receiptnoteBT";
            this.receiptnoteBT.Size = new System.Drawing.Size(138, 23);
            this.receiptnoteBT.TabIndex = 3;
            this.receiptnoteBT.Text = "采购入库单导入";
            this.receiptnoteBT.UseVisualStyleBackColor = true;
            this.receiptnoteBT.Click += new System.EventHandler(this.receiptnoteBT_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Location = new System.Drawing.Point(68, 54);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(200, 2);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "groupBox1";
            // 
            // groupBox2
            // 
            this.groupBox2.Location = new System.Drawing.Point(338, 54);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(200, 2);
            this.groupBox2.TabIndex = 5;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "groupBox2";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1152, 700);
            this.Controls.Add(this.splitContainer1);
            this.Controls.Add(this.toolStrip1);
            this.Controls.Add(this.formstatusStrip);
            this.Controls.Add(this.MainmenuStrip);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MainMenuStrip = this.MainmenuStrip;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "MainForm";
            this.Text = "BESTU8";
            this.MainmenuStrip.ResumeLayout(false);
            this.MainmenuStrip.PerformLayout();
            this.formstatusStrip.ResumeLayout(false);
            this.formstatusStrip.PerformLayout();
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.U8tabCtl.ResumeLayout(false);
            this.U8dataimporttabPage.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip MainmenuStrip;
        private System.Windows.Forms.ToolStripMenuItem 系统ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem reloginMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exitMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 帮助ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 关于ToolStripMenuItem;
        private System.Windows.Forms.StatusStrip formstatusStrip;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatususerid;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatususeridtext;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabelcompany;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatuscompanytext;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabeloperationdate;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusoperationdatetext;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel7;
        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripButton toolStripButton1;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private Silver.UI.ToolBox u8toolBox;
        private System.Windows.Forms.TabControl U8tabCtl;
        private System.Windows.Forms.TabPage U8dataimporttabPage;
        private System.Windows.Forms.ToolStripButton toolStripButton3;
        private System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button receiptnoteBT;
        private System.Windows.Forms.Button GLdataimportBT;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
    }
}

