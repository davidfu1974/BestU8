using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Silver.UI;
using System.Runtime.InteropServices;

namespace BestU8
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();

            //填写状态栏信息
            toolStripStatusLabel.Text = "已登陆";
            toolStripStatususeridtext.Text = Pubvar.gu8userdata.UserId;
            toolStripStatuscompanytext.Text = "[" + Pubvar.gu8userdata.AccID + "]" + Pubvar.gu8userdata.AccName;
            toolStripStatusoperationdatetext.Text = Pubvar.gu8userdata.operDate;

            //构建左侧菜单栏并设置基本属性
            u8toolBox.BackColor = System.Drawing.SystemColors.Control;
            u8toolBox.Dock = System.Windows.Forms.DockStyle.Fill;
            u8toolBox.TabHeight = 18;
            u8toolBox.ItemHeight = 20;
            u8toolBox.ItemSpacing = 1;
            u8toolBox.ItemHoverColor = System.Drawing.Color.OldLace;  
            u8toolBox.ItemNormalColor = System.Drawing.SystemColors.Control;
            u8toolBox.ItemSelectedColor = System.Drawing.Color.BurlyWood; 

            //创建tab菜单项 -- U8数据接口
            ToolBoxTab _tab_u8datainterface = new ToolBoxTab("U8数据接口", 1);
            _tab_u8datainterface.SmallImageIndex = 2;
            u8toolBox.AddTab(_tab_u8datainterface);
            //创建tab菜单项下子菜单项 -- U8数据接口 -- 数据导入 
            ToolBoxItem _itemdataimport = new ToolBoxItem();
            _itemdataimport.Caption = "数据导入";
            _itemdataimport.Enabled = true;
            _itemdataimport.AllowDrag = false;
            _itemdataimport.SmallImageIndex = 1;
            _itemdataimport.Object = new Rectangle(10, 10, 100, 100);
            u8toolBox[0].AddItem(_itemdataimport);
            _itemdataimport.Selected = false;
            //隐藏tab control 及标签页
            U8tabCtl.Visible = false;
            U8dataimporttabPage.Parent = null;

        }

        private void u8toolBox_ItemSelectionChanged(ToolBoxItem sender, EventArgs e)
        {
            if (sender.Caption == "数据导入")
            {
                //显示tab control 及对应标签页
                U8tabCtl.Visible = true;
                U8dataimporttabPage.Parent = U8tabCtl;
            }

        }


        private void u8toolBox_ItemMouseDown(ToolBoxItem sender, MouseEventArgs e)
        {
            if (sender.Caption == "数据导入")
            {
                //显示tab control 及对应标签页
                U8tabCtl.Visible = true;
                U8dataimporttabPage.Parent = U8tabCtl;
            }
        }

        private void GLdataimportBT_Click(object sender, EventArgs e)
        {
            Pubvar.gdataimporttype = GLdataimportBT.Text;     //总账凭证模板数据
            DataImport dataimportform = new DataImport();
            dataimportform.StartPosition = FormStartPosition.CenterParent;
            dataimportform.ShowDialog();
        }

        private void receiptnoteBT_Click(object sender, EventArgs e)
        {
            Pubvar.gdataimporttype = receiptnoteBT.Text;     //采购入库单模板数据
            DataImport dataimportform = new DataImport();
            dataimportform.StartPosition = FormStartPosition.CenterParent;
            dataimportform.ShowDialog();
        }

        private void reloginMenuItem_Click(object sender, EventArgs e)
        {
            //显示U8门户登陆界面,处理用户登陆信息
            UFSoft.U8.Framework.Login.UI.clsLogin u8LoginUI = new UFSoft.U8.Framework.Login.UI.clsLogin();
            if (!u8LoginUI.login("DP"))
            {
                MessageBox.Show("登陆失败，原因：" + u8LoginUI.ErrDescript);
                u8LoginUI.ShutDown();
                return;
            }

            //从这个类里可以获取登陆信息、数据库连接信息等等
            UFSoft.U8.Framework.LoginContext.UserData u8userdata = new UFSoft.U8.Framework.LoginContext.UserData();
            u8userdata = u8LoginUI.GetLoginInfo();

            Pubvar.gu8LoginUI = u8LoginUI;
            Pubvar.gu8userdata = u8userdata;
            Pubvar.gdataimporttype = "";

            //填写状态栏信息
            toolStripStatusLabel.Text = "已登陆";
            toolStripStatususeridtext.Text = Pubvar.gu8userdata.UserId;
            toolStripStatuscompanytext.Text = "[" + Pubvar.gu8userdata.AccID + "]" + Pubvar.gu8userdata.AccName;
            toolStripStatusoperationdatetext.Text = Pubvar.gu8userdata.operDate;

            //隐藏tab control 及标签页
            U8tabCtl.Visible = false;
            U8dataimporttabPage.Parent = null;

        }

        private void exitMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
