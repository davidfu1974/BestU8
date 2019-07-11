﻿extern alias interU8lg;
extern alias interadodb;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using UFIDA.U8.MomServiceCommon;
using UFIDA.U8.U8MOMAPIFramework;
using UFIDA.U8.U8APIFramework;
using UFIDA.U8.U8APIFramework.Meta;
using UFIDA.U8.U8APIFramework.Parameter;
using MSXML2;



namespace BestU8
{

    public partial class DataImport : Form
    {

        [DllImport("kernel32.dll")]
        public static extern IntPtr _lopen(string lpPathName, int iReadWrite);

        [DllImport("kernel32.dll")]
        public static extern bool CloseHandle(IntPtr hObject);

        public const int OF_READWRITE = 2;
        public const int OF_SHARE_DENY_NONE = 0x40;
        public readonly IntPtr HFILE_ERROR = new IntPtr(-1);



        public DataImport()
        {
            InitializeComponent();
            importtypeGB.Text = Pubvar.gdataimporttype;
            this.Text = "DataImport - " + Pubvar.gdataimporttype;
        }

        private void importdataopenfilebutton_Click(object sender, EventArgs e)
        {
            importdataopenFileDialog.InitialDirectory = "c:\\";
            importdataopenFileDialog.Filter = "Excel文件(*.xlsx)|*.xlsx|Excel文件(*.xls)|*.xls|所有文件(*.*)|*.*";
            importdataopenFileDialog.RestoreDirectory = true;
            importdataopenFileDialog.FilterIndex = 1;
            if (importdataopenFileDialog.ShowDialog() == DialogResult.OK)
            {
                importdatafiletextBox.Text = importdataopenFileDialog.FileName;
            }
        }

        private void closebutton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void importdatabutton_Click(object sender, EventArgs e)
        {
            DataSet dsexcel = new DataSet();
            DataSet dstoexcel = new DataSet();
            int importsuccessrows = 0, importfailurerows = 0;
            ExcelHelper npoidata = new ExcelHelper(importdatafiletextBox.Text);
            DataTable dtnpoidata = new DataTable();

            //为防止用户多次点击导入按钮，将按钮禁用
            importdatabutton.Enabled = false;
            closebutton.Enabled = false;
            this.ControlBox = false;

            //判断数据模板EXCEL是否被用户打开
            IntPtr vHandle = _lopen(importdatafiletextBox.Text, OF_READWRITE | OF_SHARE_DENY_NONE);
            if (vHandle == HFILE_ERROR)
            {
                MessageBox.Show("请先关闭数据模板导入EXCEL文件！");
                //启用导入按钮
                importdatabutton.Enabled = true;
                closebutton.Enabled = true;
                this.ControlBox = true;
                return;
            }
            CloseHandle(vHandle);

            //总账凭证导入
            #region
            if (Pubvar.gdataimporttype == "总账凭证导入")
            {
                string impstart, impend;
                dtnpoidata = npoidata.ExcelToDataTable("GLVouchers", true, importdatafiletextBox.Text);
                dtnpoidata.TableName = "GLVouchers";
                dsexcel.Tables.Add(dtnpoidata);
                impstart = DateTime.Now.ToLocalTime().ToString();
                importdataresulttextBox.AppendText("数据导入执行开始:" + impstart + "\n");
                importdataresulttextBox.Refresh();
                //调用总账导入功能
                bool v_importglvouchersflag = GLvouchersimport(Pubvar.gu8LoginUI.userToken, Pubvar.gu8userdata.ConnString, dsexcel, Pubvar.gu8userdata.UserId, out importsuccessrows, out importfailurerows, out dstoexcel);
                //导入结果回写EXCEL 
                string f1 = System.IO.Path.GetFileNameWithoutExtension(importdatafiletextBox.Text);//文件名没有扩展名
                string f2 = System.IO.Path.GetDirectoryName(importdatafiletextBox.Text);           //获取路径
                string fe = System.IO.Path.GetExtension(importdatafiletextBox.Text);               //文件扩展名
                string tempfile = f2 + "\\" + f1 + "_tmp" + fe;
                npoidata.DataTableToExcel(dstoexcel.Tables["GLVouchers"], "GLVouchers", true, tempfile);
                //如果存在则删除并将临时文件重命名
                if (System.IO.File.Exists(importdatafiletextBox.Text))
                {
                    System.IO.File.Delete(importdatafiletextBox.Text);
                }
                System.IO.File.Move(tempfile, importdatafiletextBox.Text);

                //执行结果回写memo text
                impend = DateTime.Now.ToLocalTime().ToString();
                importdataresulttextBox.AppendText("此次数据导入共计执行：" + (importsuccessrows + importfailurerows) + " 条 \n");
                importdataresulttextBox.AppendText("其中导入成功：" + importsuccessrows + " 条 \n");
                importdataresulttextBox.AppendText("其中导入失败：" + importfailurerows + " 条 \n");
                importdataresulttextBox.AppendText("如果导入有出错，具体原因请看导入数据模板中错误信息列，请纠正后再次执行导入！\n");
                importdataresulttextBox.AppendText("数据导入执行结束:" + impend + "  \n");
                importdataresulttextBox.Refresh();
            }
            #endregion

            //采购入库单导入
            #region
            if (Pubvar.gdataimporttype == "采购入库单导入")
            {
                string impstart, impend, v_errmsg;
                //采购入库单导入EXCEL中
                dtnpoidata = npoidata.ExcelToDataTable("ReceiptNotes", true, importdatafiletextBox.Text);
                dtnpoidata.TableName = "ReceiptNotes";
                dsexcel.Tables.Add(dtnpoidata);
                impstart = DateTime.Now.ToLocalTime().ToString();
                importdataresulttextBox.AppendText("数据导入执行开始:" + impstart + "\n");
                importdataresulttextBox.Refresh();
                //调用采购入库单导入功能
                bool v_importreceiptnoteflag = ReceiptNoteimport(Pubvar.gu8userdata, dsexcel, out importsuccessrows, out importfailurerows, out dstoexcel, out v_errmsg);
                //导入结果回写EXCEL 
                string f1 = System.IO.Path.GetFileNameWithoutExtension(importdatafiletextBox.Text);//文件名没有扩展名
                string f2 = System.IO.Path.GetDirectoryName(importdatafiletextBox.Text);           //获取路径
                string fe = System.IO.Path.GetExtension(importdatafiletextBox.Text);               //文件扩展名
                string tempfile = f2 + "\\" + f1 + "_tmp" + fe;
                npoidata.DataTableToExcel(dstoexcel.Tables["ReceiptNotes"], "ReceiptNotes", true, tempfile);
                //如果存在则删除并将临时文件重命名
                if (System.IO.File.Exists(importdatafiletextBox.Text))
                {
                    System.IO.File.Delete(importdatafiletextBox.Text);
                }
                System.IO.File.Move(tempfile, importdatafiletextBox.Text);

                //执行结果回写memo text
                impend = DateTime.Now.ToLocalTime().ToString();
                importdataresulttextBox.AppendText("此次数据导入共计执行：" + (importsuccessrows + importfailurerows) + " 条 \n");
                importdataresulttextBox.AppendText("其中导入成功：" + importsuccessrows + " 条 \n");
                importdataresulttextBox.AppendText("其中导入失败：" + importfailurerows + " 条 \n");
                importdataresulttextBox.AppendText("如果导入有出错，具体原因请看导入数据模板中错误信息列，请纠正后再次执行导入！\n");
                importdataresulttextBox.AppendText("数据导入执行结束:" + impend + "  \n");
                importdataresulttextBox.Refresh();
            }
            #endregion


            //为防止用户多次点击导入按钮，将按钮禁用
            importdatabutton.Enabled = true;
            closebutton.Enabled = true;
            this.ControlBox = true;
        }


        public bool GLvouchersimport(string usertoken, string dbconn, DataSet dsimportedvouchers, string userid, out int importsuccessrows, out int importfailurerows, out DataSet dsreturnvouchers)
        {
            string strSql = "", strTempTable = "tempdb.dbo.cus_gl_accvouchers";
            int v_importsuccessrows = 0, v_importfailurerows = 0;
            System.Object rsaffected = new System.Object();
            //创建或清除凭证导入临时表数据
            #region
            interadodb::ADODB.Recordset rs = new interadodb::ADODB.Recordset();
            interadodb::ADODB.Connection conn = new interadodb::ADODB.Connection();
            conn.Open(dbconn);
            strSql = "SELECT count(*) FROM tempdb.dbo.sysobjects WHERE name = 'cus_gl_accvouchers'";
            rs = conn.Execute(strSql, out rsaffected, -1);
            if (Convert.ToInt16(rs.Fields[0].Value) > 0)
            {
                strSql = "DELETE FROM tempdb.dbo.cus_gl_accvouchers ";
                rs = conn.Execute(strSql, out rsaffected, -1);
            }
            else
            {
                strSql = "CREATE TABLE " + strTempTable;
                strSql = strSql + "( csign             NVARCHAR (28),";         //凭证类别字
                strSql = strSql + "ino_id              SMALLINT,";              //凭证编号
                strSql = strSql + "inid                SMALLINT,";              //行号
                strSql = strSql + "cbill               NVARCHAR (80),";         //制单人
                strSql = strSql + "doutbilldate        DATETIME,";              //外部凭证制单日期
                strSql = strSql + "ccashier            NVARCHAR (80),";         //出纳签字人
                strSql = strSql + "idoc                SMALLINT DEFAULT 0,";    //附单据数
                strSql = strSql + "ctext1              NVARCHAR (50),";         //凭证头自定义项1
                strSql = strSql + "ctext2              NVARCHAR (50),";         //凭证头自定义项2
                strSql = strSql + "cexch_name          NVARCHAR (28),";         //币种名称
                strSql = strSql + "cdigest             NVARCHAR (120),";        //凭证摘要
                strSql = strSql + "ccode               NVARCHAR (40),";         //科目编码
                strSql = strSql + "md                  MONEY DEFAULT 0,";       //借方金额
                strSql = strSql + "mc                  MONEY DEFAULT 0,";       //贷方金额
                strSql = strSql + "md_f                MONEY DEFAULT 0,";       //外币借方金额
                strSql = strSql + "mc_f                MONEY DEFAULT 0,";       //外币贷方金额
                strSql = strSql + "nfrat               FLOAT DEFAULT 0,";       //汇率
                strSql = strSql + "nd_s                FLOAT DEFAULT 0,";       //数量借方
                strSql = strSql + "nc_s                FLOAT DEFAULT 0,";       //数量贷方
                strSql = strSql + "csettle             NVARCHAR (23),";         //结算方式编码
                strSql = strSql + "cn_id               NVARCHAR (30),";         //票据号
                strSql = strSql + "dt_date             DATETIME,";              //票号发生日期
                strSql = strSql + "cdept_id            NVARCHAR (12),";         //部门编码
                strSql = strSql + "cperson_id          NVARCHAR (80),";         //职员编码
                strSql = strSql + "ccus_id             NVARCHAR (80),";         //客户编码
                strSql = strSql + "csup_id             NVARCHAR (20),";         //供应商编码
                strSql = strSql + "citem_id            NVARCHAR (80),";         //物料编码
                strSql = strSql + "citem_class         NVARCHAR (22),";         //物料大类编码
                strSql = strSql + "cname               NVARCHAR (40),";         //业务员
                strSql = strSql + "ccode_equal         NVARCHAR (50),";         //对方科目编码
                strSql = strSql + "bvouchedit          BIT DEFAULT 0,";         //凭证是否可修改
                strSql = strSql + "bvouchaddordele     BIT DEFAULT 0,";         //凭证分录是否可增删
                strSql = strSql + "bvouchmoneyhold     BIT DEFAULT 0,";         //凭证合计金额是否保值
                strSql = strSql + "bvalueedit          BIT DEFAULT 0,";         //分录数值是否可修改
                strSql = strSql + "bcodeedit           BIT DEFAULT 0,";         //分录科目是否可修改
                strSql = strSql + "ccodecontrol        NVARCHAR (50),";         //分录受控科目可用状态
                strSql = strSql + "bPCSedit            BIT DEFAULT 0,";         //分录来往项是否可修改
                strSql = strSql + "bDeptedit           BIT DEFAULT 0,";         //分录部门是否可修改
                strSql = strSql + "bItemedit           BIT DEFAULT 0,";         //分录物料是否可修改
                strSql = strSql + "bCusSupInput        BIT DEFAULT 0,";         //分录往来项是否必须输入
                strSql = strSql + "coutaccset          NVARCHAR (23),";         //外部凭证账套号
                strSql = strSql + "ioutyear            SMALLINT,";              //外部凭证会计年度
                strSql = strSql + "coutsysname         NVARCHAR (50) NOT NULL,";//外部凭证系统名称 这里如果不放GL 则外部导入的凭证无法修改。
                strSql = strSql + "coutsysver          NVARCHAR (50),";         //外部凭证系统版本号
                strSql = strSql + "ioutperiod          TINYINT NOT NULL,";      //外部凭证会计期间
                strSql = strSql + "coutsign            NVARCHAR (80) NOT NULL,";//外部凭证业务类型
                strSql = strSql + "coutno_id           NVARCHAR (100) NOT NULL,";//外部凭证业务号 （相同的话表示为一张凭证）
                strSql = strSql + "doutdate            DATETIME,";              //外部凭证单据日期
                strSql = strSql + "coutbillsign        NVARCHAR (80),";         //外部凭证单据类型
                strSql = strSql + "coutid              NVARCHAR (50),";         //外部凭证单据号
                strSql = strSql + "iflag               TINYINT,";               //凭证标志
                strSql = strSql + "iBG_ControlResult   SMALLINT NULL,";         //
                strSql = strSql + "daudit_date         DATETIME NULL,";         //
                strSql = strSql + "cblueoutno_id       NVARCHAR (50) NULL,";    //
                strSql = strSql + "bWH_BgFlag          BIT,";                   //
                strSql = strSql + "cDefine1            NVARCHAR (40),";         //自定义项1
                strSql = strSql + "cDefine2            NVARCHAR (40),";         //自定义项2
                strSql = strSql + "cDefine3            NVARCHAR (40),";
                strSql = strSql + "cDefine4            DATETIME,";
                strSql = strSql + "cDefine5            INT,";
                strSql = strSql + "cDefine6            DATETIME,";
                strSql = strSql + "cDefine7            FLOAT,";
                strSql = strSql + "cDefine8            NVARCHAR (4),";
                strSql = strSql + "cDefine9            NVARCHAR (8),";
                strSql = strSql + "cDefine10           NVARCHAR (60),";
                strSql = strSql + "cDefine11           NVARCHAR (120),";
                strSql = strSql + "cDefine12           NVARCHAR (120),";
                strSql = strSql + "cDefine13           NVARCHAR (120),";
                strSql = strSql + "cDefine14           NVARCHAR (120),";
                strSql = strSql + "cDefine15           INT,";
                strSql = strSql + "cDefine16           FLOAT )";
                rs = conn.Execute(strSql, out rsaffected, -1);

            }
            #endregion
            //临时表中插入总账凭证数据
            /*
            //测试数据
            //借方
            strSql = "INSERT INTO tempdb.dbo.cus_gl_accvouchers(ioutperiod,coutsign ,cSign,coutno_id,cdigest,coutsysname,cbill,inid,ccode,cexch_name ,doutbilldate,bvouchedit,bvouchaddordele,bvouchmoneyhold,bvalueedit,bcodeedit,md) ";
            strSql = strSql + "VALUES(1, N'记', N'记', N'IMP0000001', N'测试后台导入总账凭证', N'GL', N'" + userid + "', 1, N'6402', N'人民币',  '2015-1-31', 1, 1, 1,1,1, 777)";
            rs = conn.Execute(strSql, out rsaffected, -1);
            //贷方
            strSql = "INSERT INTO tempdb.dbo.cus_gl_accvouchers(ioutperiod,coutsign ,cSign,coutno_id,cdigest,coutsysname,cbill,inid,ccode,cexch_name ,doutbilldate,bvouchedit,bvouchaddordele,bvouchmoneyhold,bvalueedit,bcodeedit,mc) ";
            strSql = strSql + "VALUES(1, N'记', N'记', N'IMP0000001', N'测试后台导入总账凭证', N'GL', N'" + userid + "', 1, N'6711', N'人民币',  '2015-1-31', 1, 1, 1,1,1, 777)";
            rs = conn.Execute(strSql, out rsaffected, -1);
            */


            //调用API保存总账凭证
            CVoucher.CVInterface glcvoucher = new CVoucher.CVInterface();
            glcvoucher.set_Connection(conn);
            glcvoucher.strTempTable = strTempTable;
            glcvoucher.LoginByUserToken(usertoken);
            //根据dataset中导入数据分组循环导入U8系统
            dsreturnvouchers = dsimportedvouchers.Clone();
            DataTable dtdistinct = dsimportedvouchers.Tables["GLVouchers"].DefaultView.ToTable(true, new string[] { "凭证ID" });
            string vougroupby = "";

            //设置progressbar步长并显示百分比
            importdataprogressBar.Minimum = 0;   // 设置进度条最小值.
            importdataprogressBar.Value = 1;    // 设置进度条初始值
            importdataprogressBar.Step = 1;     // 设置每次增加的步长
            importdataprogressBar.Maximum = dtdistinct.Rows.Count;// 设置进度条最大值.
            Graphics g = this.importdataprogressBar.CreateGraphics();

            for (int i = 0; i < dtdistinct.Rows.Count; i++)
            {

                vougroupby = dtdistinct.Rows[i]["凭证ID"].ToString();
                string filterestr = "";
                //当凭证ID为空或""时的特殊处理
                if (!string.IsNullOrEmpty(vougroupby))
                {
                    filterestr = "凭证ID = " + "'" + vougroupby + "'";
                }
                else
                {
                    filterestr = "凭证ID IS NULL  OR 凭证ID ='" + "'";
                }
                DataRow[] drgroupby = dsimportedvouchers.Tables["GLVouchers"].Select(filterestr);
                if ((!string.IsNullOrEmpty(drgroupby[0]["凭证号"].ToString())) || ((!string.IsNullOrEmpty(drgroupby[0]["是否导入"].ToString())) && (drgroupby[0]["是否导入"].ToString() == "N")) || (string.IsNullOrEmpty(drgroupby[0]["凭证ID"].ToString())))
                {
                    //复制已成功导入的数据到返回数据表中
                    for (int k = 0; k < drgroupby.Count(); k++)
                    {
                        dsreturnvouchers.Tables["GLVouchers"].ImportRow(drgroupby[k]);
                    }
                }
                else
                {
                    for (int j = 0; j < drgroupby.Count(); j++)
                    {
                        strSql = "INSERT INTO tempdb.dbo.cus_gl_accvouchers(ioutperiod,coutsign ,cSign,coutno_id,cdigest,coutsysname,cbill,inid,ccode,cexch_name ,doutbilldate,bvouchedit,bvouchaddordele,bvouchmoneyhold,bvalueedit,bcodeedit,md,mc,cdept_id,cperson_id,ccus_id,csup_id,citem_class,citem_id,cname) ";
                        strSql = strSql + "VALUES(" + drgroupby[j]["会计期间"].ToString();
                        strSql = strSql + ",'" + drgroupby[j]["凭证类别"].ToString();
                        strSql = strSql + "','" + drgroupby[j]["凭证类别"].ToString();
                        strSql = strSql + "','" + drgroupby[j]["凭证ID"].ToString();
                        strSql = strSql + "','" + drgroupby[j]["摘要"].ToString();
                        strSql = strSql + "','" + "GL";   //这里外部系统设置为总账，否则导入的凭证默认无法修改。
                        strSql = strSql + "','" + userid;
                        strSql = strSql + "'," + (j + 1).ToString();   //行号
                        strSql = strSql + ",'" + drgroupby[j]["科目编码"].ToString();
                        strSql = strSql + "','" + drgroupby[j]["币种名称"].ToString();
                        strSql = strSql + "','" + drgroupby[j]["制单日期"].ToString();
                        strSql = strSql + "'," + 1;  //bvouchedit
                        strSql = strSql + "," + 1;   //bvouchaddordele
                        strSql = strSql + "," + 1;   //bvouchmoneyhold
                        strSql = strSql + "," + 1;   //bvalueedit,bcodeedit
                        strSql = strSql + "," + 1;   //bcodeedit
                        if (string.IsNullOrEmpty(drgroupby[j]["借方金额"].ToString()))
                        {
                            strSql = strSql + "," + 0;   //md
                        }
                        else
                        {
                            strSql = strSql + "," + drgroupby[j]["借方金额"].ToString();   //md
                        }

                        if (string.IsNullOrEmpty(drgroupby[j]["贷方金额"].ToString()))
                        {
                            strSql = strSql + "," + 0;   //mc
                        }
                        else
                        {
                            strSql = strSql + "," + drgroupby[j]["贷方金额"].ToString();   //mc
                        }

                        strSql = strSql + ",'" + drgroupby[j]["部门编码"].ToString();          //部门编码
                        strSql = strSql + "','" + drgroupby[j]["职员编码"].ToString();          //职员编码
                        strSql = strSql + "','" + drgroupby[j]["客户编码"].ToString();          //客户编码
                        strSql = strSql + "','" + drgroupby[j]["供应商编码"].ToString();          //供应商编码
                        strSql = strSql + "','" + drgroupby[j]["项目大类编码"].ToString();        //物料大类编码
                        strSql = strSql + "','" + drgroupby[j]["项目编码"].ToString();            //物料编码
                        strSql = strSql + "','" + drgroupby[j]["业务员"].ToString() + "')";         //业务员
                        rs = conn.Execute(strSql, out rsaffected, -1);
                    }
                    //凭证导入U8中制单
                    bool glsaveflag = glcvoucher.SaveVoucher();
                    //回写凭证号及错误信息,一旦SaveVoucher成功执行完毕，数据库连接系统API自动关闭，需要再次打开

                    if (glsaveflag)
                    {
                        v_importsuccessrows = v_importsuccessrows + 1;
                        int importedvoucherid;
                        strSql = "SELECT distinct ino_id  FROM tempdb.dbo.cus_gl_accvouchers WHERE coutno_id ='" + vougroupby + "'";
                        conn.Open(dbconn);
                        rs = conn.Execute(strSql, out rsaffected, -1);
                        if (Convert.ToInt16(rs.Fields[0].Value) > 0)
                        {
                            importedvoucherid = Convert.ToInt16(rs.Fields[0].Value);
                        }
                        else
                        {
                            importedvoucherid = -1;
                        }

                        for (int j = 0; j < drgroupby.Count(); j++)
                        {
                            drgroupby[j]["是否导入"] = "Y";
                            drgroupby[j]["错误信息"] = "";
                            drgroupby[j]["凭证号"] = importedvoucherid;
                            drgroupby[j]["制单人"] = userid;
                        }

                    }
                    else
                    {
                        v_importfailurerows = v_importfailurerows + 1;
                        //回写凭证号及错误信息
                        for (int j = 0; j < drgroupby.Count(); j++)
                        {
                            drgroupby[j]["是否导入"] = "N";
                            drgroupby[j]["错误信息"] = glcvoucher.strErrMessage;
                        }
                    }

                    //删除导入接口表数据并关闭数据库连接
                    strSql = "DELETE FROM tempdb.dbo.cus_gl_accvouchers ";
                    rs = conn.Execute(strSql, out rsaffected, -1);

                    //复制已导入数据到返回数据表中
                    for (int k = 0; k < drgroupby.Count(); k++)
                    {
                        dsreturnvouchers.Tables["GLVouchers"].ImportRow(drgroupby[k]);
                    }
                }

                //执行PerformStep()函数
                importdataprogressBar.PerformStep();
                string str = Math.Round((100 * (i + 1.0) / dtdistinct.Rows.Count), 2).ToString("#0.00 ") + "%";
                Font font = new Font("Times New Roman", (float)10, FontStyle.Regular);
                PointF pt = new PointF(this.importdataprogressBar.Width / 2 - 17, this.importdataprogressBar.Height / 2 - 7);
                g.DrawString(str, font, Brushes.Blue, pt);

            }

            //返回数据导入是否成功标志
            importsuccessrows = v_importsuccessrows;
            importfailurerows = v_importfailurerows;
            conn.Close();
            if (v_importfailurerows != 0)
            {
                return false;
            }
            else
            {
                return true;
            }

        }


        //public bool ReceiptNoteimport01(UFSoft.U8.Framework.LoginContext.UserData u8userdata, DataSet dsimportedreceiptnotes, out int importsuccessrows, out int importfailurerows, out DataSet dsreturnreceiptnotes, out string errmsg)
        //{
        //    /*测试数据
        //    int v_importsuccessrows = 0, v_importfailurerows = 0;
        //    string[] displayresult;
        //    //采购入库单导入
        //    //第一步：构造u8login对象并登陆(引用U8API类库中的Interop.U8Login.dll),如果当前环境中有login对象则可以省去第一步
        //    interU8lg::U8Login.clsLogin u8Login = new interU8lg::U8Login.clsLogin();
        //    String sSubId = Pubvar.gu8userdata.cSubID;              // "AS";
        //    String sAccID = Pubvar.gu8userdata.AccID;               // "(default)@999"
        //    String sYear = Pubvar.gu8userdata.iYear;                 //"2014";
        //    String sUserID = Pubvar.gu8userdata.UserId;             //"demo";
        //    String sPassword = Pubvar.gu8userdata.Password;         // "";
        //    String sDate = Pubvar.gu8userdata.operDate;             //"2014-12-11";
        //    String sServer = Pubvar.gu8userdata.AppServer;          // "UF8125";
        //    String sSerial = "";

        //    if (!u8Login.Login(ref sSubId, ref sAccID, ref sYear, ref sUserID, ref sPassword, ref sDate, ref sServer, ref sSerial))
        //    {
        //        MessageBox.Show("登陆失败，原因：" + u8Login.ShareString);
        //        Marshal.FinalReleaseComObject(u8Login);

        //        dsreturnreceiptnotes = null;
        //        importsuccessrows = 0;
        //        importfailurerows = 0;
        //        errmsg = "";
        //        return false;
        //    }

        //    //第二步：构造环境上下文对象，传入login，并按需设置其它上下文参数
        //    U8EnvContext envContext = new U8EnvContext();
        //    envContext.U8Login = u8Login;

        //    //第三步：设置API地址标识(Url)：当前API：添加新单据的地址标识为：U8API/PuStoreIn/Add
        //    U8ApiAddress BestU8ApiAddress = new U8ApiAddress("U8API/PuStoreIn/Add");

        //    //第四步：构造APIBroker
        //    U8ApiBroker broker = new U8ApiBroker(BestU8ApiAddress, envContext);

        //    //第五步：API参数赋值

        //    //给普通参数sVouchType赋值。此参数的数据类型为System.String，此参数按值传递，表示单据类型：01
        //    broker.AssignNormalValue("sVouchType", Convert.ToString("01"));

        //    BusinessObject DomHead = broker.GetBoParam("DomHead");
        //    DomHead.RowCount = 1; //设置BO对象(表头)行数，只能为一行
        //    //给BO对象(表头)的字段赋值，值可以是真实类型，也可以是无类型字符串.以下代码示例只设置第一行值。各字段定义详见API服务接口定义
        //    //****************************** 以下是必输字段 ****************************
        //    DomHead[0]["id"] = "1000000410"; //主关键字段，int类型
        //    //DomHead[0]["bomfirst"] = "0"; //委外期初标志，string类型
        //    DomHead[0]["ccode"] = "testimp0006"; //入库单号，string类型
        //    DomHead[0]["ddate"] = "2015-01-12"; //入库日期，DateTime类型
        //    //DomHead[0]["iverifystate"] = "0"; //iverifystate，int类型
        //    //DomHead[0]["iswfcontrolled"] = "0"; //iswfcontrolled，int类型
        //    //DomHead[0]["cvenabbname"] = "辰环手机配件"; //供货单位，string类型
        //    //DomHead[0]["cbustype"] = "普通采购"; //业务类型，int类型
        //    DomHead[0]["cmaker"] = "demo"; //制单人，string类型
        //    DomHead[0]["iexchrate"] = "1.00"; //汇率，double类型
        //    DomHead[0]["cexch_name"] = "人民币"; //币种，string类型
        //    //DomHead[0]["ufts"] = ""; //时间戳，string类型
        //    //DomHead[0]["bpufirst"] = "0"; //采购期初标志，string类型
        //    DomHead[0]["cvencode"] = "01002"; //供货单位编码，string类型
        //    DomHead[0]["cvouchtype"] = "01"; //单据类型，string类型
        //    DomHead[0]["cwhcode"] = "04"; //仓库编码，string类型
        //    //DomHead[0]["brdflag"] = "1"; //收发标志，int类型
        //    DomHead[0]["csource"] = "采购订单"; //单据来源，int类型
        //    //DomHead[0]["iflowid"] = ""; //流程模式ID，string类型
        //    //DomHead[0]["cflowname"] = ""; //流程模式描述，string类型
        //    //DomHead[0]["csysbarcode"] = ""; //单据条码，string类型
        //    //DomHead[0]["chinvsn"] = ""; //序列号，string类型
        //    DomHead[0]["cordercode"] = "0000000042";

        //    BusinessObject domBody = broker.GetBoParam("domBody");
        //    domBody.RowCount = 10; //设置BO对象行数

        //    //****************************** 以下是必输字段 ****************************
        //    domBody[0]["autoid"] = "1000001237"; //主关键字段，int类型
        //    domBody[0]["id"] = "1000000410"; //与收发记录主表关联项，int类型
        //    domBody[0]["cinvcode"] = "01019002063"; //存货编码，string类型

        //    //domBody[0]["cinvm_unit"] = ""; //主计量单位，string类型
        //    domBody[0]["iquantity"] = "3.00"; //数量，double类型
        //    domBody[0]["editprop"] = "A"; //编辑属性：A表新增，M表修改，D表删除，string类型

        //    //domBody[0]["iMatSettleState"] = new int(); //iMatSettleState，int类型
        //    //domBody[0]["creworkmocode"] = ""; //返工订单号，string类型
        //    //domBody[0]["ireworkmodetailsid"] = ""; //返工订单子表标识，string类型
        //    //domBody[0]["iproducttype"] = ""; //产出品类型，string类型
        //    //domBody[0]["cmaininvcode"] = ""; //对应主产品，string类型
        //    //domBody[0]["imainmodetailsid"] = ""; //主产品订单子表标识，string类型
        //    //domBody[0]["isharematerialfee"] = ""; //分摊材料费，string类型
        //    //domBody[0]["cinvouchtype"] = ""; //对应入库单类型，string类型
        //    //domBody[0]["idebitids"] = ""; //借入借出单子表id，string类型
        //    //domBody[0]["imergecheckautoid"] = ""; //检验单子表ID，string类型
        //    //domBody[0]["outcopiedquantity"] = ""; //已复制数量，string类型
        //    //domBody[0]["iOldPartId"] = ""; //降级前物料编码，string类型
        //    //domBody[0]["fOldQuantity"] = ""; //降级前数量，string类型
        //    //domBody[0]["cbsysbarcode"] = ""; //单据行条码，string类型
        //    //domBody[0]["cbmemo"] = ""; //备注，string类型
        //    //domBody[0]["iFaQty"] = ""; //转资产数量，string类型
        //    //domBody[0]["isTax"] = ""; //累计结算税额，string类型
        //    //domBody[0]["irowno"] = ""; //行号，string类型
        //    //domBody[0]["cbinvsn"] = ""; //序列号，string类型
        //    //domBody[0]["strowguid"] = ""; //rowguid，string类型
        //    //domBody[0]["cplanlotcode"] = ""; //计划批号，string类型
        //    //domBody[0]["taskguid"] = ""; //taskguid，string类型
        //    //domBody[0]["bgift"] = ""; //赠品，string类型

        //    //给普通参数domPosition赋值。此参数的数据类型为System.Object，此参数按引用传递，表示货位：传空
        //    //broker.AssignNormalValue("domPosition", new System.Object());
        //    broker.AssignNormalValue("domPosition", null);

        //    //该参数errMsg为OUT型参数，由于其数据类型为System.String，为一般值类型，因此不必传入一个参数变量。在API调用返回时，可以通过GetResult("errMsg")获取其值

        //    //给普通参数cnnFrom赋值。此参数的数据类型为ADODB.Connection，此参数按引用传递，表示连接对象,如果由调用方控制事务，则需要设置此连接对象，否则传空
        //    //broker.AssignNormalValue("cnnFrom", new ADODB.Connection());
        //    broker.AssignNormalValue("cnnFrom", null);

        //    //该参数VouchId为INOUT型普通参数。此参数的数据类型为System.String，此参数按值传递。在API调用返回时，可以通过GetResult("VouchId")获取其值
        //    broker.AssignNormalValue("VouchId", Convert.ToString(""));

        //    //该参数domMsg为OUT型参数，由于其数据类型为MSXML2.IXMLDOMDocument2，非一般值类型，因此必须传入一个参数变量。在API调用返回时，可以直接使用该参数
        //    //MSXML2.IXMLDOMDocument2 domMsg = new MSXML2.IXMLDOMDocument2();
        //    MSXML2.DOMDocumentClass domMsg = new MSXML2.DOMDocumentClass();
        //    broker.AssignNormalValue("domMsg", (IXMLDOMDocument2)domMsg);

        //    //给普通参数bCheck赋值。此参数的数据类型为System.Boolean，此参数按值传递，表示是否控制可用量。
        //    broker.AssignNormalValue("bCheck", false);

        //    //给普通参数bBeforCheckStock赋值。此参数的数据类型为System.Boolean，此参数按值传递，表示检查可用量
        //    broker.AssignNormalValue("bBeforCheckStock", false);

        //    //给普通参数bIsRedVouch赋值。此参数的数据类型为System.Boolean，此参数按值传递，表示是否红字单据
        //    broker.AssignNormalValue("bIsRedVouch", false);

        //    //给普通参数sAddedState赋值。此参数的数据类型为System.String，此参数按值传递，表示传空字符串
        //    broker.AssignNormalValue("sAddedState", Convert.ToString(""));

        //    //给普通参数bReMote赋值。此参数的数据类型为System.Boolean，此参数按值传递，表示是否远程：转入false
        //    broker.AssignNormalValue("bReMote", false);

        //    //第六步：调用API
        //    if (!broker.Invoke())
        //    {
        //        //错误处理
        //        Exception apiEx = broker.GetException();
        //        if (apiEx != null)
        //        {
        //            if (apiEx is MomSysException)
        //            {
        //                MomSysException sysEx = apiEx as MomSysException;
        //                //importresule.Text = "系统异常：" + sysEx.Message + "\n\r";
        //                //todo:异常处理
        //            }
        //            else if (apiEx is MomBizException)
        //            {
        //                MomBizException bizEx = apiEx as MomBizException;
        //                //importresule.Text = "API异常：" + bizEx.Message + "\n\r";
        //                //todo:异常处理
        //            }
        //            //异常原因
        //            String exReason = broker.GetExceptionString();
        //            if (exReason.Length != 0)
        //            {
        //                //importresule.Text = "其他异常原因：" + exReason + "\n\r";
        //            }
        //        }
        //        //结束本次调用，释放API资源
        //        broker.Release();
        //        dsreturnreceiptnotes = null;
        //        importsuccessrows = 0;
        //        importfailurerows = 0;
        //        errmsg = "";
        //        return false;
        //    }

        //    //第七步：获取返回结果

        //    //获取普通返回值。此返回值数据类型为System.Boolean，此参数按值传递，表示返回值:true:成功,false:失败
        //    System.Boolean result = Convert.ToBoolean(broker.GetReturnValue());
        //    if (result)
        //    {
        //        v_importsuccessrows = v_importsuccessrows + 1;

        //    }
        //    else
        //    {
        //        v_importfailurerows = v_importfailurerows + 1;
        //    }

        //    //获取out/inout参数值

        //    //获取普通OUT参数errMsg。此返回值数据类型为System.String，在使用该参数之前，请判断是否为空
        //    System.String errMsgRet = broker.GetResult("errMsg") as System.String;

        //    //获取普通INOUT参数VouchId。此返回值数据类型为System.String，在使用该参数之前，请判断是否为空
        //    System.String VouchIdRet = broker.GetResult("VouchId") as System.String;
        //    System.String VouchIdRetttt = broker.GetResult("sVouchType") as System.String;
        //    System.String T1 = broker.GetResult("sAddedState") as System.String;

        //    //获取普通OUT参数domMsg。此返回值数据类型为MSXML2.IXMLDOMDocument2，在使用该参数之前，请判断是否为空
        //    MSXML2.IXMLDOMDocument2 domMsgRet = (MSXML2.DOMDocument)(broker.GetResult("domMsg"));
        //    //BusinessObject vdomBody = broker.GetBoParam("domBody");
        //    //BusinessObject vDomHead = broker.GetBoParam("DomHead");

        //    //第八步 ： 结束本次调用，释放API资源
        //    broker.Release();
        //    dsreturnreceiptnotes = null;
        //    importsuccessrows = 0;
        //    importfailurerows = 0;
        //    errmsg = "";
        //    return true;
        //    */



        //    int v_importsuccessrows = 0, v_importfailurerows = 0;
        //    string v_errmsg = "";

        //    SqlConnection conn = new SqlConnection();
        //    SqlDataAdapter apdata = new SqlDataAdapter();
        //    SqlCommand sqlcmd = new SqlCommand();
        //    DataTable dbaccinfo = new DataTable();
        //    DataSet poheadds = new DataSet(), polinesds = new DataSet();
        //    int pos = u8userdata.ConnString.IndexOf(";");
        //    conn.ConnectionString = u8userdata.ConnString.Remove(0, pos + 1);
        //    conn.Open();//连接数据库  
        //    sqlcmd.Connection = conn;

        //    //第一步：构造u8login对象并登陆(引用U8API类库中的Interop.U8Login.dll),如果当前环境中有login对象则可以省去第一步
        //    interU8lg::U8Login.clsLogin u8Login = new interU8lg::U8Login.clsLogin();
        //    String sSubId = u8userdata.cSubID;              // "AS";
        //    String sAccID = u8userdata.AccID;               // "(default)@999"
        //    String sYear = u8userdata.iYear;                 //"2014";
        //    String sUserID = u8userdata.UserId;             //"demo";
        //    String sPassword = u8userdata.Password;         // "";
        //    String sDate = u8userdata.operDate;             //"2014-12-11";
        //    String sServer = u8userdata.AppServer;          // "UF8125";
        //    String sSerial = "";
        //    if (!u8Login.Login(ref sSubId, ref sAccID, ref sYear, ref sUserID, ref sPassword, ref sDate, ref sServer, ref sSerial))
        //    {
        //        Marshal.FinalReleaseComObject(u8Login);
        //        v_errmsg = "数据导入登陆失败，原因：" + u8Login.ShareString;
        //        //返回数据导入是否成功标志
        //        importsuccessrows = v_importsuccessrows;
        //        importfailurerows = v_importfailurerows;
        //        dsreturnreceiptnotes = dsimportedreceiptnotes;
        //        errmsg = v_errmsg;
        //        conn.Close();
        //        return false;
        //    }


        //    dsreturnreceiptnotes = dsimportedreceiptnotes.Clone();
        //    DataTable dtdistinct = dsimportedreceiptnotes.Tables["ReceiptNotes"].DefaultView.ToTable(true, new string[] { "单据ID" });
        //    string vougroupby = "";
        //    //设置progressbar步长并显示百分比
        //    importdataprogressBar.Minimum = 0;   // 设置进度条最小值.
        //    importdataprogressBar.Value = 1;    // 设置进度条初始值
        //    importdataprogressBar.Step = 1;     // 设置每次增加的步长
        //    importdataprogressBar.Maximum = dtdistinct.Rows.Count;// 设置进度条最大值.
        //    Graphics g = this.importdataprogressBar.CreateGraphics();

        //    //第五步：API单据值及参数赋值： 根据dataset中导入数据分组循环导入U8系统，如有采购订单则以采购订单作为分组条件，否则则以分组标识作为分组条件。
        //    for (int i = 0; i < dtdistinct.Rows.Count; i++)   //分组开始 
        //    {
        //        string v_receiptnotnumber = "";

        //        //第二步：构造环境上下文对象，传入login，并按需设置其它上下文参数
        //        U8EnvContext envContext = new U8EnvContext();
        //        envContext.U8Login = u8Login;

        //        //第三步：设置API地址标识(Url)：当前API：添加新单据的地址标识为：U8API / PuStoreIn / Add
        //        U8ApiAddress BestU8ApiAddress = new U8ApiAddress("U8API/PuStoreIn/Add");
        //        //第四步：构造APIBroker
        //        U8ApiBroker broker = new U8ApiBroker(BestU8ApiAddress, envContext);

        //        vougroupby = dtdistinct.Rows[i]["单据ID"].ToString();
        //        string filterestr = "";
        //        //当凭证ID为空或""时的特殊处理
        //        if (!string.IsNullOrEmpty(vougroupby))
        //        {
        //            filterestr = "单据ID = " + "'" + vougroupby + "'";
        //        }
        //        else
        //        {
        //            filterestr = "单据ID IS NULL  OR 单据ID ='" + "'";
        //        }
        //        DataRow[] drgroupby = dsimportedreceiptnotes.Tables["ReceiptNotes"].Select(filterestr);
        //        if ((!string.IsNullOrEmpty(drgroupby[0]["单据号"].ToString())) || ((!string.IsNullOrEmpty(drgroupby[0]["是否导入"].ToString())) && (drgroupby[0]["是否导入"].ToString() == "N")) || (string.IsNullOrEmpty(drgroupby[0]["单据ID"].ToString())))
        //        {
        //            //复制已导入的数据到返回数据表中
        //            for (int k = 0; k < drgroupby.Count(); k++)
        //            {
        //                dsreturnreceiptnotes.Tables["ReceiptNotes"].ImportRow(drgroupby[k]);
        //            }
        //        }
        //        else
        //        {   //这组数据中需要保证订单号唯一
        //            string v_ponumber = drgroupby[0]["订单号"].ToString();
        //            bool podif = false;
        //            for (int j = 0; j < drgroupby.Count(); j++)
        //            {
        //                if (v_ponumber != drgroupby[j]["订单号"].ToString())
        //                {
        //                    podif = true;
        //                    break;
        //                }
        //            }
        //            if (podif)
        //            {
        //                v_importfailurerows = v_importfailurerows + 1;
        //                //回写错误信息
        //                for (int j = 0; j < drgroupby.Count(); j++)
        //                {
        //                    drgroupby[j]["是否导入"] = "N";
        //                    drgroupby[j]["错误信息"] = "一张采购入库单中存在不同订单号，请检查!";
        //                }
        //                //复制已导入的数据到返回数据表中
        //                for (int k = 0; k < drgroupby.Count(); k++)
        //                {
        //                    dsreturnreceiptnotes.Tables["ReceiptNotes"].ImportRow(drgroupby[k]);
        //                }
        //            }
        //            else
        //            {

        //                if (!string.IsNullOrEmpty(v_ponumber))
        //                {
        //                    //获取订单头及订单行信息
        //                    sqlcmd.CommandText = "SELECT * FROM dbo.PO_Pomain WHERE cPOID ='" + v_ponumber + "'";
        //                    apdata.SelectCommand = sqlcmd;
        //                    apdata.Fill(poheadds);
        //                    sqlcmd.CommandText = "SELECT * FROM dbo.PO_Podetails WHERE POID =" + poheadds.Tables[0].Rows[0]["POID"].ToString();
        //                    apdata.SelectCommand = sqlcmd;
        //                    apdata.Fill(polinesds);
        //                }

        //                //API单据值赋值
        //                #region 
        //                //设置BO对象(表头)行数，只能为一行

        //                BusinessObject DomHead = broker.GetBoParam("DomHead");
        //                DataTable rdmainid = new DataTable(), rdmaincode = new DataTable(), rdlineid = new DataTable(), rdcptcode = new DataTable();
        //                DomHead.RowCount = 1;

        //                sqlcmd.CommandText = "SELECT MAX(ID)+1 FROM dbo.RdRecord01 ";
        //                apdata.SelectCommand = sqlcmd;
        //                apdata.Fill(rdmainid);
        //                //入库单主表主关键ID "1000000404"
        //                DomHead[0]["id"] = rdmainid.Rows[0][0].ToString();
        //                sqlcmd.CommandText = "SELECT RIGHT('0000000000' + CONVERT(VARCHAR(10), max(ccode) + 1),10) FROM dbo.RdRecord01 ";
        //                apdata.SelectCommand = sqlcmd;
        //                apdata.Fill(rdmaincode);
        //                //入库单编号
        //                DomHead[0]["ccode"] = rdmaincode.Rows[0][0].ToString();
        //                v_receiptnotnumber = rdmaincode.Rows[0][0].ToString();
        //                //入库日期"2015-01-12"
        //                DomHead[0]["ddate"] = drgroupby[0]["单据日期"].ToString();
        //                //制单人    
        //                DomHead[0]["cmaker"] = u8userdata.UserId;
        //                //供应商、部门、业务类型、单据来源编码、汇率、币种
        //                if (!string.IsNullOrEmpty(v_ponumber))
        //                {
        //                    DomHead[0]["cvencode"] = poheadds.Tables[0].Rows[0]["cVenCode"].ToString();
        //                    DomHead[0]["cdepcode"] = poheadds.Tables[0].Rows[0]["cDepCode"].ToString();
        //                    DomHead[0]["cbustype"] = poheadds.Tables[0].Rows[0]["cBusType"].ToString();
        //                    DomHead[0]["csource"] = "采购订单";  //委外订单
        //                    DomHead[0]["iexchrate"] = poheadds.Tables[0].Rows[0]["nflat"].ToString();
        //                    DomHead[0]["cexch_name"] = poheadds.Tables[0].Rows[0]["cexch_name"].ToString();
        //                }
        //                else
        //                {
        //                    DomHead[0]["cvencode"] = drgroupby[0]["供应商编码"].ToString();
        //                    DomHead[0]["cdepcode"] = drgroupby[0]["部门编码"].ToString();
        //                    DomHead[0]["csource"] = drgroupby[0]["单据来源"].ToString(); //"库存";采购订单，委外订单
        //                    DomHead[0]["cbustype"] = drgroupby[0]["业务类型"].ToString(); //"普通采购";委外加工
        //                    DomHead[0]["iexchrate"] = drgroupby[0]["汇率"].ToString();
        //                    DomHead[0]["cexch_name"] = drgroupby[0]["币种"].ToString();
        //                }
        //                //单据类型这里固定是 01- 采购入库单
        //                DomHead[0]["cvouchtype"] = "01";
        //                ///仓库编码
        //                DomHead[0]["cwhcode"] = drgroupby[0]["仓库编码"].ToString();
        //                ////收发标志这里固定是收标志
        //                DomHead[0]["brdflag"] = "1";
        //                //采购及入库类别编码
        //                if (!string.IsNullOrEmpty(v_ponumber))
        //                {
        //                    sqlcmd.CommandText = "SELECT a.cPTName,a.cRdCode,b.cRdName  FROM dbo.PurchaseType AS a inner join dbo.Rd_Style AS b on a.cRdCode = b.cRdCode WHERE a.cPTCode='" + poheadds.Tables[0].Rows[0]["cPTCode"].ToString() + "'";
        //                    apdata.SelectCommand = sqlcmd;
        //                    apdata.Fill(rdcptcode);
        //                    DomHead[0]["cptcode"] = poheadds.Tables[0].Rows[0]["cPTCode"].ToString();
        //                    DomHead[0]["crdcode"] = rdcptcode.Rows[0]["cRdCode"].ToString();
        //                    DomHead[0]["cptname"] = rdcptcode.Rows[0]["cPTName"].ToString();
        //                    DomHead[0]["crdname"] = rdcptcode.Rows[0]["cRdName"].ToString();
        //                }
        //                else
        //                {
        //                    DomHead[0]["cptcode"] = "";
        //                    DomHead[0]["crdcode"] = rdcptcode.Rows[0]["cRdCode"].ToString();
        //                    DomHead[0]["cptname"] = rdcptcode.Rows[0]["cPTName"].ToString();
        //                    DomHead[0]["crdname"] = rdcptcode.Rows[0]["cRdName"].ToString();
        //                }

        //                if (!string.IsNullOrEmpty(v_ponumber))
        //                {
        //                    DomHead[0]["cordercode"] = v_ponumber;                                       //订单号，string类型
        //                    DomHead[0]["itaxrate"] = poheadds.Tables[0].Rows[0]["itaxrate"].ToString();
        //                    DomHead[0]["ipurorderid"] = poheadds.Tables[0].Rows[0]["POID"].ToString();   //采购订单ID，string类型
        //                }
        //                else
        //                {
        //                    DomHead[0]["cordercode"] = "";
        //                }

        //                BusinessObject domBody = broker.GetBoParam("domBody");
        //                domBody.RowCount = 10;
        //                //domBody.RowCount = drgroupby.Count();
        //                sqlcmd.CommandText = "SELECT MAX(autoid) FROM dbo.rdrecords01 ";
        //                apdata.SelectCommand = sqlcmd;
        //                apdata.Fill(rdlineid);

        //                Int32 v_linesidmax = Convert.ToInt32(rdlineid.Rows[0][0]);
        //                string filter = "";
        //                DataRow[] drpolines;
        //                bool itemexist = true;
        //                for (int j = 0; j < drgroupby.Count(); j++)
        //                {
        //                    domBody[j]["autoid"] = v_linesidmax + 1;                                  //"1000001229";   //主关键字段，int类型
        //                    domBody[j]["id"] = DomHead[0]["id"].ToString();                          //"1000000404"; //与收发记录主表关联项，int类型
        //                    domBody[j]["cinvcode"] = drgroupby[j]["存货编码"].ToString();     //"01019002082"; //存货编码，string类型
        //                    //domBody[j]["cinvm_unit"] = "PCS"; //主计量单位，string类型
        //                    //domBody[j]["cinvname"] = "主板"; //存货名称，string类型

        //                    domBody[j]["iquantity"] = drgroupby[j]["入库数量"].ToString();           // "777.00"; //数量，double类型
        //                    domBody[j]["editprop"] = "A";                                            //编辑属性：A表新增，M表修改，D表删除，string类型
        //                    domBody[j]["irowno"] = j + 1;                                             //行号，string类型

        //                    filter = "cInvCode = '" + drgroupby[j]["存货编码"].ToString() + "'";
        //                    if (!string.IsNullOrEmpty(v_ponumber))
        //                    {
        //                        drpolines = polinesds.Tables[0].Select(filter);
        //                        if (drpolines.Length == 0)
        //                        {
        //                            itemexist = false;
        //                            break;
        //                        }
        //                        else
        //                        {
        //                            //获取采购订单中价格及金额信息
        //                            domBody[j]["itaxrate"] = poheadds.Tables[0].Rows[0]["itaxrate"].ToString(); //税率，double类型

        //                            domBody[j]["ioritaxcost"] = drpolines[0]["iTaxPrice"].ToString();  //原币含税单价，double类型
        //                            domBody[j]["ioricost"] = drpolines[0]["iUnitPrice"].ToString();     //原币单价，double类型
        //                            domBody[j]["iorimoney"] = Math.Round(Convert.ToDouble(drgroupby[j]["入库数量"]) * Convert.ToDouble(drpolines[0]["iUnitPrice"]), 2).ToString();    //原币金额，double类型
        //                            domBody[j]["ioritaxprice"] = Math.Round(Convert.ToDouble(drgroupby[j]["入库数量"]) * Convert.ToDouble(drpolines[0]["iUnitPrice"]) * Convert.ToDouble(poheadds.Tables[0].Rows[0]["itaxrate"]) / 100.00, 2).ToString(); //原币税额，double类型
        //                            domBody[j]["iorisum"] = (Convert.ToDouble(domBody[0]["iorimoney"]) + Convert.ToDouble(domBody[0]["ioritaxprice"])).ToString(); //原币价税合计，double类型

        //                            domBody[j]["iunitcost"] = drpolines[0]["iNatUnitPrice"].ToString(); //本币无税单价 ，double类型
        //                            domBody[j]["iprice"] = Math.Round(Convert.ToDouble(drgroupby[j]["入库数量"]) * Convert.ToDouble(drpolines[0]["iNatUnitPrice"]), 2).ToString(); //本币金额，double类型
        //                            domBody[j]["iaprice"] = Math.Round(Convert.ToDouble(drgroupby[j]["入库数量"]) * Convert.ToDouble(drpolines[0]["iNatUnitPrice"]), 2).ToString();  //暂估金额，double类型
        //                            //domBody[j]["cbatch"] = "001"; //批号，string类型
        //                            domBody[j]["iposid"] = drpolines[0]["id"].ToString(); //订单子表ID，int类型
        //                            domBody[j]["facost"] = drpolines[0]["iNatUnitPrice"].ToString(); //暂估单价，double类型
        //                            domBody[j]["inquantity"] = drpolines[0]["iQuantity"].ToString(); //应收数量，double类型


        //                            domBody[j]["itaxprice"] = Math.Round(Convert.ToDouble(drgroupby[j]["入库数量"]) * Convert.ToDouble(drpolines[0]["iNatUnitPrice"]) * Convert.ToDouble(poheadds.Tables[0].Rows[0]["itaxrate"]) / 100.00, 2).ToString(); ; //本币税额，double类型
        //                            domBody[j]["isum"] = (Convert.ToDouble(domBody[j]["iprice"]) + Convert.ToDouble(domBody[j]["itaxprice"])).ToString();  //本币价税合计，double类型
        //                            domBody[j]["cpoid"] = v_ponumber; //订单号，string类型

        //                        }
        //                    }

        //                }
        //                #endregion
        //                //导入数据中item不存在PO中
        //                if (!itemexist)
        //                {
        //                    v_importfailurerows = v_importfailurerows + 1;
        //                    //回写错误信息
        //                    for (int j = 0; j < drgroupby.Count(); j++)
        //                    {
        //                        drgroupby[j]["是否导入"] = "N";
        //                        drgroupby[j]["错误信息"] = "采购入库单中入库商品在采购订单中不存在，请检查!";
        //                    }
        //                    //复制已导入的数据到返回数据表中
        //                    for (int k = 0; k < drgroupby.Count(); k++)
        //                    {
        //                        dsreturnreceiptnotes.Tables["ReceiptNotes"].ImportRow(drgroupby[k]);
        //                    }
        //                }
        //                else
        //                {
        //                    //API 参数赋值
        //                    #region
        //                    //给普通参数sVouchType赋值。此参数的数据类型为System.String，此参数按值传递，表示单据类型：01
        //                    broker.AssignNormalValue("sVouchType", Convert.ToString("01"));
        //                    //给普通参数domPosition赋值。此参数的数据类型为System.Object，此参数按引用传递，表示货位：传空
        //                    broker.AssignNormalValue("domPosition", null); //broker.AssignNormalValue("domPosition", new System.Object());
        //                    //该参数errMsg为OUT型参数，由于其数据类型为System.String，为值类型，因此不必传入参数变量。在API调用返回时，可以通过GetResult("errMsg")获取其值
        //                    //给普通参数cnnFrom赋值。此参数的数据类型为ADODB.Connection，此参数按引用传递，表示连接对象,如果由调用方控制事务，则需要设置此连接对象，否则传空
        //                    broker.AssignNormalValue("cnnFrom", null); //broker.AssignNormalValue("cnnFrom", new ADODB.Connection());
        //                    //该参数VouchId为INOUT型普通参数。此参数的数据类型为System.String，此参数按值传递。在API调用返回时，可以通过GetResult("VouchId")获取其值
        //                    broker.AssignNormalValue("VouchId", Convert.ToString(""));
        //                    //该参数domMsg为OUT型参数，由于其数据类型为MSXML2.IXMLDOMDocument2，非一般值类型，因此必须传入一个参数变量。在API调用返回时，可以直接使用该参数.
        //                    //无法直接创建接口实例，需要做类型转换 。//MSXML2.IXMLDOMDocument2 domMsg = new MSXML2.IXMLDOMDocument2();
        //                    MSXML2.DOMDocumentClass domMsg = new MSXML2.DOMDocumentClass();
        //                    broker.AssignNormalValue("domMsg", (IXMLDOMDocument2)domMsg);
        //                    //给普通参数bCheck赋值。此参数的数据类型为System.Boolean，此参数按值传递，表示是否控制可用量。
        //                    broker.AssignNormalValue("bCheck", false);
        //                    //给普通参数bBeforCheckStock赋值。此参数的数据类型为System.Boolean，此参数按值传递，表示检查可用量
        //                    broker.AssignNormalValue("bBeforCheckStock", false);
        //                    //给普通参数bIsRedVouch赋值。此参数的数据类型为System.Boolean，此参数按值传递，表示是否红字单据
        //                    broker.AssignNormalValue("bIsRedVouch", false);
        //                    //给普通参数sAddedState赋值。此参数的数据类型为System.String，此参数按值传递，表示传空字符串
        //                    broker.AssignNormalValue("sAddedState", Convert.ToString(""));
        //                    //给普通参数bReMote赋值。此参数的数据类型为System.Boolean，此参数按值传递，表示是否远程：传入false
        //                    broker.AssignNormalValue("bReMote", false);
        //                    #endregion
        //                    //第六步：调用API
        //                    #region
        //                    if (!broker.Invoke())
        //                    {
        //                        //错误处理
        //                        Exception apiEx = broker.GetException();
        //                        if (apiEx != null)
        //                        {
        //                            if (apiEx is MomSysException)
        //                            {
        //                                MomSysException sysEx = apiEx as MomSysException;
        //                                v_errmsg = "系统异常：" + sysEx.Message + "\n\r";

        //                            }
        //                            else if (apiEx is MomBizException)
        //                            {
        //                                MomBizException bizEx = apiEx as MomBizException;
        //                                v_errmsg = "API异常：" + bizEx.Message + "\n\r";

        //                            }
        //                            //异常原因
        //                            String exReason = broker.GetExceptionString();
        //                            if (exReason.Length != 0)
        //                            {
        //                                v_errmsg = "其他异常原因：" + exReason + "\n\r";
        //                            }
        //                        }
        //                        //结束本次调用，释放API资源
        //                        //broker.Release();
        //                    }
        //                    #endregion
        //                    //第七步：获取返回结果
        //                    #region
        //                    //获取普通返回值。此返回值数据类型为System.Boolean，此参数按值传递，表示返回值:true:成功,false:失败
        //                    System.Boolean result = Convert.ToBoolean(broker.GetReturnValue());
        //                    //获取out/inout参数值
        //                    //获取普通OUT参数errMsg。此返回值数据类型为System.String，在使用该参数之前，请判断是否为空
        //                    errmsg = (System.String)broker.GetResult("errMsg");
        //                    //获取普通INOUT参数VouchId。此返回值数据类型为System.String，在使用该参数之前，请判断是否为空
        //                    System.String v_vouchid = (System.String)broker.GetResult("VouchId");

        //                    //获取普通OUT参数domMsg。此返回值数据类型为MSXML2.IXMLDOMDocument2，在使用该参数之前，请判断是否为空
        //                    //MSXML2.IXMLDOMDocument2 domMsgRet = (MSXML2.DOMDocument)(broker.GetResult("domMsg"));
        //                    //BusinessObject vdomBody = broker.GetBoParam("domBody");
        //                    //BusinessObject vdomHead = broker.GetBoParam("DomHead");
        //                    #endregion
        //                    //第八步 ： 结束本次调用，释放API资源
        //                    #region
        //                    broker.Release();

        //                    if (result)
        //                    {
        //                        v_importsuccessrows = v_importsuccessrows + 1;
        //                        //回写信息
        //                        for (int j = 0; j < drgroupby.Count(); j++)
        //                        {
        //                            drgroupby[j]["是否导入"] = "Y";
        //                            drgroupby[j]["错误信息"] = "";
        //                            //drgroupby[j]["单据号"] = v_vouchid;
        //                            drgroupby[j]["单据号"] = v_receiptnotnumber;

        //                        }
        //                    }
        //                    else
        //                    {
        //                        v_importfailurerows = v_importfailurerows + 1;
        //                        //回写错误信息
        //                        for (int j = 0; j < drgroupby.Count(); j++)
        //                        {
        //                            drgroupby[j]["是否导入"] = "N";
        //                            drgroupby[j]["错误信息"] = errmsg;
        //                        }
        //                    }

        //                    //复制已导入的数据到返回数据表中
        //                    for (int k = 0; k < drgroupby.Count(); k++)
        //                    {
        //                        dsreturnreceiptnotes.Tables["ReceiptNotes"].ImportRow(drgroupby[k]);
        //                    }
        //                    #endregion
        //                }

        //            }
        //        }

        //        //执行进度条：PerformStep()函数
        //        importdataprogressBar.PerformStep();
        //        string str = Math.Round((100 * (i + 1.0) / dtdistinct.Rows.Count), 2).ToString("#0.00 ") + "%";
        //        Font font = new Font("Times New Roman", (float)10, FontStyle.Regular);
        //        PointF pt = new PointF(this.importdataprogressBar.Width / 2 - 17, this.importdataprogressBar.Height / 2 - 7);
        //        g.DrawString(str, font, Brushes.Blue, pt);

        //    } //分组结束

        //    //返回数据导入是否成功标志
        //    importsuccessrows = v_importsuccessrows;
        //    importfailurerows = v_importfailurerows;
        //    dsreturnreceiptnotes = dsimportedreceiptnotes;
        //    errmsg = v_errmsg;
        //    conn.Close();
        //    //结束本次调用，释放API资源
        //    //broker.Release();

        //    if (v_importfailurerows != 0)
        //    {
        //        return false;
        //    }
        //    else
        //    {
        //        return true;
        //    }

        //}

        public bool ReceiptNoteimport(UFSoft.U8.Framework.LoginContext.UserData u8userdata, DataSet dsimportedreceiptnotes, out int importsuccessrows, out int importfailurerows, out DataSet dsreturnreceiptnotes, out string errmsg)
        {
            int v_importsuccessrows = 0, v_importfailurerows = 0;
            string v_errmsg = "", v_groupby = "", v_receiptnotnumber = "",v_filterestr = "",v_bustype = "",v_source="", v_cinvcode = "",v_receiptnotedate="";
            bool v_exitflag = false;
            DataRow[] drgroupby, drorderlines;
            DataTable dtsql = new DataTable();
            DataSet orderhead = new DataSet(), orderlines = new DataSet(),dssql=new DataSet();
            interU8lg::U8Login.clsLogin u8Login;

            //构建导入事务数据库连接，复制导入数据集
            #region
            SqlConnection conn = new SqlConnection();
            SqlDataAdapter apdata = new SqlDataAdapter();
            SqlCommand sqlcmd = new SqlCommand();
            int pos = u8userdata.ConnString.IndexOf(";");
            conn.ConnectionString = u8userdata.ConnString.Remove(0, pos + 1);
            conn.Open();//连接数据库  
            sqlcmd.Connection = conn;
            dsreturnreceiptnotes = dsimportedreceiptnotes.Clone();
            #endregion
            //第一步：构造u8login对象并登陆(引用U8API类库中的Interop.U8Login.dll),如果当前环境中有login对象则可以省去第一步
            #region
            try
            {
                u8Login = new interU8lg::U8Login.clsLogin();
                String sSubId = u8userdata.cSubID;              // "AS";
                String sAccID = u8userdata.AccID;               // "(default)@999"
                String sYear = u8userdata.iYear;                 //"2014";
                String sUserID = u8userdata.UserId;             //"demo";
                String sPassword = u8userdata.Password;         // "";
                String sDate = u8userdata.operDate;             //"2014-12-11";
                String sServer = u8userdata.AppServer;          // "UF8125";
                String sSerial = "";
                if (!u8Login.Login(ref sSubId, ref sAccID, ref sYear, ref sUserID, ref sPassword, ref sDate, ref sServer, ref sSerial))
                {
                    Marshal.FinalReleaseComObject(u8Login);
                    v_errmsg = "数据导入登陆失败，原因：" + u8Login.ShareString;
                    //返回数据导入是否成功标志
                    importsuccessrows = v_importsuccessrows;
                    importfailurerows = v_importfailurerows;
                    dsreturnreceiptnotes = dsimportedreceiptnotes;
                    errmsg = v_errmsg;
                    conn.Close();
                    return false;
                }
            }
            catch (Exception ex)
            {
                v_errmsg = "数据导入登陆失败，原因：" + ex.Message;
                //返回数据导入是否成功标志
                importsuccessrows = v_importsuccessrows;
                importfailurerows = v_importfailurerows;
                dsreturnreceiptnotes = dsimportedreceiptnotes;
                errmsg = v_errmsg;
                conn.Close();
                return false;
            }
            #endregion
            //按照采购入库单导入数据模板，根据订单号进行分组，默认无订单号为一组.
            DataTable dtdistinct = dsimportedreceiptnotes.Tables["ReceiptNotes"].DefaultView.ToTable(true, new string[] { "订单号", "单据日期" });
            //进度条初始化
            #region
            //设置progressbar步长并显示百分比
            importdataprogressBar.Minimum = 0;                              // 设置进度条最小值.
            importdataprogressBar.Value = 1;                                // 设置进度条初始值
            importdataprogressBar.Step = 1;                                 // 设置每次增加的步长
            importdataprogressBar.Maximum = dtdistinct.Rows.Count;        // 设置进度条最大值.
            Graphics g = this.importdataprogressBar.CreateGraphics();
            #endregion
            //调用 U8 API ：单据及参数赋值并根据DataTable中分组数据循环导入U8系统.
            #region
            for (int i = 0; i < dtdistinct.Rows.Count; i++)   
            {
                //第二步：构造环境上下文对象，传入login，并按需设置其它上下文参数
                U8EnvContext envContext = new U8EnvContext();
                envContext.U8Login = u8Login;
                //第三步：设置API地址标识(Url)- 采购入库单： U8API / PuStoreIn / Add
                U8ApiAddress BestU8ApiAddress = new U8ApiAddress("U8API/PuStoreIn/Add");
                //第四步：构造APIBroker
                U8ApiBroker broker = new U8ApiBroker(BestU8ApiAddress, envContext);
                //按照分组标识数据获取导入数据模板该组数组值
                v_groupby = dtdistinct.Rows[i]["订单号"].ToString();
                v_receiptnotedate = dtdistinct.Rows[i]["单据日期"].ToString();
                if (!string.IsNullOrEmpty(v_groupby))
                {
                    v_filterestr = "订单号 = " + "'" + v_groupby + "' AND 单据日期 = '" + v_receiptnotedate + "'";
                }
                else
                {
                    v_filterestr = "(订单号 IS NULL  OR 订单号 ='" + "') AND 单据日期 = '" + v_receiptnotedate + "'";
                }
                drgroupby = dsimportedreceiptnotes.Tables["ReceiptNotes"].Select(v_filterestr);
                //执行导入模板数据初步校验 
                //1.针对已导入数据即“单据号”列不为空或“是否导入”列不为空，或“错误信息"列不为空则直接复制到返回datatable中，不参加数据导入API调用
                #region
                if ((!string.IsNullOrEmpty(drgroupby[0]["单据号"].ToString()))||(!string.IsNullOrEmpty(drgroupby[0]["是否导入"].ToString()))|| (!string.IsNullOrEmpty(drgroupby[0]["错误信息"].ToString())))
                {
                    //复制已导入的数据到返回数据表中
                    for (int k = 0; k < drgroupby.Count(); k++)
                    {
                        dsreturnreceiptnotes.Tables["ReceiptNotes"].ImportRow(drgroupby[k]);
                    }
                    continue;
                }
                #endregion
                //2.有订单号则单据来源必须为"采购订单"或"委外订单"，业务类型为普通采购或委外加工，无订单号则来源为"库存"。业务类型为普通采购。
                #region
                if (!string.IsNullOrEmpty(v_groupby))
                {
                    for (int j = 0; j < drgroupby.Count(); j++)
                    {
                        v_bustype = drgroupby[j]["业务类型"].ToString();
                        v_source  = drgroupby[j]["单据来源"].ToString();
                        if ((v_bustype != "普通采购") && (v_bustype != "委外加工"))
                        {
                            v_exitflag = true;
                            break;
                        }
                        if ((v_source != "采购订单") && (v_source != "委外订单"))
                        {
                            v_exitflag = true;
                            break;
                        }
                    }
                    if (v_exitflag)
                    {
                        v_importfailurerows = v_importfailurerows + 1;
                        //回写错误信息
                        for (int j = 0; j < drgroupby.Count(); j++)
                        {
                            drgroupby[j]["是否导入"] = "N";
                            drgroupby[j]["错误信息"] = "单据来源或业务类型错误，请检查!";
                        }
                        //复制已导入的数据到返回数据表中
                        for (int k = 0; k < drgroupby.Count(); k++)
                        {
                            dsreturnreceiptnotes.Tables["ReceiptNotes"].ImportRow(drgroupby[k]);
                        }
                        continue;
                    }
                }
                else
                {
                    v_filterestr = "订单号 IS NULL  OR 订单号 ='" + "'";
                    for (int j = 0; j < drgroupby.Count(); j++)
                    {
                        v_bustype = drgroupby[j]["业务类型"].ToString();
                        v_source = drgroupby[j]["单据来源"].ToString();
                        if (v_bustype != "普通采购") 
                        {
                            v_exitflag = true;
                            break;
                        }
                        if (v_source != "库存")
                        {
                            v_exitflag = true;
                            break;
                        }
                    }
                    if (v_exitflag)
                    {
                        v_importfailurerows = v_importfailurerows + 1;
                        //回写错误信息
                        for (int j = 0; j < drgroupby.Count(); j++)
                        {
                            drgroupby[j]["是否导入"] = "N";
                            drgroupby[j]["错误信息"] = "单据来源或业务类型错误，请检查!";
                        }
                        //复制已导入的数据到返回数据表中
                        for (int k = 0; k < drgroupby.Count(); k++)
                        {
                            dsreturnreceiptnotes.Tables["ReceiptNotes"].ImportRow(drgroupby[k]);
                        }
                        continue;
                    }
                }

                #endregion
                //3.导入数据中item是否存在PO中
                #region
                if ((!string.IsNullOrEmpty(v_groupby)) && (drgroupby[0]["业务类型"].ToString() == "普通采购"))
                {
                    for (int j = 0; j < drgroupby.Count(); j++)
                    {
                        v_cinvcode = drgroupby[j]["存货编码"].ToString();
                        sqlcmd.CommandText = "SELECT b.cInvCode FROM dbo.PO_Pomain as a inner join dbo.PO_Podetails as b on a.POID = b.POID where a.cPOID = '" + v_groupby + "'  AND b.cInvCode = '" + v_cinvcode +"'" ;
                        apdata.SelectCommand = sqlcmd;
                        orderlines.Reset();
                        apdata.Fill(orderlines);
                        if (orderlines.Tables[0].Rows.Count==0)
                        {
                            v_exitflag = true;
                            break;
                        }
                    }
                    if (v_exitflag)
                    {
                        v_importfailurerows = v_importfailurerows + 1;
                        //回写错误信息
                        for (int j = 0; j < drgroupby.Count(); j++)
                        {
                            drgroupby[j]["是否导入"] = "N";
                            drgroupby[j]["错误信息"] = "采购订单中不存在需要导入的存货编码，请检查!";
                        }
                        //复制已导入的数据到返回数据表中
                        for (int k = 0; k < drgroupby.Count(); k++)
                        {
                            dsreturnreceiptnotes.Tables["ReceiptNotes"].ImportRow(drgroupby[k]);
                        }
                        continue;
                    }
                }

                if ((!string.IsNullOrEmpty(v_groupby)) && (drgroupby[0]["业务类型"].ToString() == "委外加工"))
                {
                    for (int j = 0; j < drgroupby.Count(); j++)
                    {
                        v_cinvcode = drgroupby[j]["存货编码"].ToString();
                        sqlcmd.CommandText = "SELECT b.cInvCode FROM dbo.OM_MOMain as a inner join dbo.OM_MODetails as b on  a.MOID = b.MOID where a.cCode = '" + v_groupby + "'  AND b.cInvCode = '" + v_cinvcode + "'";
                        apdata.SelectCommand = sqlcmd;
                        orderlines.Reset();
                        apdata.Fill(orderlines);
                        if (orderlines.Tables[0].Rows.Count == 0)
                        {
                            v_exitflag = true;
                            break;
                        }
                    }
                    if (v_exitflag)
                    {
                        v_importfailurerows = v_importfailurerows + 1;
                        //回写错误信息
                        for (int j = 0; j < drgroupby.Count(); j++)
                        {
                            drgroupby[j]["是否导入"] = "N";
                            drgroupby[j]["错误信息"] = "委外订单中不存在需要导入的存货编码，请检查!";
                        }
                        //复制已导入的数据到返回数据表中
                        for (int k = 0; k < drgroupby.Count(); k++)
                        {
                            dsreturnreceiptnotes.Tables["ReceiptNotes"].ImportRow(drgroupby[k]);
                        }
                        continue;
                    }
                }
                #endregion
                //执行API单据赋值及参数赋值.
                #region
                if ((!string.IsNullOrEmpty(v_groupby))&&(drgroupby[0]["业务类型"].ToString() == "普通采购"))
                {
                    //获取采购订单头及订单行信息
                    sqlcmd.CommandText = "SELECT * FROM dbo.PO_Pomain WHERE cPOID ='" + v_groupby + "'";
                    apdata.SelectCommand = sqlcmd;
                    orderhead.Reset();
                    apdata.Fill(orderhead);
                    sqlcmd.CommandText = "SELECT * FROM dbo.PO_Podetails WHERE POID =" + orderhead.Tables[0].Rows[0]["POID"].ToString();
                    apdata.SelectCommand = sqlcmd;
                    orderlines.Reset();
                    apdata.Fill(orderlines);
                }

                if ((!string.IsNullOrEmpty(v_groupby)) && (drgroupby[0]["业务类型"].ToString() == "委外加工"))
                {
                    //获取采购订单头及订单行信息
                    sqlcmd.CommandText = "SELECT * FROM dbo.OM_MOMain WHERE cCode ='" + v_groupby + "'";
                    apdata.SelectCommand = sqlcmd;
                    orderhead.Reset();
                    apdata.Fill(orderhead);
                    sqlcmd.CommandText = "SELECT * FROM dbo.OM_MODetails  WHERE MOID =" + orderhead.Tables[0].Rows[0]["MOID"].ToString();
                    apdata.SelectCommand = sqlcmd;
                    orderlines.Reset();
                    apdata.Fill(orderlines);
                }
                //设置BO对象(表头)行数，只能为一行
                BusinessObject DomHead = broker.GetBoParam("DomHead");
                DomHead.RowCount = 1;
                sqlcmd.CommandText = "SELECT MAX(ID)+1 FROM dbo.RdRecord01 ";
                apdata.SelectCommand = sqlcmd;
                dssql.Reset();
                apdata.Fill(dssql);
                dtsql.Reset();
                dtsql = dssql.Tables[0];
                DomHead[0]["id"] = dtsql.Rows[0][0].ToString();                 //入库单主表主关键ID 

                sqlcmd.CommandText = "SELECT RIGHT('0000000000' + CONVERT(VARCHAR(10), max(ccode) + 1),10) FROM dbo.RdRecord01 ";
                apdata.SelectCommand = sqlcmd;
                dssql.Reset();
                apdata.Fill(dssql);
                dtsql.Reset();
                dtsql = dssql.Tables[0];
                v_receiptnotnumber = dtsql.Rows[0][0].ToString();  
                DomHead[0]["ccode"] = v_receiptnotnumber;                       //入库单编号
                DomHead[0]["ddate"] = drgroupby[0]["单据日期"].ToString();      //入库日期
                DomHead[0]["cmaker"] = u8userdata.UserId;                       //制单人

                if ((!string.IsNullOrEmpty(v_groupby)) && (drgroupby[0]["业务类型"].ToString() == "普通采购"))
                {
                    DomHead[0]["cvencode"]  = orderhead.Tables[0].Rows[0]["cVenCode"].ToString();                   //供应商编号
                    DomHead[0]["cdepcode"]  = orderhead.Tables[0].Rows[0]["cDepCode"].ToString();                   //部门编号
                    DomHead[0]["cbustype"]  = orderhead.Tables[0].Rows[0]["cBusType"].ToString();                   //业务类型
                    DomHead[0]["csource"]   = "采购订单";                                                           //单据来源                                               
                    DomHead[0]["iexchrate"] = orderhead.Tables[0].Rows[0]["nflat"].ToString();                      //汇率
                    DomHead[0]["cexch_name"] = orderhead.Tables[0].Rows[0]["cexch_name"].ToString();                //币种
                    DomHead[0]["ipurorderid"] = orderhead.Tables[0].Rows[0]["POID"].ToString();                     //采购订单ID
                }
                if ((!string.IsNullOrEmpty(v_groupby)) && (drgroupby[0]["业务类型"].ToString() == "委外加工"))
                {
                    DomHead[0]["cvencode"] = orderhead.Tables[0].Rows[0]["cVenCode"].ToString();                    
                    DomHead[0]["cdepcode"] = orderhead.Tables[0].Rows[0]["cDepCode"].ToString();
                    DomHead[0]["cbustype"] = orderhead.Tables[0].Rows[0]["cBusType"].ToString();
                    DomHead[0]["csource"] = "委外订单";
                    DomHead[0]["iexchrate"] = orderhead.Tables[0].Rows[0]["nflat"].ToString();
                    DomHead[0]["cexch_name"] = orderhead.Tables[0].Rows[0]["cexch_name"].ToString();
                    DomHead[0]["ipurorderid"] = orderhead.Tables[0].Rows[0]["MOID"].ToString();                     //委外订单ID
                }
                // 无订单号为参照则视作库存直接收货
                if ((string.IsNullOrEmpty(v_groupby)) && (drgroupby[0]["业务类型"].ToString() == "普通采购"))
                {
                    DomHead[0]["cvencode"] = drgroupby[0]["供应商编码"].ToString();
                    DomHead[0]["cdepcode"] = drgroupby[0]["部门编码"].ToString();
                    DomHead[0]["csource"] = drgroupby[0]["单据来源"].ToString();            //"库存"
                    DomHead[0]["cbustype"] = drgroupby[0]["业务类型"].ToString();           //"普通采购"
                    DomHead[0]["iexchrate"] = drgroupby[0]["汇率"].ToString();
                    DomHead[0]["cexch_name"] = drgroupby[0]["币种"].ToString();
                }
                
                DomHead[0]["cvouchtype"] = "01";                                                                    //单据类型这里固定是 01- 采购入库单
                DomHead[0]["cwhcode"] = drgroupby[0]["仓库编码"].ToString();                                        //仓库编码
                DomHead[0]["brdflag"] = "1";                                                                        //收发标志这里固定是收标志
                
                if (!string.IsNullOrEmpty(v_groupby))                                                               //采购类型及入库类别编码
                {
                    sqlcmd.CommandText = "SELECT a.cPTName,a.cRdCode,b.cRdName  FROM dbo.PurchaseType AS a inner join dbo.Rd_Style AS b on a.cRdCode = b.cRdCode WHERE a.cPTCode='" + orderhead.Tables[0].Rows[0]["cPTCode"].ToString() + "'";
                    apdata.SelectCommand = sqlcmd;
                    dssql.Reset();
                    apdata.Fill(dssql);
                    dtsql.Reset();
                    dtsql = dssql.Tables[0];
                    DomHead[0]["cptcode"] = orderhead.Tables[0].Rows[0]["cPTCode"].ToString();
                    DomHead[0]["crdcode"] = dtsql.Rows[0]["cRdCode"].ToString();
                    DomHead[0]["cordercode"] = v_groupby;                                                           //订单号
                    DomHead[0]["itaxrate"] = orderhead.Tables[0].Rows[0]["itaxrate"].ToString();                    //税率
                    
                }
                else
                {
                    DomHead[0]["cptcode"] = drgroupby[0]["采购类型编码"].ToString();
                    DomHead[0]["crdcode"] = drgroupby[0]["入库类别编码"].ToString();
                }
                //设置BO对象(表体)行数，只能为一行
                BusinessObject domBody = broker.GetBoParam("domBody");
                domBody.RowCount = 10;
                sqlcmd.CommandText = "SELECT MAX(autoid) FROM dbo.rdrecords01 ";
                apdata.SelectCommand = sqlcmd;
                dssql.Reset();
                apdata.Fill(dssql);
                dtsql.Reset();
                dtsql = dssql.Tables[0];
                Int32 v_linesidmax = Convert.ToInt32(dtsql.Rows[0][0]);
                for (int j = 0; j < drgroupby.Count(); j++)
                {
                    domBody[j]["autoid"] = v_linesidmax + 1;                                    //入库单子表主关键字段
                    domBody[j]["id"] = DomHead[0]["id"].ToString();                             //入库单主表主关键字段
                    domBody[j]["cinvcode"] = drgroupby[j]["存货编码"].ToString();               //存货编码
                    domBody[j]["iquantity"] = drgroupby[j]["入库数量"].ToString();              //入库数量
                    domBody[j]["editprop"] = "A";                                               //编辑属性：A表新增，M表修改，D表删除
                    domBody[j]["irowno"] = j + 1;                                               //行号
                    if ((!string.IsNullOrEmpty(v_groupby)) && (drgroupby[0]["业务类型"].ToString() == "普通采购"))
                    {
                        v_filterestr = " cInvCode = '" + drgroupby[j]["存货编码"].ToString() + "'";
                        drorderlines = orderlines.Tables[0].Select(v_filterestr);
                        domBody[j]["itaxrate"] = orderhead.Tables[0].Rows[0]["itaxrate"].ToString();        //税率
                        domBody[j]["ioritaxcost"] = drorderlines[0]["iTaxPrice"].ToString();                //原币含税单价
                        domBody[j]["ioricost"] = drorderlines[0]["iUnitPrice"].ToString();                  //原币单价
                        domBody[j]["iorimoney"] = Math.Round(Convert.ToDouble(drgroupby[j]["入库数量"]) * Convert.ToDouble(drorderlines[0]["iUnitPrice"]), 2).ToString();       //原币金额
                        domBody[j]["ioritaxprice"] = Math.Round(Convert.ToDouble(drgroupby[j]["入库数量"]) * Convert.ToDouble(drorderlines[0]["iUnitPrice"]) * Convert.ToDouble(orderhead.Tables[0].Rows[0]["itaxrate"]) / 100.00, 2).ToString();   //原币税额
                        domBody[j]["iorisum"] = (Convert.ToDouble(domBody[0]["iorimoney"]) + Convert.ToDouble(domBody[0]["ioritaxprice"])).ToString();      //原币价税合计
                        domBody[j]["iunitcost"] = drorderlines[0]["iNatUnitPrice"].ToString();              //本币无税单价
                        domBody[j]["iprice"] = Math.Round(Convert.ToDouble(drgroupby[j]["入库数量"]) * Convert.ToDouble(drorderlines[0]["iNatUnitPrice"]), 2).ToString();       //本币金额
                        domBody[j]["iaprice"] = Math.Round(Convert.ToDouble(drgroupby[j]["入库数量"]) * Convert.ToDouble(drorderlines[0]["iNatUnitPrice"]), 2).ToString();         //暂估金额
                        domBody[j]["facost"] = drorderlines[0]["iNatUnitPrice"].ToString();                 //暂估单价
                        domBody[j]["inquantity"] = drorderlines[0]["iQuantity"].ToString();                 //应收数量
                        domBody[j]["itaxprice"] = Math.Round(Convert.ToDouble(drgroupby[j]["入库数量"]) * Convert.ToDouble(drorderlines[0]["iNatUnitPrice"]) * Convert.ToDouble(orderhead.Tables[0].Rows[0]["itaxrate"]) / 100.00, 2).ToString(); ; //本币税额
                        domBody[j]["isum"] = (Convert.ToDouble(domBody[j]["iprice"]) + Convert.ToDouble(domBody[j]["itaxprice"])).ToString();  //本币价税合计
                        domBody[j]["cpoid"] = v_groupby;                                                  //订单号，string类型
                        domBody[j]["iposid"] = drorderlines[0]["id"].ToString();                           //订单子表ID
                        if (!string.IsNullOrEmpty(drgroupby[j]["入库件数"].ToString()))
                        {
                            domBody[0]["inum"] = drgroupby[j]["入库件数"].ToString();                                //入库件数
                            domBody[0]["iinvexchrate"] = Math.Round(Convert.ToDouble(drgroupby[j]["入库数量"]) / Convert.ToDouble(drgroupby[j]["入库件数"]), 4).ToString();  //换算率
                        }
                        domBody[0]["innum"] = "0.00";                                                           //应收件数
                        //获取入库物料信息：库存单位及是否批次控制
                        sqlcmd.CommandText = "SELECT bInvBatch,cSTComUnitCode FROM dbo.Inventory WHERE cInvCode = '" + drgroupby[j]["存货编码"].ToString() + "'";
                        apdata.SelectCommand = sqlcmd;
                        dssql.Reset();
                        apdata.Fill(dssql);
                        dtsql.Reset();
                        dtsql = dssql.Tables[0];
                        domBody[0]["cassunit"] = dtsql.Rows[0]["cSTComUnitCode"].ToString();                    //库存单位码，string类型
                        if (Convert.ToInt16(dtsql.Rows[0]["bInvBatch"]) == 1)
                        {
                            domBody[0]["cbatch"] = "001";                                                       //批号
                        }

                        if (!string.IsNullOrEmpty(drorderlines[0]["citemcode"].ToString()))
                        {
                            domBody[0]["citemcode"] = drorderlines[0]["citemcode"].ToString();                      //项目编码
                            domBody[0]["cname"] = drorderlines[0]["citemname"].ToString();                          //项目名称
                        }
                        if (!string.IsNullOrEmpty(drorderlines[0]["citem_class"].ToString()))
                        {
                            domBody[0]["citem_class"] = drorderlines[0]["citem_class"].ToString();                  //项目大类编码
                            //获取项目大类名称
                            sqlcmd.CommandText = "SELECT citem_name FROM fitem WHERE citem_class ='" + drorderlines[0]["citem_class"].ToString() + "'";
                            apdata.SelectCommand = sqlcmd;
                            dssql.Reset();
                            apdata.Fill(dssql);
                            dtsql.Reset();
                            dtsql = dssql.Tables[0];
                            domBody[0]["citemcname"] = dtsql.Rows[0]["citem_name"].ToString();                      //项目大类名称，string类型
                        }

                    }

                    if ((!string.IsNullOrEmpty(v_groupby)) && (drgroupby[0]["业务类型"].ToString() == "委外加工"))
                    {
                        v_filterestr = " cInvCode = '" + drgroupby[j]["存货编码"].ToString() + "'";
                        drorderlines = orderlines.Tables[0].Select(v_filterestr);
                        /*
                        
                        domBody[j]["ioritaxcost"] = drorderlines[0]["iTaxPrice"].ToString();                //原币含税单价
                        domBody[j]["ioricost"] = drorderlines[0]["iUnitPrice"].ToString();                  //原币单价
                        domBody[j]["iorimoney"] = Math.Round(Convert.ToDouble(drgroupby[j]["入库数量"]) * Convert.ToDouble(drorderlines[0]["iUnitPrice"]), 2).ToString();       //原币金额
                        domBody[j]["ioritaxprice"] = Math.Round(Convert.ToDouble(drgroupby[j]["入库数量"]) * Convert.ToDouble(drorderlines[0]["iUnitPrice"]) * Convert.ToDouble(orderhead.Tables[0].Rows[0]["itaxrate"]) / 100.00, 2).ToString();   //原币税额
                        domBody[j]["iorisum"] = (Convert.ToDouble(domBody[0]["iorimoney"]) + Convert.ToDouble(domBody[0]["ioritaxprice"])).ToString();      //原币价税合计
                        domBody[j]["iunitcost"] = drorderlines[0]["iNatUnitPrice"].ToString();              //本币无税单价
                        domBody[j]["iprice"] = Math.Round(Convert.ToDouble(drgroupby[j]["入库数量"]) * Convert.ToDouble(drorderlines[0]["iNatUnitPrice"]), 2).ToString();       //本币金额
                        domBody[j]["iaprice"] = Math.Round(Convert.ToDouble(drgroupby[j]["入库数量"]) * Convert.ToDouble(drorderlines[0]["iNatUnitPrice"]), 2).ToString();         //暂估金额
                        domBody[j]["facost"] = drorderlines[0]["iNatUnitPrice"].ToString();                 //暂估单价
                        domBody[j]["itaxprice"] = Math.Round(Convert.ToDouble(drgroupby[j]["入库数量"]) * Convert.ToDouble(drorderlines[0]["iNatUnitPrice"]) * Convert.ToDouble(orderhead.Tables[0].Rows[0]["itaxrate"]) / 100.00, 2).ToString(); ; //本币税额
                        domBody[j]["isum"] = (Convert.ToDouble(domBody[j]["iprice"]) + Convert.ToDouble(domBody[j]["itaxprice"])).ToString();  //本币价税合计
                        */
                        domBody[j]["itaxrate"] = 0.00;                                                      //税率 orderhead.Tables[0].Rows[0]["itaxrate"].ToString();        
                        domBody[j]["cpoid"] = v_groupby;                                                    //订单号，string类型
                        domBody[0]["iomodid"] = drorderlines[0]["MODetailsID"].ToString();                  //委外订单子表ID，int类型
                        domBody[j]["inquantity"] = drorderlines[0]["iQuantity"].ToString();                 //应收数量
                        domBody[0]["iprocesscost"] = drorderlines[0]["iUnitPrice"].ToString();              //加工费单价
                        domBody[0]["iprocessfee"] = Math.Round(Convert.ToDouble(drgroupby[j]["入库数量"]) * Convert.ToDouble(drorderlines[0]["iUnitPrice"]), 2).ToString();       //加工费


                        //获取入库物料信息：是否批次控制
                        sqlcmd.CommandText = "SELECT bInvBatch FROM dbo.Inventory WHERE cInvCode = '" + drgroupby[j]["存货编码"].ToString() + "'";
                        apdata.SelectCommand = sqlcmd;
                        dssql.Reset();
                        apdata.Fill(dssql);
                        dtsql.Reset();
                        dtsql = dssql.Tables[0];
                        if (Convert.ToInt16(dtsql.Rows[0]["bInvBatch"]) == 1)
                        {
                            domBody[0]["cbatch"] = "001";                                                       //批号
                        }

                        if (!string.IsNullOrEmpty(drorderlines[0]["citemcode"].ToString()))
                        {
                            domBody[0]["citemcode"] = drorderlines[0]["citemcode"].ToString();                      //项目编码
                            domBody[0]["cname"] = drorderlines[0]["citemname"].ToString();                          //项目名称
                        }
                        if (!string.IsNullOrEmpty(drorderlines[0]["citem_class"].ToString()))
                        {
                            domBody[0]["citem_class"] = drorderlines[0]["citem_class"].ToString();                  //项目大类编码
                            //获取项目大类名称
                            sqlcmd.CommandText = "SELECT citem_name FROM fitem WHERE citem_class ='" + drorderlines[0]["citem_class"].ToString() + "'";
                            apdata.SelectCommand = sqlcmd;
                            dssql.Reset();
                            apdata.Fill(dssql);
                            dtsql.Reset();
                            dtsql = dssql.Tables[0];
                            domBody[0]["citemcname"] = dtsql.Rows[0]["citem_name"].ToString();                      //项目大类名称，string类型
                        }
                    }



                }
                //API 通用参数赋值
                //给普通参数sVouchType赋值。此参数的数据类型为System.String，此参数按值传递，表示单据类型：01
                broker.AssignNormalValue("sVouchType", Convert.ToString("01"));
                //给普通参数domPosition赋值。此参数的数据类型为System.Object，此参数按引用传递，表示货位：传空
                broker.AssignNormalValue("domPosition", null); //broker.AssignNormalValue("domPosition", new System.Object());
                //该参数errMsg为OUT型参数，由于其数据类型为System.String，为值类型，因此不必传入参数变量。在API调用返回时，可以通过GetResult("errMsg")获取其值
                //给普通参数cnnFrom赋值。此参数的数据类型为ADODB.Connection，此参数按引用传递，表示连接对象,如果由调用方控制事务，则需要设置此连接对象，否则传空
                broker.AssignNormalValue("cnnFrom", null); //broker.AssignNormalValue("cnnFrom", new ADODB.Connection());
                //该参数VouchId为INOUT型普通参数。此参数的数据类型为System.String，此参数按值传递。在API调用返回时，可以通过GetResult("VouchId")获取其值
                broker.AssignNormalValue("VouchId", Convert.ToString(""));
                //该参数domMsg为OUT型参数，由于其数据类型为MSXML2.IXMLDOMDocument2，非一般值类型，因此必须传入一个参数变量。在API调用返回时，可以直接使用该参数.
                //无法直接创建接口实例，需要做类型转换 。//MSXML2.IXMLDOMDocument2 domMsg = new MSXML2.IXMLDOMDocument2();
                MSXML2.DOMDocumentClass domMsg = new MSXML2.DOMDocumentClass();
                broker.AssignNormalValue("domMsg", (IXMLDOMDocument2)domMsg);
                //给普通参数bCheck赋值。此参数的数据类型为System.Boolean，此参数按值传递，表示是否控制可用量。
                broker.AssignNormalValue("bCheck", false);
                //给普通参数bBeforCheckStock赋值。此参数的数据类型为System.Boolean，此参数按值传递，表示检查可用量
                broker.AssignNormalValue("bBeforCheckStock", false);
                //给普通参数bIsRedVouch赋值。此参数的数据类型为System.Boolean，此参数按值传递，表示是否红字单据
                broker.AssignNormalValue("bIsRedVouch", false);
                //给普通参数sAddedState赋值。此参数的数据类型为System.String，此参数按值传递，表示传空字符串
                broker.AssignNormalValue("sAddedState", Convert.ToString(""));
                //给普通参数bReMote赋值。此参数的数据类型为System.Boolean，此参数按值传递，表示是否远程：传入false
                broker.AssignNormalValue("bReMote", false);
                #endregion
                //第六步：调用API
                #region
                if (!broker.Invoke())
                {
                    //错误处理
                    Exception apiEx = broker.GetException();
                    if (apiEx != null)
                    {
                        if (apiEx is MomSysException)
                        {
                            MomSysException sysEx = apiEx as MomSysException;
                            v_errmsg = "系统异常：" + sysEx.Message + "\n\r";

                        }
                        else if (apiEx is MomBizException)
                        {
                            MomBizException bizEx = apiEx as MomBizException;
                            v_errmsg = "API异常：" + bizEx.Message + "\n\r";

                        }
                        //异常原因
                        String exReason = broker.GetExceptionString();
                        if (exReason.Length != 0)
                        {
                            v_errmsg = "其他异常原因：" + exReason + "\n\r";
                        }
                    }
                }
                #endregion
                //第七步：获取返回结果
                #region
                //获取普通返回值。此返回值数据类型为System.Boolean，此参数按值传递，表示返回值:true:成功,false:失败
                System.Boolean result = Convert.ToBoolean(broker.GetReturnValue());
                //获取out/inout参数值
                //获取普通OUT参数errMsg。此返回值数据类型为System.String，在使用该参数之前，请判断是否为空
                v_errmsg = (System.String)broker.GetResult("errMsg");
                //获取普通INOUT参数VouchId。此返回值数据类型为System.String，在使用该参数之前，请判断是否为空
                System.String v_vouchid = (System.String)broker.GetResult("VouchId");
                //获取普通OUT参数domMsg。此返回值数据类型为MSXML2.IXMLDOMDocument2，在使用该参数之前，请判断是否为空
                //MSXML2.IXMLDOMDocument2 domMsgRet = (MSXML2.DOMDocument)(broker.GetResult("domMsg"));
                //BusinessObject vdomBody = broker.GetBoParam("domBody");
                //BusinessObject vdomHead = broker.GetBoParam("DomHead");
                #endregion
                //第八步 ： 结束本次调用，释放API资源
                #region
                broker.Release();
                if (result)
                {
                    v_importsuccessrows = v_importsuccessrows + 1;
                    //回写信息
                    for (int j = 0; j < drgroupby.Count(); j++)
                    {
                        drgroupby[j]["是否导入"] = "Y";
                        drgroupby[j]["错误信息"] = "";
                        drgroupby[j]["单据号"] = v_receiptnotnumber;
                    }
                }
                else
                {
                    v_importfailurerows = v_importfailurerows + 1;
                    //回写错误信息
                    for (int j = 0; j < drgroupby.Count(); j++)
                    {
                        drgroupby[j]["是否导入"] = "N";
                        drgroupby[j]["错误信息"] = v_errmsg;
                    }
                }

                //复制已导入的数据到返回数据表中
                for (int k = 0; k < drgroupby.Count(); k++)
                {
                    dsreturnreceiptnotes.Tables["ReceiptNotes"].ImportRow(drgroupby[k]);
                }
                #endregion

                //执行进度条：PerformStep()函数
                importdataprogressBar.PerformStep();
                string str = Math.Round((100 * (i + 1.0) / dtdistinct.Rows.Count), 2).ToString("#0.00 ") + "%";
                Font font = new Font("Times New Roman", (float)10, FontStyle.Regular);
                PointF pt = new PointF(this.importdataprogressBar.Width / 2 - 17, this.importdataprogressBar.Height / 2 - 7);
                g.DrawString(str, font, Brushes.Blue, pt);

            }
            #endregion
            //结束本次数据导入调用,返回数据导入是否成功标志,
            Marshal.FinalReleaseComObject(u8Login);
            importsuccessrows = v_importsuccessrows;
            importfailurerows = v_importfailurerows;
            dsreturnreceiptnotes = dsimportedreceiptnotes;
            errmsg = v_errmsg;
            conn.Close();
            if (v_importfailurerows != 0)
                return false;
            else
                return true;

        }
    }
}
