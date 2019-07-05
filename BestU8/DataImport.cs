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
                importdataresulttextBox.AppendText("如果导入有出错，具体原因请看导入数据模板中错误信息列，请纠正后再次执行导入！\n" ) ;
                importdataresulttextBox.AppendText("数据导入执行结束:" + impend + "  \n");
            }
            #endregion

            //采购入库单导入
            #region
            if (Pubvar.gdataimporttype == "采购入库单导入")
            {
                string impstart, impend,v_errmsg;
                //采购入库单导入EXCEL中
                dtnpoidata = npoidata.ExcelToDataTable("ReceiptNotes", true, importdatafiletextBox.Text);
                dtnpoidata.TableName = "ReceiptNotes";
                dsexcel.Tables.Add(dtnpoidata);
                impstart = DateTime.Now.ToLocalTime().ToString();
                importdataresulttextBox.AppendText("数据导入执行开始:" + impstart + "\n");
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
                    filterestr = "凭证ID IS NULL  OR 凭证ID ='" + "'" ;
                }
                DataRow[] drgroupby = dsimportedvouchers.Tables["GLVouchers"].Select(filterestr);
                if ((!string.IsNullOrEmpty(drgroupby[0]["凭证号"].ToString())) || ((!string.IsNullOrEmpty(drgroupby[0]["是否导入"].ToString())) && (drgroupby[0]["是否导入"].ToString() == "N"))|| (string.IsNullOrEmpty(drgroupby[0]["凭证ID"].ToString())))
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

        public bool ReceiptNoteimport(UFSoft.U8.Framework.LoginContext.UserData u8userdata, DataSet dsimportedreceiptnotes, out int importsuccessrows, out int importfailurerows, out DataSet dsreturnreceiptnotes, out string errmsg)
        {
            int v_importsuccessrows = 0, v_importfailurerows = 0;
            string v_errmsg = "",dbname = "";
            SqlConnection conn = new SqlConnection(); 
            SqlDataAdapter apdata = new SqlDataAdapter();
            SqlCommand sqlcmd = new SqlCommand();
            DataTable dbaccinfo = new DataTable();
            DataSet poheadds = new DataSet(), polinesds = new DataSet();
            int pos = u8userdata.ConnString.IndexOf(";");
            conn.ConnectionString = u8userdata.ConnString.Remove(0, pos + 1);
            conn.Open();//连接数据库  
            sqlcmd.Connection = conn;
            /*
            //获取账套数据库名称
            sqlcmd.CommandText = "SELECT iYear FROM UFSystem.dbo.UA_Account WHWERE caccid = '" + u8userdata.AccID + "'";
            apdata.SelectCommand = sqlcmd;
            apdata.Fill(dbaccinfo);
            dbname = "UFDATA_" + u8userdata.AccID + "_" + dbaccinfo.Rows[0].ToString()+".";
            */
            //第一步：构造u8login对象并登陆(引用U8API类库中的Interop.U8Login.dll),如果当前环境中有login对象则可以省去第一步
            interU8lg::U8Login.clsLogin u8Login = new interU8lg::U8Login.clsLogin();
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

            //第二步：构造环境上下文对象，传入login，并按需设置其它上下文参数
            U8EnvContext envContext = new U8EnvContext();
            envContext.U8Login = u8Login;

            //第三步：设置API地址标识(Url)：当前API：添加新单据的地址标识为：U8API/PuStoreIn/Add
            U8ApiAddress BestU8ApiAddress = new U8ApiAddress("U8API/PuStoreIn/Add");

            //第四步：构造APIBroker
            U8ApiBroker broker = new U8ApiBroker(BestU8ApiAddress, envContext);

            //第五步：API单据值及参数赋值： 根据dataset中导入数据分组循环导入U8系统，如有采购订单则以采购订单作为分组条件，否则则以分组标识作为分组条件。

            dsreturnreceiptnotes = dsimportedreceiptnotes.Clone();
            DataTable dtdistinct = dsimportedreceiptnotes.Tables["ReceiptNotes"].DefaultView.ToTable(true, new string[] { "单据ID" });
            string vougroupby = "";
            //设置progressbar步长并显示百分比
            importdataprogressBar.Minimum = 0;   // 设置进度条最小值.
            importdataprogressBar.Value = 1;    // 设置进度条初始值
            importdataprogressBar.Step = 1;     // 设置每次增加的步长
            importdataprogressBar.Maximum = dtdistinct.Rows.Count;// 设置进度条最大值.
            Graphics g = this.importdataprogressBar.CreateGraphics();
            for (int i = 0; i < dtdistinct.Rows.Count; i++)   //分组开始
            {
                vougroupby = dtdistinct.Rows[i]["单据ID"].ToString();
                string filterestr = "";
                //当凭证ID为空或""时的特殊处理
                if (!string.IsNullOrEmpty(vougroupby))
                {
                    filterestr = "单据ID = " + "'" + vougroupby + "'";
                }
                else
                {
                    filterestr = "单据ID IS NULL  OR 单据ID ='" + "'";
                }
                DataRow[] drgroupby = dsimportedreceiptnotes.Tables["ReceiptNotes"].Select(filterestr);
                if ((!string.IsNullOrEmpty(drgroupby[0]["单据号"].ToString())) || ((!string.IsNullOrEmpty(drgroupby[0]["是否导入"].ToString())) && (drgroupby[0]["是否导入"].ToString() == "N")) || (string.IsNullOrEmpty(drgroupby[0]["单据ID"].ToString())))
                {
                    //复制已导入的数据到返回数据表中
                    for (int k = 0; k < drgroupby.Count(); k++)
                    {
                        dsreturnreceiptnotes.Tables["ReceiptNotes"].ImportRow(drgroupby[k]);
                    }
                }
                else
                {   //这组数据中需要保证订单号唯一
                    string v_ponumber = drgroupby[0]["订单号"].ToString();
                    bool podif = false;
                    for (int j = 0; j < drgroupby.Count(); j++)
                    {
                        if (v_ponumber != drgroupby[j]["订单号"].ToString())
                        {
                            podif = true;
                            break;
                        }
                    }
                    if (podif)
                    {
                        v_importfailurerows = v_importfailurerows + 1;
                        //回写错误信息
                        for (int j = 0; j < drgroupby.Count(); j++)
                        {
                            drgroupby[j]["是否导入"] = "N";
                            drgroupby[j]["错误信息"] = "一张采购入库单中存在不同订单号，请检查!";
                        }
                        //复制已导入的数据到返回数据表中
                        for (int k = 0; k < drgroupby.Count(); k++)
                        {
                            dsreturnreceiptnotes.Tables["ReceiptNotes"].ImportRow(drgroupby[k]);
                        }
                    }
                    else
                    {

                        if (!string.IsNullOrEmpty(v_ponumber))  
                        {
                            //获取订单头及订单行信息
                            sqlcmd.CommandText = "SELECT * FROM dbo.PO_Pomain WHERE cPOID ='" + v_ponumber + "'";
                            apdata.SelectCommand = sqlcmd;
                            apdata.Fill(poheadds);
                            sqlcmd.CommandText = "SELECT * FROM dbo.PO_Podetails WHERE POID =" + poheadds.Tables[0].Rows[0]["POID"].ToString() ;
                            apdata.SelectCommand = sqlcmd;
                            apdata.Fill(polinesds);
                        }

                        //API单据值赋值
                        #region 
                        //设置BO对象(表头)行数，只能为一行
                        BusinessObject DomHead = broker.GetBoParam("DomHead");
                        DataTable rdmainid = new DataTable(), rdmaincode = new DataTable(), rdlineid = new DataTable();

                        DomHead.RowCount = 1;

                        sqlcmd.CommandText = "SELECT MAX(ID)+1 FROM dbo.RdRecord01 ";
                        apdata.SelectCommand = sqlcmd;
                        apdata.Fill(rdmainid);
                        DomHead[0]["id"] = rdmainid.Rows[0][0].ToString();                  //"1000000404"; //主关键字段，int类型

                        sqlcmd.CommandText = "SELECT RIGHT('0000000000' + CONVERT(VARCHAR(10), max(ccode) + 1),10) FROM dbo.RdRecord01 ";
                        apdata.SelectCommand = sqlcmd;
                        apdata.Fill(rdmaincode);
                        DomHead[0]["ccode"] = rdmaincode.Rows[0][0].ToString();               //"testimp0006"; //入库单号，string类型

                        DomHead[0]["ddate"] = drgroupby[0]["单据日期"].ToString();      //"2015-01-12"; //入库日期，DateTime类型
                        DomHead[0]["cbustype"]=drgroupby[0]["业务类型"].ToString();     //"普通采购"; //业务类型，int类型
                        DomHead[0]["csource"] = drgroupby[0]["单据来源"].ToString();    //"库存"; //单据来源，int类型
                        DomHead[0]["cmaker"] = u8userdata.UserId ;                      //制单人，string类型      
                        DomHead[0]["iexchrate"] = drgroupby[0]["汇率"].ToString();      //汇率，double类型
                        DomHead[0]["cexch_name"] = drgroupby[0]["币种"].ToString();     // "人民币"; //币种，string类型
                        DomHead[0]["cvencode"] = drgroupby[0]["供应商编码"].ToString();       //"01002"; //供货单位编码，string类型
                        //这里固定是 01- 采购入库单 
                        DomHead[0]["cvouchtype"] = "01";                                      //单据类型，string类型 
                        DomHead[0]["cwhcode"] = drgroupby[0]["仓库编码"].ToString();        //"04";仓库编码，string类型
                        //这里固定是收标志
                        DomHead[0]["brdflag"] = "1";                                        //收发标志，int类型
                        DomHead[0]["crdcode"] = drgroupby[0]["入库类别编码"].ToString();        //入库类别编码，string类型  采购入库
                        DomHead[0]["cdepcode"] = drgroupby[0]["部门编码"].ToString();          // "0401"; //部门编码，string类型  采购部
                        if (!string.IsNullOrEmpty(v_ponumber))
                        {
                            DomHead[0]["cordercode"] = v_ponumber;                                       //订单号，string类型
                            DomHead[0]["itaxrate"] = poheadds.Tables[0].Rows[0]["itaxrate"].ToString();
                            DomHead[0]["ipurorderid"] = poheadds.Tables[0].Rows[0]["POID"].ToString();   //采购订单ID，string类型

                        }
                        else
                        {
                            DomHead[0]["cordercode"] = "";
                        }
                        //设置BO对象(表体）
                        BusinessObject domBody = broker.GetBoParam("domBody");
                        domBody.RowCount = drgroupby.Count();
                        sqlcmd.CommandText = "SELECT MAX(autoid) FROM dbo.rdrecords01 ";
                        apdata.SelectCommand = sqlcmd;
                        apdata.Fill(rdlineid);
                        Int32 v_linesidmax = Convert.ToInt32(rdlineid.Rows[0][0]);
                        string filter = "";
                        DataRow[] drpolines;
                        bool itemexist = true;
                        for (int j = 0; j < drgroupby.Count(); j++)
                        {
                            domBody[j]["autoid"] = v_linesidmax +1;                                  //"1000001229";   //主关键字段，int类型
                            domBody[j]["id"] = DomHead[0]["id"].ToString();                          //"1000000404"; //与收发记录主表关联项，int类型
                            domBody[j]["cinvcode"] = drgroupby[j]["存货编码"].ToString();            //"01019002082"; //存货编码，string类型
                            domBody[j]["iquantity"] = drgroupby[j]["入库数量"].ToString();           // "777.00"; //数量，double类型
                            domBody[j]["editprop"] = "A";                                            //编辑属性：A表新增，M表修改，D表删除，string类型
                            domBody[j]["irowno"] = j+1 ;                                             //行号，string类型

                            filter = "cInvCode = '" + drgroupby[j]["存货编码"].ToString()+"'";
                            if (!string.IsNullOrEmpty(v_ponumber))
                            {
                                drpolines = polinesds.Tables[0].Select(filter);
                                if (drpolines.Length ==0 )
                                {
                                    itemexist = false;
                                    break;
                                }
                                else
                                {
                                    //获取采购订单中价格及金额信息
                                    domBody[j]["itaxrate"] = poheadds.Tables[0].Rows[0]["itaxrate"].ToString(); //税率，double类型

                                    domBody[j]["ioritaxcost"] = drpolines[0]["iTaxPrice"].ToString();  //原币含税单价，double类型
                                    domBody[j]["ioricost"] = drpolines[0]["iUnitPrice"].ToString();     //原币单价，double类型
                                    domBody[j]["iorimoney"] = Math.Round(Convert.ToDouble(drgroupby[j]["入库数量"])* Convert.ToDouble(drpolines[0]["iUnitPrice"]),2).ToString();    //原币金额，double类型
                                    domBody[j]["ioritaxprice"] = Math.Round(Convert.ToDouble(drgroupby[j]["入库数量"]) * Convert.ToDouble(drpolines[0]["iUnitPrice"]) * Convert.ToDouble(poheadds.Tables[0].Rows[0]["itaxrate"]),2).ToString(); //原币税额，double类型
                                    domBody[j]["iorisum"] = (Convert.ToDouble(domBody[0]["iorimoney"]) + Convert.ToDouble(domBody[0]["ioritaxprice"])).ToString(); //原币价税合计，double类型

                                    domBody[j]["iunitcost"] = drpolines[0]["iNatUnitPrice"].ToString(); //本币无税单价 ，double类型
                                    domBody[j]["iprice"] = Math.Round(Convert.ToDouble(drgroupby[j]["入库数量"]) * Convert.ToDouble(drpolines[0]["iNatUnitPrice"]), 2).ToString(); //本币金额，double类型
                                    domBody[j]["iaprice"] = Math.Round(Convert.ToDouble(drgroupby[j]["入库数量"]) * Convert.ToDouble(drpolines[0]["iNatUnitPrice"]), 2).ToString();  //暂估金额，double类型
                                    //domBody[j]["cbatch"] = "001"; //批号，string类型
                                    domBody[j]["iposid"] = drpolines[0]["id"].ToString(); //订单子表ID，int类型
                                    domBody[j]["facost"] = drpolines[0]["iNatUnitPrice"].ToString(); //暂估单价，double类型
                                    domBody[j]["inquantity"] = drpolines[0]["iQuantity"]; //应收数量，double类型
                                    

                                    domBody[j]["itaxprice"] = Math.Round(Convert.ToDouble(drgroupby[j]["入库数量"]) * Convert.ToDouble(drpolines[0]["iNatUnitPrice"]) * Convert.ToDouble(poheadds.Tables[0].Rows[0]["itaxrate"]), 2).ToString(); ; //本币税额，double类型
                                    domBody[j]["isum"] = (Convert.ToDouble(domBody[j]["iprice"]) + Convert.ToDouble(domBody[j]["itaxprice"])).ToString();  //本币价税合计，double类型
                                    domBody[j]["cpoid"] = v_ponumber; //订单号，string类型

                                }
                            }

                        }
                        #endregion
                        //导入数据中item不存在PO中
                        if (!itemexist)
                        {
                            v_importfailurerows = v_importfailurerows + 1;
                            //回写错误信息
                            for (int j = 0; j < drgroupby.Count(); j++)
                            {
                                drgroupby[j]["是否导入"] = "N";
                                drgroupby[j]["错误信息"] = "采购入库单中入库商品在采购订单中不存在，请检查!";
                            }
                            //复制已导入的数据到返回数据表中
                            for (int k = 0; k < drgroupby.Count(); k++)
                            {
                                dsreturnreceiptnotes.Tables["ReceiptNotes"].ImportRow(drgroupby[k]);
                            }
                        }
                        else
                        {
                            
                            //API 参数赋值
                            #region
                            //给普通参数sVouchType赋值。此参数的数据类型为System.String，此参数按值传递，表示单据类型：01 --采购入库单
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
                                //结束本次调用，释放API资源
                                broker.Release();
                            }
                            #endregion
                            //第七步：获取返回结果
                            #region
                            //获取普通返回值。此返回值数据类型为System.Boolean，此参数按值传递，表示返回值:true:成功,false:失败
                            System.Boolean result = Convert.ToBoolean(broker.GetReturnValue());
                            //获取out/inout参数值
                            //获取普通OUT参数errMsg。此返回值数据类型为System.String，在使用该参数之前，请判断是否为空
                            errmsg = (System.String)broker.GetResult("errMsg");
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
                                    drgroupby[j]["单据号"] = v_vouchid;
                                }
                            }
                            else
                            {
                                v_importfailurerows = v_importfailurerows + 1;
                                //回写错误信息
                                for (int j = 0; j < drgroupby.Count(); j++)
                                {
                                    drgroupby[j]["是否导入"] = "N";
                                    drgroupby[j]["错误信息"] = errmsg;
                                }
                            }

                            //复制已导入的数据到返回数据表中
                            for (int k = 0; k < drgroupby.Count(); k++)
                            {
                                dsreturnreceiptnotes.Tables["ReceiptNotes"].ImportRow(drgroupby[k]);
                            }
                            #endregion
                        }

                    }
                }

                //执行进度条：PerformStep()函数
                importdataprogressBar.PerformStep();
                string str = Math.Round((100 * (i + 1.0) / dtdistinct.Rows.Count), 2).ToString("#0.00 ") + "%";
                Font font = new Font("Times New Roman", (float)10, FontStyle.Regular);
                PointF pt = new PointF(this.importdataprogressBar.Width / 2 - 17, this.importdataprogressBar.Height / 2 - 7);
                g.DrawString(str, font, Brushes.Blue, pt);

            } //分组结束

            //返回数据导入是否成功标志
            importsuccessrows = v_importsuccessrows;
            importfailurerows = v_importfailurerows;
            dsreturnreceiptnotes = dsimportedreceiptnotes;
            errmsg = v_errmsg;
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

    }
}
