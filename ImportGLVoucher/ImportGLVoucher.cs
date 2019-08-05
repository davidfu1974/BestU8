using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using UFIDA.U8.Portal.Proxy.editors;
using UFIDA.U8.Portal.Common.Core;
using UFIDA.U8.Portal.Framework.MainFrames;
using UFSoft.U8.Framework.Login.UI;
using UFIDA.U8.Portal.Framework.Actions;
using UFIDA.U8.Portal.Proxy.Actions;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace ImportGLVoucher
{
    public partial class ImportGLVoucher : UserControl
    {
        [DllImport("kernel32.dll")]
        public static extern IntPtr _lopen(string lpPathName, int iReadWrite);

        [DllImport("kernel32.dll")]
        public static extern bool CloseHandle(IntPtr hObject);

        public const int OF_READWRITE = 2;
        public const int OF_SHARE_DENY_NONE = 0x40;
        public readonly IntPtr HFILE_ERROR = new IntPtr(-1);


        public ImportGLVoucher()
        {
            InitializeComponent();
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

        private void importdatabutton_Click(object sender, EventArgs e)
        {
            DataSet dsexcel = new DataSet();
            DataSet dstoexcel = new DataSet();
            int importsuccessrows = 0, importfailurerows = 0;
            ExcelHelper npoidata = new ExcelHelper();
            System.Data.DataTable dtnpoidata = new System.Data.DataTable();

            //为防止用户多次点击导入按钮，将按钮禁用
            importdatabutton.Enabled = false;
            //判断数据模板EXCEL是否被用户打开
            IntPtr vHandle = _lopen(importdatafiletextBox.Text, OF_READWRITE | OF_SHARE_DENY_NONE);
            if (vHandle == HFILE_ERROR)
            {
                MessageBox.Show("请先关闭数据模板导入Excel文件！");
                importdatabutton.Enabled = true;
                return;
            }
            CloseHandle(vHandle);

            //总账凭证导入
            string impstart, impend, v_errmsg;
            dtnpoidata = npoidata.ReadExcelToDatatble("GLVouchers", importdatafiletextBox.Text, 1, 1);
            dtnpoidata.TableName = "GLVouchers";
            dsexcel.Tables.Add(dtnpoidata);
            impstart = DateTime.Now.ToLocalTime().ToString();
            importdataresulttextBox.AppendText("数据导入执行开始......:   " + impstart + "  \n");
            importdataresulttextBox.AppendText("\n");
            importdataresulttextBox.Refresh();
            //ExcelHelper.wl.WriteLogs("Debug ......");

            //调用总账导入功能
            try
            {
                bool v_importglvouchersflag = GLvouchersimport(Pubvar.userToken, Pubvar.ConnString, dsexcel, Pubvar.UserId, out importsuccessrows, out importfailurerows, out dstoexcel, out v_errmsg);
                //导入结果回写EXCEL 
                npoidata.WriteDataTableToUpdateExcel(dstoexcel.Tables["GLVouchers"], "GLVouchers", importdatafiletextBox.Text);
                //执行结果回写memo text
                impend = DateTime.Now.ToLocalTime().ToString();
                importdataresulttextBox.AppendText("此次数据导入共计执行：" + (importsuccessrows + importfailurerows) + " 条 \n");
                importdataresulttextBox.AppendText("\n");
                importdataresulttextBox.AppendText("其中导入成功：" + importsuccessrows + " 条 , 导入失败： " + importfailurerows + " 条 \n");
                importdataresulttextBox.AppendText("\n");
                importdataresulttextBox.AppendText("如果导入有出错，具体原因请看导入数据模板中错误信息列，请纠正后再次执行导入！\n");
                importdataresulttextBox.AppendText("\n");
                if ((!v_importglvouchersflag) && (!string.IsNullOrEmpty(v_errmsg)))
                {
                    importdataresulttextBox.AppendText("系统调用出错：" + v_errmsg + "  \n");
                    importdataresulttextBox.AppendText(" \n");
                }
                importdataresulttextBox.AppendText("数据导入执行结束......:   " + impend + "  \n");
                importdataresulttextBox.AppendText(" \n");
                importdataresulttextBox.Refresh();
                importdatabutton.Enabled = true;
            }
            catch (Exception ex)
            {

                //ExcelHelper.wl.WriteLogs(ex.Message);
            }

        }


        public bool GLvouchersimport(string usertoken, string dbconn, DataSet dsimportedvouchers, string userid, out int importsuccessrows, out int importfailurerows, out DataSet dsreturnvouchers, out string errmsg)
        {
            string strSql = "", strTempTable = "tempdb.dbo.bestu8cus_gl_accvouchers";
            int v_importsuccessrows = 0, v_importfailurerows = 0;
            System.Object rsaffected = new System.Object();
            //创建或清除凭证导入临时表数据
            #region
            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Connection conn = new ADODB.Connection();
            conn.Open(dbconn);
            strSql = "SELECT count(*) FROM tempdb.dbo.sysobjects WHERE name = 'bestu8cus_gl_accvouchers'";
            rs = conn.Execute(strSql, out rsaffected, -1);
            if (Convert.ToInt16(rs.Fields[0].Value) > 0)
            {
                strSql = "DELETE FROM " + strTempTable;
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
            //调用API保存总账凭证
            CVoucher.CVInterface glcvoucher = new CVoucher.CVInterface();
            glcvoucher.set_Connection(conn);
            glcvoucher.strTempTable = strTempTable;
            glcvoucher.LoginByUserToken(usertoken);
            //根据dataset中导入数据分组循环导入U8系统
            dsreturnvouchers = dsimportedvouchers.Clone();
            //添加对账套的校验逻辑
            System.Data.DataTable dtdistinctaccid= dsimportedvouchers.Tables["GLVouchers"].DefaultView.ToTable(true, new string[] { "账套" });
            if (dtdistinctaccid.Rows.Count>1)
            {
                //返回数据导入是否成功标志
                importsuccessrows = 0;
                importfailurerows = 0;
                conn.Close();
                errmsg = "存在账套不唯一，请核查模板账套列数据";
                return false;

            }
            else
            {
                if (dtdistinctaccid.Rows[0]["账套"].ToString() != Pubvar.accid)
                {
                    //返回数据导入是否成功标志
                    importsuccessrows = 0;
                    importfailurerows = 0;
                    conn.Close();
                    errmsg = "导入模板账套("+ dtdistinctaccid.Rows[0]["账套"].ToString() + ")与用户登陆U8账套(" + Pubvar.accid + ") 不一致，请核查！";
                    return false;

                }
            }
            System.Data.DataTable dtdistinct = dsimportedvouchers.Tables["GLVouchers"].DefaultView.ToTable(true, new string[] { "凭证ID" });
            string vougroupby = "";
            //设置progressbar步长并显示百分比
            importdataprogressBar.Minimum = 0;   // 设置进度条最小值.
            importdataprogressBar.Value = 1;    // 设置进度条初始值
            importdataprogressBar.Step = 1;     // 设置每次增加的步长
            importdataprogressBar.Maximum = dtdistinct.Rows.Count;// 设置进度条最大值.
            //Graphics g = this.importdataprogressBar.CreateGraphics();

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
                        strSql = "INSERT INTO " + strTempTable + "(ioutperiod,coutsign ,cSign,coutno_id,cdigest,ctext1,coutsysname,cbill,inid,ccode,cexch_name ,doutbilldate,bvouchedit,bvouchaddordele,bvouchmoneyhold,bvalueedit,bcodeedit,md_f,mc_f,md,mc,nfrat,cdept_id,cperson_id,ccus_id,csup_id,citem_class,citem_id,cDefine12,cDefine13) ";
                        strSql = strSql + "VALUES(" + drgroupby[j]["会计期间"].ToString();
                        strSql = strSql + ",'" + drgroupby[j]["凭证类别"].ToString();
                        strSql = strSql + "','" + drgroupby[j]["凭证类别"].ToString();
                        strSql = strSql + "','" + drgroupby[j]["凭证ID"].ToString();
                        strSql = strSql + "','" + drgroupby[j]["摘要"].ToString();
                        strSql = strSql + "','" + drgroupby[j]["原凭证号"].ToString();
                        
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
                        //外币借贷
                        if (string.IsNullOrEmpty(drgroupby[j]["借方原币金额"].ToString()))
                        {
                            strSql = strSql + "," + 0;   //md_f
                        }
                        else
                        {
                            strSql = strSql + "," + drgroupby[j]["借方原币金额"].ToString();   //md
                        }

                        if (string.IsNullOrEmpty(drgroupby[j]["贷方原币金额"].ToString()))
                        {
                            strSql = strSql + "," + 0;   //mc_f
                        }
                        else
                        {
                            strSql = strSql + "," + drgroupby[j]["贷方原币金额"].ToString();   //mc
                        }

                        //本位币借贷
                        if (string.IsNullOrEmpty(drgroupby[j]["借方本位币金额"].ToString()))
                        {
                            strSql = strSql + "," + 0;   //md
                        }
                        else
                        {
                            strSql = strSql + "," + drgroupby[j]["借方本位币金额"].ToString();   //md
                        }

                        if (string.IsNullOrEmpty(drgroupby[j]["贷方本位币金额"].ToString()))
                        {
                            strSql = strSql + "," + 0;   //mc
                        }
                        else
                        {
                            strSql = strSql + "," + drgroupby[j]["贷方本位币金额"].ToString();   //mc
                        }
                        //汇率
                        if (string.IsNullOrEmpty(drgroupby[j]["汇率"].ToString()))
                        {
                            strSql = strSql + "," + 0;   
                        }
                        else
                        {
                            strSql = strSql + "," + drgroupby[j]["汇率"].ToString();   
                        }
                        strSql = strSql + ",'" + drgroupby[j]["部门编码"].ToString();               //部门编码
                        strSql = strSql + "','" + drgroupby[j]["职员编码"].ToString();              //职员编码
                        strSql = strSql + "','" + drgroupby[j]["客户编码"].ToString();              //客户编码
                        strSql = strSql + "','" + drgroupby[j]["供应商编码"].ToString();            //供应商编码
                        strSql = strSql + "','" + drgroupby[j]["项目大类编码"].ToString();          //物料大类编码
                        strSql = strSql + "','" + drgroupby[j]["项目编码"].ToString();              //物料编码
                        strSql = strSql + "','" + drgroupby[j]["政府项目"].ToString();              //政府项目
                        strSql = strSql + "','" + drgroupby[j]["资金来源"].ToString() + "')";       //资金来源
                        rs = conn.Execute(strSql, out rsaffected, -1);
                    }
                    //凭证导入U8中制单
                    bool glsaveflag = glcvoucher.SaveVoucher();
                    //回写凭证号及错误信息,一旦SaveVoucher成功执行完毕，数据库连接系统API自动关闭，需要再次打开
                    if (glsaveflag)
                    {
                        v_importsuccessrows = v_importsuccessrows + 1;
                        int importedvoucherid;
                        strSql = "SELECT distinct ino_id  FROM " + strTempTable + " WHERE coutno_id ='" + vougroupby + "'";
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
                    strSql = "DELETE FROM " + strTempTable ;
                    rs = conn.Execute(strSql, out rsaffected, -1);

                    //复制已导入数据到返回数据表中
                    for (int k = 0; k < drgroupby.Count(); k++)
                    {
                        dsreturnvouchers.Tables["GLVouchers"].ImportRow(drgroupby[k]);
                    }
                }

                //执行PerformStep()函数
                importdataprogressBar.PerformStep();
                //string str = Math.Round((100 * (i + 1.0) / dtdistinct.Rows.Count), 2).ToString("#0.00 ") + "%";
                //System.Drawing.Font font = new System.Drawing.Font("Times New Roman", (float)10, FontStyle.Regular);
                //PointF pt = new PointF(this.importdataprogressBar.Width / 2 - 17, this.importdataprogressBar.Height / 2 - 7);
                //g.DrawString(str, font, Brushes.Yellow, pt);

            }

            //返回数据导入是否成功标志
            importsuccessrows = v_importsuccessrows;
            importfailurerows = v_importfailurerows;
            conn.Close();
            errmsg = "";
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

    //U8 portal NetLoginable 实现
    public class BestLoginable : UFIDA.U8.Portal.Proxy.supports.NetLoginable
    {
        public BestLoginable()
        {
        }

        public override object CallFunction(string cMenuId, string cMenuName, string cAuthId, string cCmdLine)
        {
            INetUserControl BestImportGLVouchersctl = new BestImportGLVouchersControl();
            base.ShowEmbedControl(BestImportGLVouchersctl, cMenuId, true);
            return null;
        }
        public override bool SubSysLogin()
        {
            GlobalParameters.gLoginable = this;
            return base.SubSysLogin();
        }

        public override bool SubSysLogOff()
        {
            return base.SubSysLogOff();
        }

    }

    public static class GlobalParameters
    {
        private static BestLoginable _myLoginable = null;

        public static BestLoginable gLoginable
        {
            get
            {
                return _myLoginable;
            }
            set
            {
                _myLoginable = value;
            }
        }
    }

    public class BestImportGLVouchersControl : UFIDA.U8.Portal.Proxy.editors.INetUserControl
    {
        ImportGLVoucher bestglimportusercontrol = null;
        private IEditorInput _editInput = null;
        private IEditorPart _editPart = null;
        private string _title;

        public IEditorInput EditorInput
        {
            get
            {
                return _editInput;
            }
            set
            {
                _editInput = value;
            }
        }

        public IEditorPart EditorPart
        {
            get
            {
                return _editPart;
            }
            set
            {
                _editPart = value;
            }
        }

        public string Title
        {
            get
            {
                return this._title;
            }
            set
            {
                this._title = value;
            }
        }

        public bool CloseEvent()
        {
            return true;
        }

        public System.Windows.Forms.Control CreateControl(clsLogin login, string MenuID, string Paramters)
        {
            Pubvar.ConnString = login.GetLoginInfo().ConnString.ToString();
            Pubvar.UserId = login.GetLoginInfo().UserName.ToString();
            Pubvar.userToken = login.userToken.ToString();
            Pubvar.accid = login.GetLoginInfo().AccID.ToString();

            //初始化自定义用户控件对象
            bestglimportusercontrol = new ImportGLVoucher();

            return bestglimportusercontrol;
        }

        public class NetSampleDelegate : IActionDelegate
        {
            public void Run(UFIDA.U8.Portal.Framework.Actions.IAction action)
            {
                string id = action.Id;
                if (id == "about")
                {
                    //MessageBox.Show("关闭按钮");
                }
            }

            public void SelectionChanged(UFIDA.U8.Portal.Framework.Actions.IAction action, ISelection selection)
            {

            }
        }

        public UFIDA.U8.Portal.Proxy.Actions.NetAction[] CreateToolbar(clsLogin login)
        {
            //IActionDelegate nsd = new NetSampleDelegate();
            //NetAction ac = new NetAction("about", nsd);
            //NetAction[] aclist;
            //aclist = new NetAction[1];
            //ac.Text = "关于";
            //aclist[0] = ac;
            NetAction[] aclist;
            aclist = new NetAction[1];
            aclist[0] = null;
            return aclist;
        }

    }

    public static class Pubvar
    {
        public static string ConnString;

        public static string UserId;

        public static string userToken;

        public static string accid;
    }

    public class ExcelHelper
    {
        public static WriteLog wl = new WriteLog();

        public System.Data.DataTable ReadExcelToDatatble(string worksheetName, string saveAsLocation, int HeaderLine, int ColumnStart)
        {
            System.Data.DataTable dataTable = new System.Data.DataTable();
            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook excelworkBook;
            Microsoft.Office.Interop.Excel.Worksheet excelSheet;
            Microsoft.Office.Interop.Excel.Range range;
            try
            {
                // Start Excel and get Application object.
                excel = new Microsoft.Office.Interop.Excel.Application();

                // for making Excel visible
                excel.Visible = false;
                excel.DisplayAlerts = false;

                // Creation a new Workbook
                excelworkBook = excel.Workbooks.Open(saveAsLocation);

                // Workk sheet
                excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.Worksheets.Item[worksheetName];
                range = excelSheet.UsedRange;

                int cl = range.Columns.Count;
                // loop through each row and add values to our sheet
                int rowcount = range.Rows.Count; ;

                for (int j = ColumnStart; j <= cl; j++)
                {
                    dataTable.Columns.Add(Convert.ToString(((Range)(range.Cells[HeaderLine, j])).Value2), typeof(string));
                }
                for (int i = HeaderLine + 1; i <= rowcount; i++)
                {
                    DataRow dr = dataTable.NewRow();
                    for (int j = ColumnStart; j <= cl; j++)
                    {
                        //判断是否为日期格式的单元格
                        string dateformat = ((Range)(range.Cells[i, j])).NumberFormat.ToString();
                        if (dateformat.IndexOf("yyyy") == -1)
                        {
                            dr[j - ColumnStart] = Convert.ToString(((Range)(range.Cells[i, j])).Value2);
                        }
                        else
                        {
                            dr[j - ColumnStart] = DateTime.FromOADate(Convert.ToDouble(((Range)(range.Cells[i, j])).Value2)).ToString("yyyy-MM-dd");
                        }
                    }
                    // on the first iteration we add the column headers
                    dataTable.Rows.InsertAt(dr, dataTable.Rows.Count + 1);
                }
                //now save the workbook and exit Excel
                excelworkBook.Close();
                excel.Quit();
                return dataTable;
            }
            catch (Exception ex)
            {
                return null;
            }
            finally
            {
                excelSheet = null;
                range = null;
                excelworkBook = null;
            }

        }

        public bool WriteDataTableToExcel(System.Data.DataTable dataTable, string worksheetName, string saveAsLocation)
        {
            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook excelworkBook;
            Microsoft.Office.Interop.Excel.Worksheet excelSheet;
            Microsoft.Office.Interop.Excel.Range excelCellrange;

            try
            {
                // Start Excel and get Application object.
                excel = new Microsoft.Office.Interop.Excel.Application();

                // for making Excel visible
                excel.Visible = false;
                excel.DisplayAlerts = false;

                // Creation a new Workbook
                excelworkBook = excel.Workbooks.Add(Type.Missing);

                // Workk sheet
                excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.ActiveSheet;
                excelSheet.Name = worksheetName;


                // loop through each row and add values to our sheet
                int rowcount = 1;

                foreach (DataRow datarow in dataTable.Rows)
                {
                    rowcount += 1;
                    for (int i = 1; i <= dataTable.Columns.Count; i++)
                    {
                        // on the first iteration we add the column headers
                        if (rowcount == 3)
                        {
                            excelSheet.Cells[2, i] = dataTable.Columns[i - 1].ColumnName;
                            excelSheet.Cells.Font.Color = System.Drawing.Color.Black;

                        }

                        excelSheet.Cells[rowcount, i] = datarow[i - 1].ToString();

                        //for alternate rows
                        if (rowcount > 2)
                        {
                            if (i == dataTable.Columns.Count)
                            {
                                if (rowcount % 2 == 0)
                                {
                                    excelCellrange = excelSheet.Range[excelSheet.Cells[rowcount, 1], excelSheet.Cells[rowcount, dataTable.Columns.Count]];
                                    FormattingExcelCells(excelCellrange, "#CCCCFF", System.Drawing.Color.Black, false);
                                }

                            }
                        }

                    }

                }

                // now we resize the columns
                excelCellrange = excelSheet.Range[excelSheet.Cells[1, 1], excelSheet.Cells[rowcount, dataTable.Columns.Count]];
                excelCellrange.EntireColumn.AutoFit();
                Microsoft.Office.Interop.Excel.Borders border = excelCellrange.Borders;
                border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border.Weight = 2d;


                excelCellrange = excelSheet.Range[excelSheet.Cells[1, 1], excelSheet.Cells[2, dataTable.Columns.Count]];
                FormattingExcelCells(excelCellrange, "#000099", System.Drawing.Color.White, true);


                //now save the workbook and exit Excel


                excelworkBook.SaveAs(saveAsLocation); ;
                excelworkBook.Close();
                excel.Quit();
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
            finally
            {
                excelSheet = null;
                excelCellrange = null;
                excelworkBook = null;
            }

        }

        public bool WriteDataTableToUpdateExcel(System.Data.DataTable dataTable, string worksheetName, string saveAsLocation)
        {
            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook excelworkBook;
            Microsoft.Office.Interop.Excel.Worksheet excelSheet;

            try
            {
                // Start Excel and get Application object.
                excel = new Microsoft.Office.Interop.Excel.Application();

                // for making Excel visible
                excel.Visible = false;
                excel.DisplayAlerts = false;

                // Creation a new Workbook
                excelworkBook = excel.Workbooks.Open(saveAsLocation);

                // Workk sheet
                excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.Worksheets.Item[worksheetName];
                // loop through each row and add values to our sheet,exclude head columns
                int rowcount = 2;
                foreach (DataRow datarow in dataTable.Rows)
                {
                    for (int i = 1; i <= dataTable.Columns.Count; i++)
                    {
                        ((Range)(excelSheet.Cells[rowcount, i])).Value2 = datarow[i - 1].ToString();
                    }
                    rowcount = rowcount + 1;
                }
                //now save the workbook and exit Excel
                excelworkBook.Save();
                excelworkBook.Close();
                excel.Quit();
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
            finally
            {
                excelSheet = null;
                excelworkBook = null;
            }

        }

        public void FormattingExcelCells(Microsoft.Office.Interop.Excel.Range range, string HTMLcolorCode, System.Drawing.Color fontColor, bool IsFontbool)
        {
            range.Interior.Color = System.Drawing.ColorTranslator.FromHtml(HTMLcolorCode);
            range.Font.Color = System.Drawing.ColorTranslator.ToOle(fontColor);
            if (IsFontbool == true)
            {
                range.Font.Bold = IsFontbool;
            }
        }

    }

    public class WriteLog
    {
        public void WriteLogs(string strmessage)
        {
            try
            {
                string text = AppDomain.CurrentDomain.SetupInformation.ApplicationBase.ToString() + "\\log";
                if (!Directory.Exists(text))
                {
                    Directory.CreateDirectory(text);
                }
                string text2 = text + "\\record.log";
                FileStream fileStream;
                if (File.Exists(text2))
                {
                    new FileInfo(text2);
                    fileStream = new FileStream(text2, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
                    fileStream.Seek(0L, SeekOrigin.End);
                }
                else
                {
                    fileStream = new FileStream(text2, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
                    fileStream.Seek(0L, SeekOrigin.End);
                }
                StreamWriter streamWriter = new StreamWriter(fileStream);
                streamWriter.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.") + DateTime.Now.Millisecond.ToString());
                streamWriter.Write(strmessage);
                streamWriter.WriteLine();
                streamWriter.Close();
                fileStream.Close();
            }
            catch (Exception)
            {
            }
        }

        public void WriteLogs(string strmessage, string pathname)
        {
            try
            {
                string text = AppDomain.CurrentDomain.SetupInformation.ApplicationBase.ToString() + "\\log";
                if (!Directory.Exists(text))
                {
                    Directory.CreateDirectory(text);
                }
                string text2 = text + "\\" + pathname + ".log";
                FileStream fileStream;
                if (File.Exists(text2))
                {
                    new FileInfo(text2);
                    fileStream = new FileStream(text2, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
                    fileStream.Seek(0L, SeekOrigin.End);
                }
                else
                {
                    fileStream = new FileStream(text2, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
                    fileStream.Seek(0L, SeekOrigin.End);
                }
                StreamWriter streamWriter = new StreamWriter(fileStream);
                streamWriter.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.") + DateTime.Now.Millisecond.ToString());
                streamWriter.Write(strmessage);
                streamWriter.WriteLine();
                streamWriter.Close();
                fileStream.Close();
            }
            catch (Exception)
            {
            }
        }

        public static string ReadLogs(string FilePath)
        {
            string text = "";
            if (File.Exists(FilePath))
            {
                StreamReader streamReader = File.OpenText(FilePath);
                for (string text2 = streamReader.ReadLine(); text2 != null; text2 = streamReader.ReadLine())
                {
                    text = text + text2 + "\r\n";
                }
                streamReader.Close();
            }
            return text;
        }
    }
}
