﻿extern alias interU8lg;
extern alias interadodb;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Data;
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
        public DataImport()
        {
            InitializeComponent();
            importtypeGB.Text = Pubvar.gdataimporttype;
            this.Text = "DataImport - " + Pubvar.gdataimporttype;
        }

        private void importdataopenfilebutton_Click(object sender, EventArgs e)
        {
            //OpenFileDialog openFileDialog = new OpenFileDialog();
            importdataopenFileDialog.InitialDirectory = "c:\\";
            importdataopenFileDialog.Filter = "Excel文件(*.xls)|*.xls|Excel文件(*.xlsx)|*.xlsx|所有文件(*.*)|*.*";
            importdataopenFileDialog.RestoreDirectory = true;
            importdataopenFileDialog.FilterIndex = 1;
            if (importdataopenFileDialog.ShowDialog() == DialogResult.OK)
            {
                importdatafiletextBox.Text = importdataopenFileDialog.FileName;
                //读取文件内容
                /*
                File fileOpen = new File(fName);
                isFileHaveName = true;
                richTextBox1.Text = fileOpen.ReadFile();
                richTextBox1.AppendText("");
                */
            }
        }

        private void closebutton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void importdatabutton_Click(object sender, EventArgs e)
        {
            //总账凭证导入
            if (Pubvar.gdataimporttype == "总账凭证导入")
            {
                GLvouchersimp();
            }

            //采购入库单导入
            if (Pubvar.gdataimporttype == "采购入库单导入")
            {
                ReceiptNoteimp();
            }
        }

        private void GLvouchersimp()
        {
            int v_importsuccessrows = 0, v_importfailurerows = 0;
            //根据总账导入EXCEL模板将数据导入到dataset 中
            DataSet v_vouchersfromexcel = new DataSet();
            DataSet v_returnvouchers = new DataSet();

            //调用总账导入功能
            bool v_importglvouchersflag = GLvouchersimport(Pubvar.gu8LoginUI.userToken, Pubvar.gu8userdata.ConnString, v_vouchersfromexcel, Pubvar.gu8userdata.UserId, out v_importsuccessrows, out v_importfailurerows, out v_returnvouchers);
            MessageBox.Show(v_importglvouchersflag.ToString());
        }

        private void ReceiptNoteimp()
        {
            int v_importsuccessrows = 0, v_importfailurerows = 0;
            String v_errmsg;
            //根据采购入库单导入EXCEL模板将数据导入到dataset 中
            DataSet v_receiptnotesfromexcel = new DataSet();
            DataSet v_returnreceiptnotes = new DataSet();

            //调用采购入库单导入功能
            bool v_importreceiptnoteflag = ReceiptNoteimport(Pubvar.gu8userdata, v_receiptnotesfromexcel, out v_importsuccessrows, out v_importfailurerows, out v_returnreceiptnotes, out v_errmsg);
            MessageBox.Show(v_importreceiptnoteflag.ToString());
        }

        public bool GLvouchersimport(string usertoken, string dbconn, DataSet dsimportedvouchers, string userid, out int importsuccessrows, out int importfailurerows, out DataSet dsreturnvouchers)
        {
            string strSql = "", strTempTable = "tempdb.dbo.cus_gl_accvouchers";
            int v_importsuccessrows = 0, v_importfailurerows = 0;
            System.Object rsaffected = new System.Object();
            //创建或清除凭证导入临时表数据
            interadodb::ADODB.Recordset rs = new interadodb::ADODB.Recordset();
            interadodb::ADODB.Connection conn = new interadodb::ADODB.Connection();
            conn.Open(dbconn);
            //创建临时表，如已创建删除表内记录
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

            //临时表中插入总账凭证数据

            //测试数据
            //借方
            strSql = "INSERT INTO tempdb.dbo.cus_gl_accvouchers(ioutperiod,coutsign ,cSign,coutno_id,cdigest,coutsysname,cbill,inid,ccode,cexch_name ,doutbilldate,bvouchedit,bvouchaddordele,bvouchmoneyhold,bvalueedit,bcodeedit,md) ";
            strSql = strSql + "VALUES(1, N'记', N'记', N'IMP0000001', N'测试后台导入总账凭证', N'GL', N'" + userid + "', 1, N'6402', N'人民币',  '2015-1-31', 1, 1, 1,1,1, 777)";
            rs = conn.Execute(strSql, out rsaffected, -1);
            //贷方
            strSql = "INSERT INTO tempdb.dbo.cus_gl_accvouchers(ioutperiod,coutsign ,cSign,coutno_id,cdigest,coutsysname,cbill,inid,ccode,cexch_name ,doutbilldate,bvouchedit,bvouchaddordele,bvouchmoneyhold,bvalueedit,bcodeedit,mc) ";
            strSql = strSql + "VALUES(1, N'记', N'记', N'IMP0000001', N'测试后台导入总账凭证', N'GL', N'" + userid + "', 1, N'6711', N'人民币',  '2015-1-31', 1, 1, 1,1,1, 777)";
            rs = conn.Execute(strSql, out rsaffected, -1);
            

            //调用API保存总账凭证
            CVoucher.CVInterface glcvoucher = new CVoucher.CVInterface();
            glcvoucher.set_Connection(conn);
            glcvoucher.strTempTable = strTempTable;
            glcvoucher.LoginByUserToken(usertoken);
            //根据dataset中导入数据分组循环导入U8系统


            if (glcvoucher.SaveVoucher())
            {
                v_importsuccessrows = v_importsuccessrows + 1;
            }
            else
            {
                v_importfailurerows = v_importfailurerows + 1;
            }


            //返回数据导入是否成功标志
            importsuccessrows = v_importsuccessrows;
            importfailurerows = v_importfailurerows;
            dsreturnvouchers = dsimportedvouchers;

            if (v_importfailurerows != 0)
            {
                return false;
            }
            else
            {
                return true;
            }

        }

        public bool ReceiptNoteimport(UFSoft.U8.Framework.LoginContext.UserData u8userdata, DataSet dsimportedreceiptnotes, out int importsuccessrows, out int importfailurerows, out DataSet dsreturnreceiptnotes, out String errmsg)
        {
            int v_importsuccessrows = 0, v_importfailurerows = 0;
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
                errmsg = "登陆失败，原因：" + u8Login.ShareString;
                //返回数据导入是否成功标志
                importsuccessrows = v_importsuccessrows;
                importfailurerows = v_importfailurerows;
                dsreturnreceiptnotes = dsimportedreceiptnotes;
                return false;
            }

            //第二步：构造环境上下文对象，传入login，并按需设置其它上下文参数
            U8EnvContext envContext = new U8EnvContext();
            envContext.U8Login = u8Login;


            //第三步：设置API地址标识(Url)：当前API：添加新单据的地址标识为：U8API/PuStoreIn/Add
            U8ApiAddress BestU8ApiAddress = new U8ApiAddress("U8API/PuStoreIn/Add");

            //第四步：构造APIBroker
            U8ApiBroker broker = new U8ApiBroker(BestU8ApiAddress, envContext);


            //第五步：API单据值及参数赋值

            //API单据值赋值
            #region 
            //设置BO对象(表头)行数，只能为一行
            BusinessObject DomHead = broker.GetBoParam("DomHead");
            DomHead.RowCount = 1;
            //给BO对象(表头)的字段赋值，值可以是真实类型，也可以是无类型字符串.以下代码示例只设置第一行值。各字段定义详见API服务接口定义
            /****************************** 以下是必输字段 ****************************/
            DomHead[0]["id"] = "1000000404"; //主关键字段，int类型
            DomHead[0]["ccode"] = "testimp0006"; //入库单号，string类型
            DomHead[0]["ddate"] = "2015-01-12"; //入库日期，DateTime类型
            //DomHead[0]["cbustype"] = "普通采购"; //业务类型，int类型
            DomHead[0]["cmaker"] = "demo"; //制单人，string类型
            DomHead[0]["iexchrate"] = "1.00"; //汇率，double类型
            DomHead[0]["cexch_name"] = "人民币"; //币种，string类型
            DomHead[0]["cvencode"] = "01002"; //供货单位编码，string类型
            DomHead[0]["cvouchtype"] = "01"; //单据类型，string类型
            DomHead[0]["cwhcode"] = "04"; //仓库编码，string类型
            //DomHead[0]["brdflag"] = "1"; //收发标志，int类型
            DomHead[0]["csource"] = "库存"; //单据来源，int类型
            //设置BO对象行数
            BusinessObject domBody = broker.GetBoParam("domBody");
            domBody.RowCount = 10;
            /****************************** 以下是必输字段 ****************************/
            domBody[0]["autoid"] = "1000001229"; //主关键字段，int类型
            domBody[0]["id"] = "1000000404"; //与收发记录主表关联项，int类型
            domBody[0]["cinvcode"] = "01019002082"; //存货编码，string类型
            domBody[0]["iquantity"] = "777.00"; //数量，double类型
            domBody[0]["editprop"] = "A"; //编辑属性：A表新增，M表修改，D表删除，string类型
            //domBody[0]["cinvouchtype"] = ""; //对应入库单类型，string类型
            //domBody[0]["cbmemo"] = ""; //备注，string类型
            //domBody[0]["irowno"] = ""; //行号，string类型
            #endregion
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
            if (!broker.Invoke())
            {
                //错误处理
                Exception apiEx = broker.GetException();
                if (apiEx != null)
                {
                    if (apiEx is MomSysException)
                    {
                        MomSysException sysEx = apiEx as MomSysException;
                        errmsg = "系统异常：" + sysEx.Message + "\n\r";
                        //todo:异常处理
                    }
                    else if (apiEx is MomBizException)
                    {
                        MomBizException bizEx = apiEx as MomBizException;
                        errmsg = "API异常：" + bizEx.Message + "\n\r";
                        //todo:异常处理
                    }
                    //异常原因
                    String exReason = broker.GetExceptionString();
                    if (exReason.Length != 0)
                    {
                        errmsg = "其他异常原因：" + exReason + "\n\r";
                    }
                }
                //结束本次调用，释放API资源
                broker.Release();
            }

            //第七步：获取返回结果

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

            //第八步 ： 结束本次调用，释放API资源
            broker.Release();
            if (result)
            {
                v_importsuccessrows = v_importsuccessrows + 1;

            }
            else
            {
                v_importfailurerows = v_importfailurerows + 1;
            }
            //返回数据导入是否成功标志
            importsuccessrows = v_importsuccessrows;
            importfailurerows = v_importfailurerows;
            dsreturnreceiptnotes = dsimportedreceiptnotes;
            return result;
        }

    }
}