using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using UFIDA.U8.Portal.Proxy.editors;
using UFIDA.U8.Portal.Common.Core;
using UFIDA.U8.Portal.Framework.MainFrames;
using UFSoft.U8.Framework.Login.UI;
using UFIDA.U8.Portal.Framework.Actions;
using UFIDA.U8.Portal.Proxy.Actions;
using System.IO;
using Microsoft.Office.Interop.Excel;
using UFIDA.U8.MomServiceCommon;
using UFIDA.U8.U8MOMAPIFramework;
using UFIDA.U8.U8APIFramework;
using UFIDA.U8.U8APIFramework.Meta;
using UFIDA.U8.U8APIFramework.Parameter;
using MSXML2;

namespace ImportReceiptNotes
{
    public partial class ImportReceiptNotes : UserControl
    {
        [DllImport("kernel32.dll")]
        public static extern IntPtr _lopen(string lpPathName, int iReadWrite);

        [DllImport("kernel32.dll")]
        public static extern bool CloseHandle(IntPtr hObject);

        public const int OF_READWRITE = 2;
        public const int OF_SHARE_DENY_NONE = 0x40;
        public readonly IntPtr HFILE_ERROR = new IntPtr(-1);
        public ImportReceiptNotes()
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
                MessageBox.Show("请先关闭数据模板导入EXCEL文件！");
                //启用导入按钮
                importdatabutton.Enabled = true;
                return;
            }
            CloseHandle(vHandle);

            //采购入库单导入
            string impstart, impend, v_errmsg;
            //采购入库单导入EXCEL中
            dtnpoidata = npoidata.ReadExcelToDatatble("ReceiptNotes", importdatafiletextBox.Text, 1, 1);
            dtnpoidata.TableName = "ReceiptNotes";
            dsexcel.Tables.Add(dtnpoidata);
            impstart = DateTime.Now.ToLocalTime().ToString();
            importdataresulttextBox.AppendText("数据导入执行开始:" + impstart + "\n");
            importdataresulttextBox.Refresh();
            //调用采购入库单导入功能
            bool v_importreceiptnoteflag = ReceiptNoteimport(Pubvar.u8userdata, dsexcel, out importsuccessrows, out importfailurerows, out dstoexcel, out v_errmsg);
            //导入结果回写EXCEL 
            npoidata.WriteDataTableToUpdateExcel(dstoexcel.Tables["ReceiptNotes"], "ReceiptNotes", importdatafiletextBox.Text);
            //执行结果回写memo text
            impend = DateTime.Now.ToLocalTime().ToString();
            importdataresulttextBox.AppendText("此次数据导入共计执行：" + (importsuccessrows + importfailurerows) + " 条 \n");
            importdataresulttextBox.AppendText("其中导入成功：" + importsuccessrows + " 条 \n");
            importdataresulttextBox.AppendText("其中导入失败：" + importfailurerows + " 条 \n");
            importdataresulttextBox.AppendText("如果导入有出错，具体原因请看导入数据模板中错误信息列，请纠正后再次执行导入！\n");
            if ((!v_importreceiptnoteflag) && (!string.IsNullOrEmpty(v_errmsg)))
            {
                importdataresulttextBox.AppendText("系统调用出错：" + v_errmsg + "\n");
            }
            importdataresulttextBox.AppendText("数据导入执行结束:" + impend + "  \n");
            importdataresulttextBox.Refresh();

            //为防止用户多次点击导入按钮，将按钮禁用
            importdatabutton.Enabled = true;

        }

        public bool ReceiptNoteimport(UFSoft.U8.Framework.LoginContext.UserData u8userdata, DataSet dsimportedreceiptnotes, out int importsuccessrows, out int importfailurerows, out DataSet dsreturnreceiptnotes, out string errmsg)
        {
            int v_importsuccessrows = 0, v_importfailurerows = 0;
            string v_errmsg = "", v_groupby = "", v_receiptnotnumber = "", v_filterestr = "", v_bustype = "", v_source = "", v_cinvcode = "", v_cinvname = "", v_receiptnotedate = "", strprocess = "", v_warehoursecode = "";
            bool v_exitflag = false;
            DataRow[] drgroupby, drorderlines;
            System.Data.DataTable dtsql = new System.Data.DataTable();
            DataSet orderhead = new DataSet(), orderlines = new DataSet(), dssql = new DataSet();
            U8Login.clsLogin u8Login;


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
                u8Login = new U8Login.clsLogin();
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
            //按照采购入库单导入数据模板，根据订单号、单据日期,仓库编码进行分组，基于采购管理业务参数必有订单选项限制，这里不考虑订单号为空的情况.
            System.Data.DataTable dtdistinct = dsimportedreceiptnotes.Tables["ReceiptNotes"].DefaultView.ToTable(true, new string[] { "订单号", "单据日期", "仓库编码" });
            //进度条初始化
            #region
            //初始化进度条
            importdataprogressBar.Minimum = 0;                              // 设置进度条最小值.
            importdataprogressBar.Value = 1;                                // 设置进度条初始值
            importdataprogressBar.Step = 1;                                 // 设置每次增加的步长
            importdataprogressBar.Maximum = dtdistinct.Rows.Count;          // 设置进度条最大值.
            //System.Drawing.Font font = new System.Drawing.Font("Times New Roman", (float)10, FontStyle.Regular);
            //PointF pt = new PointF(this.importdataprogressBar.Width / 2 - 17, this.importdataprogressBar.Height / 2 - 7);
            //Graphics g = this.importdataprogressBar.CreateGraphics();
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
                //按照分组标识数据获取导入数据模板该组数据值
                v_groupby = dtdistinct.Rows[i]["订单号"].ToString();
                v_receiptnotedate = dtdistinct.Rows[i]["单据日期"].ToString();
                v_warehoursecode = dtdistinct.Rows[i]["仓库编码"].ToString();
                if (!string.IsNullOrEmpty(v_groupby))
                {
                    v_filterestr = "订单号 = " + "'" + v_groupby + "' AND 单据日期 = '" + v_receiptnotedate + "' AND 仓库编码='" + v_warehoursecode + "'";
                }
                else
                {
                    v_filterestr = "(订单号 IS NULL  OR 订单号 ='" + "') AND 单据日期 = '" + v_receiptnotedate + "' AND 仓库编码='" + v_warehoursecode + "'";
                }
                drgroupby = dsimportedreceiptnotes.Tables["ReceiptNotes"].Select(v_filterestr);
                //执行导入模板数据初步校验 
                //1.订单号不能为空
                #region
                if (string.IsNullOrEmpty(v_groupby))
                {
                    v_importfailurerows = v_importfailurerows + 1;
                    //回写错误信息
                    for (int j = 0; j < drgroupby.Count(); j++)
                    {
                        drgroupby[j]["是否导入"] = "N";
                        drgroupby[j]["错误信息"] = "订单号不能为空，请检查!";
                    }
                    //复制已导入的数据到返回数据表中
                    for (int k = 0; k < drgroupby.Count(); k++)
                    {
                        dsreturnreceiptnotes.Tables["ReceiptNotes"].ImportRow(drgroupby[k]);
                    }

                    //执行进度条：PerformStep()函数
                    importdataprogressBar.PerformStep();
                    //strprocess = Math.Round((100 * (i + 1.0) / dtdistinct.Rows.Count), 2).ToString("#0.00 ") + "%";
                    //g.DrawString(strprocess, font, Brushes.Blue, pt);
                    v_exitflag = false;
                    continue;
                }
                #endregion
                //2.针对已导入数据即“单据号”列不为空或“是否导入”列不为空，或“错误信息"列不为空则直接复制到返回datatable中，不参加数据导入API调用
                #region
                if ((!string.IsNullOrEmpty(drgroupby[0]["单据号"].ToString())) || (!string.IsNullOrEmpty(drgroupby[0]["是否导入"].ToString())) || (!string.IsNullOrEmpty(drgroupby[0]["错误信息"].ToString())))
                {
                    //复制已导入的数据到返回数据表中
                    for (int k = 0; k < drgroupby.Count(); k++)
                    {
                        dsreturnreceiptnotes.Tables["ReceiptNotes"].ImportRow(drgroupby[k]);
                    }
                    //执行进度条：PerformStep()函数
                    importdataprogressBar.PerformStep();
                    //strprocess = Math.Round((100 * (i + 1.0) / dtdistinct.Rows.Count), 2).ToString("#0.00 ") + "%";
                    //g.DrawString(strprocess, font, Brushes.Blue, pt);
                    v_exitflag = false;
                    continue;
                }
                #endregion
                //3.有订单号则单据来源必须为"采购订单"或"委外订单"  //，业务类型为普通采购或委外加工。
                #region
                if (!string.IsNullOrEmpty(v_groupby))
                {
                    for (int j = 0; j < drgroupby.Count(); j++)
                    {
                        v_source = drgroupby[j]["单据来源"].ToString();
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
                            drgroupby[j]["错误信息"] = "单据来源错误，请检查!";
                        }
                        //复制已导入的数据到返回数据表中
                        for (int k = 0; k < drgroupby.Count(); k++)
                        {
                            dsreturnreceiptnotes.Tables["ReceiptNotes"].ImportRow(drgroupby[k]);
                        }
                        //执行进度条：PerformStep()函数
                        importdataprogressBar.PerformStep();
                        //strprocess = Math.Round((100 * (i + 1.0) / dtdistinct.Rows.Count), 2).ToString("#0.00 ") + "%";
                        //g.DrawString(strprocess, font, Brushes.Blue, pt);
                        v_exitflag = false;
                        continue;
                    }
                }

                #endregion
                //4.导入数据中item是否存在PO中
                #region
                if ((!string.IsNullOrEmpty(v_groupby)) && (drgroupby[0]["单据来源"].ToString() == "采购订单"))
                {
                    for (int j = 0; j < drgroupby.Count(); j++)
                    {
                        v_cinvname = drgroupby[j]["存货名称"].ToString();
                        sqlcmd.CommandText = "SELECT b.cInvCode FROM dbo.PO_Pomain as a inner join dbo.PO_Podetails as b on a.POID = b.POID  inner join dbo.Inventory as c  on b.cInvCode = c.cInvCode WHERE a.cPOID = '" + v_groupby + "' AND c.cInvName = '" + v_cinvname + "'";
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
                            drgroupby[j]["错误信息"] = "采购订单中不存在需要导入的存货编码，请检查!";
                        }
                        //复制已导入的数据到返回数据表中
                        for (int k = 0; k < drgroupby.Count(); k++)
                        {
                            dsreturnreceiptnotes.Tables["ReceiptNotes"].ImportRow(drgroupby[k]);
                        }
                        //执行进度条：PerformStep()函数
                        importdataprogressBar.PerformStep();
                        //strprocess = Math.Round((100 * (i + 1.0) / dtdistinct.Rows.Count), 2).ToString("#0.00 ") + "%";
                        //g.DrawString(strprocess, font, Brushes.Blue, pt);
                        v_exitflag = false;
                        continue;
                    }
                }

                if ((!string.IsNullOrEmpty(v_groupby)) && (drgroupby[0]["单据来源"].ToString() == "委外订单"))
                {
                    for (int j = 0; j < drgroupby.Count(); j++)
                    {
                        v_cinvname = drgroupby[j]["存货名称"].ToString();
                        sqlcmd.CommandText = "SELECT b.cInvCode FROM dbo.OM_MOMain as a inner join dbo.OM_MODetails as b on a.MOID = b.MOID inner join dbo.Inventory as c  on b.cInvCode = c.cInvCode WHERE a.cCode = '" + v_groupby + "'  AND c.cInvName = '" + v_cinvname + "'";
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
                        //执行进度条：PerformStep()函数
                        importdataprogressBar.PerformStep();
                        //strprocess = Math.Round((100 * (i + 1.0) / dtdistinct.Rows.Count), 2).ToString("#0.00 ") + "%";
                        //g.DrawString(strprocess, font, Brushes.Blue, pt);
                        v_exitflag = false;
                        continue;
                    }
                }
                #endregion
                //5.导入数据中订单号是否存在
                #region
                if ((!string.IsNullOrEmpty(v_groupby)) && (drgroupby[0]["单据来源"].ToString() == "采购订单"))
                {
                    sqlcmd.CommandText = "SELECT a.cPOID FROM dbo.PO_Pomain as a WHERE a.cPOID = '" + v_groupby + "'";
                    apdata.SelectCommand = sqlcmd;
                    orderhead.Reset();
                    apdata.Fill(orderhead);
                    if (orderhead.Tables[0].Rows.Count == 0)
                    {
                        v_exitflag = true;
                        break;
                    }

                    if (v_exitflag)
                    {
                        v_importfailurerows = v_importfailurerows + 1;
                        //回写错误信息
                        for (int j = 0; j < drgroupby.Count(); j++)
                        {
                            drgroupby[j]["是否导入"] = "N";
                            drgroupby[j]["错误信息"] = "采购订单中不存在需要导入的订单编号，请检查!";
                        }
                        //复制已导入的数据到返回数据表中
                        for (int k = 0; k < drgroupby.Count(); k++)
                        {
                            dsreturnreceiptnotes.Tables["ReceiptNotes"].ImportRow(drgroupby[k]);
                        }
                        //执行进度条：PerformStep()函数
                        importdataprogressBar.PerformStep();
                        //strprocess = Math.Round((100 * (i + 1.0) / dtdistinct.Rows.Count), 2).ToString("#0.00 ") + "%";
                        //g.DrawString(strprocess, font, Brushes.Blue, pt);
                        v_exitflag = false;
                        continue;
                    }
                }

                if ((!string.IsNullOrEmpty(v_groupby)) && (drgroupby[0]["单据来源"].ToString() == "委外订单"))
                {

                    sqlcmd.CommandText = "SELECT a.cCode FROM dbo.OM_MOMain as a  WHERE a.cCode = '" + v_groupby + "'";
                    apdata.SelectCommand = sqlcmd;
                    orderhead.Reset();
                    apdata.Fill(orderhead);
                    if (orderhead.Tables[0].Rows.Count == 0)
                    {
                        v_exitflag = true;
                        break;
                    }
                    if (v_exitflag)
                    {
                        v_importfailurerows = v_importfailurerows + 1;
                        //回写错误信息
                        for (int j = 0; j < drgroupby.Count(); j++)
                        {
                            drgroupby[j]["是否导入"] = "N";
                            drgroupby[j]["错误信息"] = "委外订单中不存在需要导入的订单编号，请检查!";
                        }
                        //复制已导入的数据到返回数据表中
                        for (int k = 0; k < drgroupby.Count(); k++)
                        {
                            dsreturnreceiptnotes.Tables["ReceiptNotes"].ImportRow(drgroupby[k]);
                        }
                        //执行进度条：PerformStep()函数
                        importdataprogressBar.PerformStep();
                        //strprocess = Math.Round((100 * (i + 1.0) / dtdistinct.Rows.Count), 2).ToString("#0.00 ") + "%";
                        //g.DrawString(strprocess, font, Brushes.Blue, pt);
                        v_exitflag = false;
                        continue;
                    }
                }
                #endregion
                //执行API单据赋值及参数赋值.
                #region
                if ((!string.IsNullOrEmpty(v_groupby)) && (drgroupby[0]["单据来源"].ToString() == "采购订单"))
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

                if ((!string.IsNullOrEmpty(v_groupby)) && (drgroupby[0]["单据来源"].ToString() == "委外订单"))
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


                if ((!string.IsNullOrEmpty(v_groupby)) && (drgroupby[0]["单据来源"].ToString() == "采购订单"))
                {
                    DomHead[0]["cvencode"] = orderhead.Tables[0].Rows[0]["cVenCode"].ToString();                   //供应商编号
                    DomHead[0]["cdepcode"] = orderhead.Tables[0].Rows[0]["cDepCode"].ToString();                   //部门编号
                    DomHead[0]["cbustype"] = orderhead.Tables[0].Rows[0]["cBusType"].ToString();                   //业务类型
                    DomHead[0]["csource"] = "采购订单";                                                           //单据来源                                               
                    DomHead[0]["iexchrate"] = orderhead.Tables[0].Rows[0]["nflat"].ToString();                      //汇率
                    DomHead[0]["cexch_name"] = orderhead.Tables[0].Rows[0]["cexch_name"].ToString();                //币种
                    DomHead[0]["ipurorderid"] = orderhead.Tables[0].Rows[0]["POID"].ToString();                     //采购订单ID
                    DomHead[0]["cmemo"] = orderhead.Tables[0].Rows[0]["cmemo"].ToString();                          //入库单备注信息

                }
                if ((!string.IsNullOrEmpty(v_groupby)) && (drgroupby[0]["单据来源"].ToString() == "委外订单"))
                {
                    DomHead[0]["cvencode"] = orderhead.Tables[0].Rows[0]["cVenCode"].ToString();
                    DomHead[0]["cdepcode"] = orderhead.Tables[0].Rows[0]["cDepCode"].ToString();
                    DomHead[0]["cbustype"] = orderhead.Tables[0].Rows[0]["cBusType"].ToString();
                    DomHead[0]["csource"] = "委外订单";
                    DomHead[0]["iexchrate"] = orderhead.Tables[0].Rows[0]["nflat"].ToString();
                    DomHead[0]["cexch_name"] = orderhead.Tables[0].Rows[0]["cexch_name"].ToString();
                    DomHead[0]["ipurorderid"] = orderhead.Tables[0].Rows[0]["MOID"].ToString();                     //委外订单ID
                    DomHead[0]["cmemo"] = orderhead.Tables[0].Rows[0]["cmemo"].ToString();                          //入库单备注信息
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
                    //通过存货名称获取存货编码
                    sqlcmd.CommandText = "select  TOP 1 cInvCode from dbo.Inventory where cInvName = '" + drgroupby[j]["存货名称"].ToString() + "'";
                    apdata.SelectCommand = sqlcmd;
                    dssql.Reset();
                    apdata.Fill(dssql);
                    dtsql.Reset();
                    dtsql = dssql.Tables[0];
                    domBody[j]["cinvcode"] = dtsql.Rows[0]["cInvCode"].ToString();              //存货编码
                    domBody[j]["iquantity"] = drgroupby[j]["入库数量"].ToString();              //入库数量
                    domBody[j]["editprop"] = "A";                                               //编辑属性：A表新增，M表修改，D表删除
                    domBody[j]["irowno"] = j + 1;                                               //行号

                    if ((!string.IsNullOrEmpty(v_groupby)) && (drgroupby[0]["单据来源"].ToString() == "采购订单"))
                    {
                        v_filterestr = " cInvCode = '" + domBody[j]["cinvcode"] + "'";
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
                        sqlcmd.CommandText = "SELECT bInvBatch,cSTComUnitCode FROM dbo.Inventory WHERE cInvCode = '" + domBody[j]["cinvcode"] + "'";
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

                    if ((!string.IsNullOrEmpty(v_groupby)) && (drgroupby[0]["单据来源"].ToString() == "委外订单"))
                    {
                        v_filterestr = " cInvCode = '" + domBody[j]["cinvcode"] + "'";
                        drorderlines = orderlines.Tables[0].Select(v_filterestr);
                        domBody[j]["itaxrate"] = 0.00;                                                      //税率 orderhead.Tables[0].Rows[0]["itaxrate"].ToString();        
                        domBody[j]["cpoid"] = v_groupby;                                                    //订单号，string类型
                        domBody[0]["iomodid"] = drorderlines[0]["MODetailsID"].ToString();                  //委外订单子表ID，int类型
                        domBody[j]["inquantity"] = drorderlines[0]["iQuantity"].ToString();                 //应收数量
                        domBody[0]["iprocesscost"] = drorderlines[0]["iUnitPrice"].ToString();              //加工费单价
                        domBody[0]["iprocessfee"] = Math.Round(Convert.ToDouble(drgroupby[j]["入库数量"]) * Convert.ToDouble(drorderlines[0]["iUnitPrice"]), 2).ToString();       //加工费

                        if (!string.IsNullOrEmpty(drgroupby[j]["入库件数"].ToString()))
                        {
                            domBody[0]["inum"] = drgroupby[j]["入库件数"].ToString();                                //入库件数
                            domBody[0]["iinvexchrate"] = Math.Round(Convert.ToDouble(drgroupby[j]["入库数量"]) / Convert.ToDouble(drgroupby[j]["入库件数"]), 4).ToString();     //换算率
                        }
                        domBody[0]["innum"] = "0.00";

                        //获取入库物料信息：是否批次控制
                        sqlcmd.CommandText = "SELECT bInvBatch FROM dbo.Inventory WHERE cInvCode = '" + domBody[j]["cinvcode"] + "'";
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
                //strprocess = Math.Round((100 * (i + 1.0) / dtdistinct.Rows.Count), 2).ToString("#0.00 ") + "%";
                //g.DrawString(strprocess, font, Brushes.Blue, pt);

            }
            #endregion
            //结束本次数据导入调用,返回数据导入是否成功标志,
            Marshal.FinalReleaseComObject(u8Login);
            importsuccessrows = v_importsuccessrows;
            importfailurerows = v_importfailurerows;
            dsreturnreceiptnotes = dsimportedreceiptnotes;
            errmsg = "";
            conn.Close();
            if (v_importfailurerows != 0)
                return false;
            else
                return true;

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
            INetUserControl BestImportReceiptNotesctl = new BestImportReceiptNotesControl();
            base.ShowEmbedControl(BestImportReceiptNotesctl, cMenuId, true);
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

    public class BestImportReceiptNotesControl : UFIDA.U8.Portal.Proxy.editors.INetUserControl
    {
        ImportReceiptNotes bestrnimportusercontrol = null;
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
            //从这个类里可以获取登陆信息、数据库连接信息等等
            Pubvar.u8userdata = new UFSoft.U8.Framework.LoginContext.UserData();
            Pubvar.u8userdata = login.GetLoginInfo();

            //初始化自定义用户控件对象
            bestrnimportusercontrol = new ImportReceiptNotes();

            return bestrnimportusercontrol;
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
            IActionDelegate nsd = new NetSampleDelegate();
            NetAction ac = new NetAction("about", nsd);
            NetAction[] aclist;
            aclist = new NetAction[1];
            ac.Text = "关于";
            aclist[0] = ac;
            return aclist;
        }

    }

    public static class Pubvar
    {
        public static string ConnString;

        public static string UserId;

        public static string userToken;

        public static UFSoft.U8.Framework.LoginContext.UserData u8userdata;
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
