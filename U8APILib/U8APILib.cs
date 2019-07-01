using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Runtime.InteropServices;
using UFIDA.U8.MomServiceCommon;
using UFIDA.U8.U8MOMAPIFramework;
using UFIDA.U8.U8APIFramework;
using UFIDA.U8.U8APIFramework.Meta;
using UFIDA.U8.U8APIFramework.Parameter;
using MSXML2;


namespace U8APILib
{
    public class U8APILibClass
    {
        public U8APILibClass()
        {
        }
        //采购入库单导入
        public bool ReceiptNoteimport(Be UFSoft.U8.Framework.LoginContext.UserData u8userdata, DataSet dsimportedreceiptnotes, out int importsuccessrows, out int importfailurerows, out DataSet dsreturnreceiptnotes, out String errmsg)
        {
            int v_importsuccessrows = 0, v_importfailurerows = 0;
            //第一步：构造u8login对象并登陆(引用U8API类库中的Interop.U8Login.dll),如果当前环境中有login对象则可以省去第一步
            U8Login.clsLogin u8Login = new U8Login.clsLogin();
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
            //MomCallContext tt = new MomCallContext();
            //tt.EnvContext = u8Login;
            //tt.

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
                        errmsg =  "其他异常原因：" + exReason + "\n\r";
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
            System.String v_vouchid = (System.String)broker.GetResult("VouchId") ;

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
