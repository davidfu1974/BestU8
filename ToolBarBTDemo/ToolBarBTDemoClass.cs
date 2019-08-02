using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using MSXML2;
using U8Login;
using UAPVoucherControl85;

namespace ToolBarBTDemo
{
    //定义COM组件接口声明
    [ComVisible(true)]
    [Guid("022B3940-275E-455C-AD04-44C433DEB54F")]
    //U8 工具条插件接口定义
    public interface IToolBarBTDemoSample
    {
        void Init(object objLogin, object objForm, object objVoucher, object msbar);
        void RunCommand(object objLogin, object objForm, object objVoucher, string sKey, object VarentValue, string other);
    }

    //定义COM组件接口实现
    [ComVisible(true)]
    [Guid("0AF8AAAB-4B67-46B7-A59B-1954D3544BAE")]
    [ProgId("ToolBarBTDemo.ToolBarBTDemoClass")]
    public class ToolBarBTDemoClass: IToolBarBTDemoSample
    {
        public void Init(object objLogin, object objForm, object objVoucher, object msbar)
        {
            //MessageBox.Show("Init");
        }

        public void RunCommand(object objLogin, object objForm, object objVoucher, string sKey, object VarentValue, string other)
        {
            //VarentValue为在UFMeta_XX表中预置的cVariant的值。
            MessageBox.Show("RunCommand01 - 按钮传入的值：" + Convert.ToString(VarentValue));

            //获取登陆信息
            clsLogin clsLogin;
            clsLogin = objLogin as clsLogin;
            MessageBox.Show("用户登陆信息值：" + Convert.ToString(clsLogin.cUserName) + "-" + Convert.ToString(clsLogin.userToken));
            try
            {
                //获取单据信息,将单据对象转换成可操作类型
                IXMLDOMDocument2 domhead, dombody;
                IXMLDOMNodeList headval, bodyval;
                string csocode = "", ddate = "", ccuscode = "", ccusname = "";
                string cinvcode = "", iquantity = "", iunitprice = "", cinvname = "";
                ctlVoucher voucher = (ctlVoucher)objVoucher;

                if (voucher != null)
                {

                    domhead = voucher.GetHeadDom();
                    dombody = voucher.GetLineDom();

                    if (domhead !=null)
                    {
                        headval = domhead.selectNodes("//rs:data/z:row");
                    }
                    else
                    {
                        MessageBox.Show("单据头信息为空！");
                        return;
                    }

                    if (dombody !=null)
                    {
                        bodyval = dombody.selectNodes("//rs:data/z:row");
                    }
                    else
                    {
                        MessageBox.Show("单据行信息为空！");
                        return;
                    }

                    //销售订单头信息
                    foreach (IXMLDOMElement item in headval)
                    {
                        csocode = item.attributes.getNamedItem("csocode").nodeValue.ToString();
                        ddate = item.attributes.getNamedItem("ddate").nodeValue.ToString();
                        ccuscode = item.attributes.getNamedItem("ccuscode").nodeValue.ToString();
                        ccusname = item.attributes.getNamedItem("ccusname").nodeValue.ToString();
                    }

                    //销售订单行信息
                    foreach (IXMLDOMElement item in bodyval)
                    {
                        cinvcode = item.attributes.getNamedItem("cinvcode").nodeValue.ToString();
                        iquantity = item.attributes.getNamedItem("iquantity").nodeValue.ToString();
                        iunitprice = item.attributes.getNamedItem("iunitprice").nodeValue.ToString();
                        cinvname = item.attributes.getNamedItem("cinvname").nodeValue.ToString();
                    }

                    MessageBox.Show("订单头信息：" + csocode + "-" + ddate + "-" + ccuscode + "-" + ccusname + "\r\n" + "订单行信息：行记录数为" + bodyval.length.ToString() + " ; " + cinvcode + "-" + iquantity + "-" + iunitprice + "-" + cinvname);
                }
                else
                {
                    MessageBox.Show("单据是空的!");
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }


            //string tmpMenuID = sKey.Replace("_CUSTDEF", "");
            //switch (tmpMenuID)
            //{
            //    case "EXT_SAMPLE":
            //        voucher.AddNewLineEvent += voucher_AddNewLineEvent; //增加新事件处理
            //        FrmVouchDetail tmpDetail = new FrmVouchDetail(this);
            //        tmpDetail.LoadVouchDetail(voucher.get_headerText("ID"));
            //        tmpDetail.Show(Form.FromChildHandle((IntPtr)voucher.hwnd)); //显示顶层窗体，非模式窗体
            //        break;
            //    default:
            //        break;
            //}
        }
    }
}
