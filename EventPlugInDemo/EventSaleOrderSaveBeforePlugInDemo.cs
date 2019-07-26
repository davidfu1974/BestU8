using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using UFIDA.U8.MomServiceCommon;
using UFIDA.U8.U8MOMAPIFramework;
using UFIDA.U8.U8APIFramework;
using UFIDA.U8.U8APIFramework.Meta;
using UFIDA.U8.U8APIFramework.Parameter;
using MSXML2;
using U8Login;
using System.Windows.Forms;
using System.IO;

// 该DEMO主要通过销售订单保存前触发用户自定义开发的业务逻辑，此处显示销售订单头及行一些基本信息。
namespace EventPlugInDemo
{
    public class EventSaleOrderSaveBeforePlugInDemo
    {
        public EventSaleOrderSaveBeforePlugInDemo()
        {
        }

        public bool Save_before(ref IXMLDOMDocument2 domhead, ref IXMLDOMDocument2 dombody, ref string errMsg)
        {
            MomCallContext currentMomCallContext = MomCallContextCache.Instance.CurrentMomCallContext;
            clsLogin clsLogin = (clsLogin)currentMomCallContext.U8Login ;
            IXMLDOMNodeList headval = domhead.selectNodes("//rs:data/z:row");
            IXMLDOMNodeList bodyval = dombody.selectNodes("//rs:data/z:row");

            string csocode = "", ddate = "", ccuscode = "", ccusname = "";
            string cinvcode = "", iquantity = "", iunitprice = "", cinvname = "";

            wl.WriteLogs("-------head start ------" + domhead.xml + "-----head end ------------");

            wl.WriteLogs("-------body start ------" + dombody.xml + "-----body end ------------");

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

            MessageBox.Show("订单头信息：" + csocode + "-" + ddate + "-" + ccuscode + "-" + ccusname + "\r\n" + "订单行信息：行记录数为" + bodyval.length.ToString() +" ; "+ cinvcode + "-" + iquantity + "-" + iunitprice + "-" + cinvname);

            return true;
        }

        public static WriteLog wl = new WriteLog();
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
