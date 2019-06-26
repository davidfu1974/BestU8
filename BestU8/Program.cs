using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Runtime.InteropServices;


namespace BestU8
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new MainForm());

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

            //构建u8Login并执行登陆
            U8Login.clsLogin u8Login = new U8Login.clsLogin();
            String sSubId = u8userdata.cSubID;              // "AS";
            String sAccID = u8userdata.AccID;               // "(default)@999"
            String sYear = u8userdata.iYear;                 //"2014";
            String sUserID = u8userdata.UserId;             //"demo";
            String sPassword = u8userdata.Password;         // "";
            String sDate = u8userdata.operDate;             //"2014-12-11";
            String sServer = u8userdata.AppServer;          // "UF8125";
            String sSerial = "";                            //u8userdata.AppServerSerial;
            if (!u8Login.Login(ref sSubId, ref sAccID, ref sYear, ref sUserID, ref sPassword, ref sDate, ref sServer, ref sSerial))
            {
                MessageBox.Show("登陆失败，原因：" + u8Login.ShareString);
                Marshal.FinalReleaseComObject(u8Login);
                return;
            }
            else  // 显示操作主界面并保存登陆信息
            {
                Pubvar.gu8LoginUI = u8LoginUI;
                Pubvar.gu8Login = u8Login;
                Pubvar.gu8userdata = u8userdata;
                Pubvar.gdataimporttype = "";
                Application.Run(new MainForm());

            }
            
        }
    }
}
