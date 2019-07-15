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
            Application.Run(new MainForm());

        }
    }
}
