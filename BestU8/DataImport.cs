using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using U8GLVouchers;
using U8APILib;

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
                GLvouchersimport();
            }

            //采购入库单导入
            if (Pubvar.gdataimporttype == "采购入库单导入")
            {
                ReceiptNoteimport();
            }
        }

        private void GLvouchersimport()
        {
            BestU8GLVouchers v_importglvouchers = new BestU8GLVouchers();
            int v_importsuccessrows =0,v_importfailurerows =0;
            //根据总账导入EXCEL模板将数据导入到dataset 中
            DataSet v_vouchersfromexcel = new DataSet();
            DataSet v_returnvouchers = new DataSet();

            //调用总账导入功能
            bool  v_importglvouchersflag= v_importglvouchers.GLvouchersimport(Pubvar.gu8LoginUI.userToken, Pubvar.gu8userdata.ConnString, v_vouchersfromexcel, Pubvar.gu8userdata.UserId,out v_importsuccessrows,out v_importfailurerows,out v_returnvouchers);

        }

        private void ReceiptNoteimport()
        {
            U8APILibClass v_importreceiptnotes = new U8APILibClass();
            int v_importsuccessrows = 0, v_importfailurerows = 0;
            String v_errmsg;
            //根据采购入库单导入EXCEL模板将数据导入到dataset 中
            DataSet v_receiptnotesfromexcel = new DataSet();
            DataSet v_returnreceiptnotes = new DataSet();

            //调用采购入库单导入功能
            bool v_importreceiptnoteflag = v_importreceiptnotes.ReceiptNoteimport(Pubvar.gu8userdata, v_receiptnotesfromexcel, out v_importsuccessrows, out v_importfailurerows,out v_returnreceiptnotes, out v_errmsg);
            MessageBox.Show(v_importreceiptnoteflag.ToString());
        }

    }
}
