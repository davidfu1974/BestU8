using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

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
    }
}
