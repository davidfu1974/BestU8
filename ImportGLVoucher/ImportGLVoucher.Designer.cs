namespace ImportGLVoucher
{
    partial class ImportGLVoucher
    {
        /// <summary> 
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary> 
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.importtypeGB = new System.Windows.Forms.GroupBox();
            this.importdataopenfilebutton = new System.Windows.Forms.Button();
            this.importdatabutton = new System.Windows.Forms.Button();
            this.importdatafiletextBox = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.importdataresulttextBox = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.importdataprogressBar = new System.Windows.Forms.ProgressBar();
            this.label2 = new System.Windows.Forms.Label();
            this.importdataopenFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.importtypeGB.SuspendLayout();
            this.SuspendLayout();
            // 
            // importtypeGB
            // 
            this.importtypeGB.Controls.Add(this.importdataopenfilebutton);
            this.importtypeGB.Controls.Add(this.importdatabutton);
            this.importtypeGB.Controls.Add(this.importdatafiletextBox);
            this.importtypeGB.Controls.Add(this.label1);
            this.importtypeGB.Location = new System.Drawing.Point(21, 30);
            this.importtypeGB.Name = "importtypeGB";
            this.importtypeGB.Size = new System.Drawing.Size(947, 95);
            this.importtypeGB.TabIndex = 13;
            this.importtypeGB.TabStop = false;
            this.importtypeGB.Text = "总账凭证数据导入";
            // 
            // importdataopenfilebutton
            // 
            this.importdataopenfilebutton.Location = new System.Drawing.Point(673, 38);
            this.importdataopenfilebutton.Name = "importdataopenfilebutton";
            this.importdataopenfilebutton.Size = new System.Drawing.Size(113, 33);
            this.importdataopenfilebutton.TabIndex = 2;
            this.importdataopenfilebutton.Text = "打开";
            this.importdataopenfilebutton.UseVisualStyleBackColor = true;
            this.importdataopenfilebutton.Click += new System.EventHandler(this.importdataopenfilebutton_Click);
            // 
            // importdatabutton
            // 
            this.importdatabutton.Location = new System.Drawing.Point(812, 38);
            this.importdatabutton.Name = "importdatabutton";
            this.importdatabutton.Size = new System.Drawing.Size(102, 33);
            this.importdatabutton.TabIndex = 13;
            this.importdatabutton.Text = "数据导入";
            this.importdatabutton.UseVisualStyleBackColor = true;
            this.importdatabutton.Click += new System.EventHandler(this.importdatabutton_Click);
            // 
            // importdatafiletextBox
            // 
            this.importdatafiletextBox.Location = new System.Drawing.Point(98, 45);
            this.importdatafiletextBox.Name = "importdatafiletextBox";
            this.importdatafiletextBox.Size = new System.Drawing.Size(549, 21);
            this.importdatafiletextBox.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(15, 45);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(77, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "凭证模板数据";
            // 
            // importdataresulttextBox
            // 
            this.importdataresulttextBox.Location = new System.Drawing.Point(119, 212);
            this.importdataresulttextBox.Multiline = true;
            this.importdataresulttextBox.Name = "importdataresulttextBox";
            this.importdataresulttextBox.Size = new System.Drawing.Size(817, 315);
            this.importdataresulttextBox.TabIndex = 17;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(60, 215);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 12);
            this.label3.TabIndex = 16;
            this.label3.Text = "导入结果";
            // 
            // importdataprogressBar
            // 
            this.importdataprogressBar.Location = new System.Drawing.Point(119, 154);
            this.importdataprogressBar.Name = "importdataprogressBar";
            this.importdataprogressBar.Size = new System.Drawing.Size(817, 23);
            this.importdataprogressBar.TabIndex = 15;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(60, 154);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(53, 12);
            this.label2.TabIndex = 14;
            this.label2.Text = "导入进度";
            // 
            // ImportGLVoucher
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.importtypeGB);
            this.Controls.Add(this.importdataresulttextBox);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.importdataprogressBar);
            this.Controls.Add(this.label2);
            this.Name = "ImportGLVoucher";
            this.Size = new System.Drawing.Size(1008, 574);
            this.importtypeGB.ResumeLayout(false);
            this.importtypeGB.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox importtypeGB;
        private System.Windows.Forms.Button importdataopenfilebutton;
        private System.Windows.Forms.Button importdatabutton;
        private System.Windows.Forms.TextBox importdatafiletextBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox importdataresulttextBox;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ProgressBar importdataprogressBar;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.OpenFileDialog importdataopenFileDialog;
    }
}
