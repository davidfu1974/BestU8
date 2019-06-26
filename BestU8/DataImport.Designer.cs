namespace BestU8
{
    partial class DataImport
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.importtypeGB = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.importdataopenFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.importdatafiletextBox = new System.Windows.Forms.TextBox();
            this.importdataopenfilebutton = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.importdataprogressBar = new System.Windows.Forms.ProgressBar();
            this.label3 = new System.Windows.Forms.Label();
            this.importdataresulttextBox = new System.Windows.Forms.TextBox();
            this.importdatabutton = new System.Windows.Forms.Button();
            this.closebutton = new System.Windows.Forms.Button();
            this.importtypeGB.SuspendLayout();
            this.SuspendLayout();
            // 
            // importtypeGB
            // 
            this.importtypeGB.Controls.Add(this.importdataopenfilebutton);
            this.importtypeGB.Controls.Add(this.importdatafiletextBox);
            this.importtypeGB.Controls.Add(this.label1);
            this.importtypeGB.Location = new System.Drawing.Point(36, 23);
            this.importtypeGB.Name = "importtypeGB";
            this.importtypeGB.Size = new System.Drawing.Size(771, 86);
            this.importtypeGB.TabIndex = 0;
            this.importtypeGB.TabStop = false;
            this.importtypeGB.Text = "XXXXX数据导入";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(39, 35);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(77, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "导入模板数据";
            // 
            // importdataopenFileDialog
            // 
            this.importdataopenFileDialog.FileName = "openFileDialog";
            // 
            // importdatafiletextBox
            // 
            this.importdatafiletextBox.Location = new System.Drawing.Point(122, 32);
            this.importdatafiletextBox.Name = "importdatafiletextBox";
            this.importdatafiletextBox.Size = new System.Drawing.Size(523, 21);
            this.importdatafiletextBox.TabIndex = 1;
            // 
            // importdataopenfilebutton
            // 
            this.importdataopenfilebutton.Location = new System.Drawing.Point(668, 29);
            this.importdataopenfilebutton.Name = "importdataopenfilebutton";
            this.importdataopenfilebutton.Size = new System.Drawing.Size(75, 23);
            this.importdataopenfilebutton.TabIndex = 2;
            this.importdataopenfilebutton.Text = "打开";
            this.importdataopenfilebutton.UseVisualStyleBackColor = true;
            this.importdataopenfilebutton.Click += new System.EventHandler(this.importdataopenfilebutton_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(99, 138);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(53, 12);
            this.label2.TabIndex = 1;
            this.label2.Text = "导入进度";
            // 
            // importdataprogressBar
            // 
            this.importdataprogressBar.Location = new System.Drawing.Point(158, 133);
            this.importdataprogressBar.Name = "importdataprogressBar";
            this.importdataprogressBar.Size = new System.Drawing.Size(621, 23);
            this.importdataprogressBar.TabIndex = 2;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(99, 204);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 12);
            this.label3.TabIndex = 3;
            this.label3.Text = "导入结果";
            // 
            // importdataresulttextBox
            // 
            this.importdataresulttextBox.Location = new System.Drawing.Point(158, 201);
            this.importdataresulttextBox.Multiline = true;
            this.importdataresulttextBox.Name = "importdataresulttextBox";
            this.importdataresulttextBox.Size = new System.Drawing.Size(621, 137);
            this.importdataresulttextBox.TabIndex = 4;
            // 
            // importdatabutton
            // 
            this.importdatabutton.Location = new System.Drawing.Point(596, 391);
            this.importdatabutton.Name = "importdatabutton";
            this.importdatabutton.Size = new System.Drawing.Size(75, 23);
            this.importdatabutton.TabIndex = 5;
            this.importdatabutton.Text = "数据导入";
            this.importdatabutton.UseVisualStyleBackColor = true;
            // 
            // closebutton
            // 
            this.closebutton.Location = new System.Drawing.Point(704, 391);
            this.closebutton.Name = "closebutton";
            this.closebutton.Size = new System.Drawing.Size(75, 23);
            this.closebutton.TabIndex = 7;
            this.closebutton.Text = "关闭";
            this.closebutton.UseVisualStyleBackColor = true;
            this.closebutton.Click += new System.EventHandler(this.closebutton_Click);
            // 
            // DataImport
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(827, 450);
            this.Controls.Add(this.closebutton);
            this.Controls.Add(this.importdatabutton);
            this.Controls.Add(this.importdataresulttextBox);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.importdataprogressBar);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.importtypeGB);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "DataImport";
            this.Text = "DataImport";
            this.importtypeGB.ResumeLayout(false);
            this.importtypeGB.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox importtypeGB;
        private System.Windows.Forms.Button importdataopenfilebutton;
        private System.Windows.Forms.TextBox importdatafiletextBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.OpenFileDialog importdataopenFileDialog;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ProgressBar importdataprogressBar;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox importdataresulttextBox;
        private System.Windows.Forms.Button importdatabutton;
        private System.Windows.Forms.Button closebutton;
    }
}