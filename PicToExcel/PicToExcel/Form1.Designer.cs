namespace PicToExcel
{
    partial class Form1
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

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.txtPicPath = new System.Windows.Forms.TextBox();
            this.lblPicPath = new System.Windows.Forms.Label();
            this.btnPicPath = new System.Windows.Forms.Button();
            this.btnConvert = new System.Windows.Forms.Button();
            this.btnClear = new System.Windows.Forms.Button();
            this.txtXlsPath = new System.Windows.Forms.TextBox();
            this.lblXlsPath = new System.Windows.Forms.Label();
            this.btnXlsPath = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.txtXlsName = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtSpool = new System.Windows.Forms.NumericUpDown();
            ((System.ComponentModel.ISupportInitialize)(this.txtSpool)).BeginInit();
            this.SuspendLayout();
            // 
            // txtPicPath
            // 
            this.txtPicPath.Location = new System.Drawing.Point(85, 25);
            this.txtPicPath.Name = "txtPicPath";
            this.txtPicPath.ReadOnly = true;
            this.txtPicPath.Size = new System.Drawing.Size(270, 21);
            this.txtPicPath.TabIndex = 0;
            // 
            // lblPicPath
            // 
            this.lblPicPath.AutoSize = true;
            this.lblPicPath.Location = new System.Drawing.Point(14, 29);
            this.lblPicPath.Name = "lblPicPath";
            this.lblPicPath.Size = new System.Drawing.Size(65, 12);
            this.lblPicPath.TabIndex = 1;
            this.lblPicPath.Text = "图片路径：";
            // 
            // btnPicPath
            // 
            this.btnPicPath.Location = new System.Drawing.Point(369, 24);
            this.btnPicPath.Name = "btnPicPath";
            this.btnPicPath.Size = new System.Drawing.Size(75, 23);
            this.btnPicPath.TabIndex = 2;
            this.btnPicPath.Text = "选择图片";
            this.btnPicPath.UseVisualStyleBackColor = true;
            this.btnPicPath.Click += new System.EventHandler(this.btnPicPath_Click);
            // 
            // btnConvert
            // 
            this.btnConvert.Location = new System.Drawing.Point(30, 209);
            this.btnConvert.Name = "btnConvert";
            this.btnConvert.Size = new System.Drawing.Size(75, 23);
            this.btnConvert.TabIndex = 2;
            this.btnConvert.Text = "NPOI转化";
            this.btnConvert.UseVisualStyleBackColor = true;
            this.btnConvert.Click += new System.EventHandler(this.btnConvert_Click);
            // 
            // btnClear
            // 
            this.btnClear.Location = new System.Drawing.Point(353, 209);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(75, 23);
            this.btnClear.TabIndex = 2;
            this.btnClear.Text = "清空";
            this.btnClear.UseVisualStyleBackColor = true;
            this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
            // 
            // txtXlsPath
            // 
            this.txtXlsPath.Location = new System.Drawing.Point(85, 86);
            this.txtXlsPath.Name = "txtXlsPath";
            this.txtXlsPath.ReadOnly = true;
            this.txtXlsPath.Size = new System.Drawing.Size(270, 21);
            this.txtXlsPath.TabIndex = 0;
            // 
            // lblXlsPath
            // 
            this.lblXlsPath.AutoSize = true;
            this.lblXlsPath.Location = new System.Drawing.Point(14, 90);
            this.lblXlsPath.Name = "lblXlsPath";
            this.lblXlsPath.Size = new System.Drawing.Size(65, 12);
            this.lblXlsPath.TabIndex = 1;
            this.lblXlsPath.Text = "结果目录：";
            // 
            // btnXlsPath
            // 
            this.btnXlsPath.Location = new System.Drawing.Point(369, 85);
            this.btnXlsPath.Name = "btnXlsPath";
            this.btnXlsPath.Size = new System.Drawing.Size(75, 23);
            this.btnXlsPath.TabIndex = 2;
            this.btnXlsPath.Text = "结果路径";
            this.btnXlsPath.UseVisualStyleBackColor = true;
            this.btnXlsPath.Click += new System.EventHandler(this.btnXlsPath_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(28, 142);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 12);
            this.label1.TabIndex = 1;
            this.label1.Text = "颜色精度：";
            // 
            // txtXlsName
            // 
            this.txtXlsName.Location = new System.Drawing.Point(291, 138);
            this.txtXlsName.Name = "txtXlsName";
            this.txtXlsName.Size = new System.Drawing.Size(137, 21);
            this.txtXlsName.TabIndex = 0;
            this.txtXlsName.Text = "转化文件";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(220, 142);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(65, 12);
            this.label2.TabIndex = 1;
            this.label2.Text = "文件名称：";
            // 
            // txtSpool
            // 
            this.txtSpool.Location = new System.Drawing.Point(99, 138);
            this.txtSpool.Name = "txtSpool";
            this.txtSpool.Size = new System.Drawing.Size(70, 21);
            this.txtSpool.TabIndex = 3;
            this.txtSpool.Value = new decimal(new int[] {
            10,
            0,
            0,
            0});
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(470, 262);
            this.Controls.Add(this.txtSpool);
            this.Controls.Add(this.btnClear);
            this.Controls.Add(this.btnConvert);
            this.Controls.Add(this.btnXlsPath);
            this.Controls.Add(this.btnPicPath);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lblXlsPath);
            this.Controls.Add(this.lblPicPath);
            this.Controls.Add(this.txtXlsName);
            this.Controls.Add(this.txtXlsPath);
            this.Controls.Add(this.txtPicPath);
            this.Name = "Form1";
            this.Text = "PicToXls";
            ((System.ComponentModel.ISupportInitialize)(this.txtSpool)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtPicPath;
        private System.Windows.Forms.Label lblPicPath;
        private System.Windows.Forms.Button btnPicPath;
        private System.Windows.Forms.Button btnConvert;
        private System.Windows.Forms.Button btnClear;
        private System.Windows.Forms.TextBox txtXlsPath;
        private System.Windows.Forms.Label lblXlsPath;
        private System.Windows.Forms.Button btnXlsPath;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtXlsName;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.NumericUpDown txtSpool;
    }
}

