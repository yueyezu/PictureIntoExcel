using System;
using System.IO;
using System.Windows.Forms;
using PicToExcel.Code;
using PicToExcel.Convert;

namespace PicToExcel
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 选择图片文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnPicPath_Click(object sender, EventArgs e)
        {
            FileDialog files = new OpenFileDialog();
            files.Filter = "图片文件(*.jpg;*.png;*.bmp)|*.jpg;*.png;*.bmp|所有文件|*.*";

            if (files.ShowDialog() == DialogResult.OK)
            {
                txtPicPath.Text = files.FileName;
            }
        }
        /// <summary>
        /// 结果路径选择
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnXlsPath_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowser = new FolderBrowserDialog();
            if (folderBrowser.ShowDialog() == DialogResult.OK)
            {
                txtXlsPath.Text = folderBrowser.SelectedPath;
            }
        }

        /// <summary>
        /// 清空选择的内容
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnClear_Click(object sender, EventArgs e)
        {
            txtXlsPath.Text = "";
            txtPicPath.Text = "";
        }

        /// <summary>
        /// 开始转化
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnConvert_Click(object sender, EventArgs e)
        {
            string picPath = txtPicPath.Text;
            string xlsPath = txtXlsPath.Text;
            string xlsName = txtXlsName.Text;
            int spool = System.Convert.ToInt32(txtSpool.Value);

            if (string.IsNullOrEmpty(picPath))
            {
                UiHelp.ShowWarning("请选择要转化的图片");
                return;
            }
            if (string.IsNullOrEmpty(xlsPath) || string.IsNullOrEmpty(xlsName))
            {
                UiHelp.ShowWarning("请设置转化后文件的路径和名称。");
                return;
            }

            PicIntoExcelByNpoi intoExcelByNpoi = new PicIntoExcelByNpoi(picPath, xlsPath, spool);

            try
            {
                string resultPath = intoExcelByNpoi.StartConvert(xlsName + ".xls");
                UiHelp.ShowInfo("转化完成。");

                resultPath = resultPath.Substring(0, resultPath.LastIndexOf("\\"));
                System.Diagnostics.Process.Start("Explorer.exe", resultPath);
            }
            catch (Exception ex)
            {
                UiHelp.ShowError("转化失败！" + ex.Message);

            }
        }
    }
}
