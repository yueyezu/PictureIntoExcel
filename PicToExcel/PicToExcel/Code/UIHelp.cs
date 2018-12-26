using System.Windows.Forms;

namespace PicToExcel.Code
{
    public class UiHelp
    {
        /// <summary>
        /// 显示正常提示框
        /// </summary>
        /// <param name="content"></param>
        /// <param name="title"></param>
        public static void ShowInfo(string content, string title = "")
        {
            if (string.IsNullOrEmpty(title))
            {
                title = "提示";
            }
            MessageBox.Show(content, title, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// 显示错误提示框
        /// </summary>
        /// <param name="content"></param>
        /// <param name="title"></param>
        public static void ShowError(string content, string title = "")
        {
            if (string.IsNullOrEmpty(title))
            {
                title = "错误";
            }
            MessageBox.Show(content, title, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        /// <summary>
        /// 显示提醒提示
        /// </summary>
        /// <param name="content"></param>
        /// <param name="title"></param>
        public static void ShowWarning(string content, string title = "")
        {
            if (string.IsNullOrEmpty(title))
            {
                title = "警告";
            }
            MessageBox.Show(content, title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        /// <summary>
        /// 显示询问提示
        /// </summary>
        /// <param name="content"></param>
        /// <param name="title"></param>
        public static bool ShowConfirm(string content, string title = "")
        {
            if (string.IsNullOrEmpty(title))
            {
                title = "提示";
            }
            DialogResult result = MessageBox.Show(content, title, MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk);
            return result == DialogResult.Yes;
        }

    }
}
