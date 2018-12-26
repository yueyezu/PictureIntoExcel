using System;
using System.IO;
using System.Text;

namespace PicToExcel.Code
{
    public class LogHelper
    {
        private static string logPath = "";
        private string logName = "";

        public LogHelper(string logName)
        {
            this.logName = logName;
            //if (string.IsNullOrEmpty(logPath))
            //{

            //}
            //logPath = ConfigurationManager.AppSettings.Get("logPath");
            if (string.IsNullOrEmpty(logPath))
            {
                logPath = AppDomain.CurrentDomain.BaseDirectory + "\\log\\";
            }
        }

        private void Log(string logType, string msg)
        {
            if (!Directory.Exists(logPath))
            {
                Directory.CreateDirectory(logPath);
            }
            string filePath = logPath + DateTime.Now.ToString("yyMMdd") + ".log";
            try
            {
                StreamWriter strmW = new StreamWriter(filePath, true, Encoding.UTF8);
                strmW.WriteLine(DateTime.Now.ToString("[hh:mm:ss] [") + logType + "] " + logName + "--" + msg);
                strmW.Close();
            }
            catch
            {
            }
        }

        /// <summary>
        /// 输出对象到文件
        /// </summary>
        /// <param name="obj"></param>
        public void Log(object obj)
        {
            Log("Info", obj.ToString());
        }

        /// <summary>
        /// 输出警告信息
        /// </summary>
        /// <param name="msg">警告信息</param>
        public void LogWarningMsg(string msg)
        {
            Log("Wornning", msg);
        }

        /// <summary>
        /// 输出警告信息
        /// </summary>
        /// <param name="msg">错误信息</param>
        public void LogErrorMsg(string msg)
        {
            Log("Error", msg);
        }

    }
}
